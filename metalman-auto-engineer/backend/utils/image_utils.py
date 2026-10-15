from PIL import Image as PILImage
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

EMU_PER_PIXEL = 9525          # Office's fixed px->EMU conversion at 96 DPI
DEFAULT_ROW_HEIGHT_PT = 15.0  # Excel's default row height when unset
DEFAULT_COL_WIDTH_CHARS = 8.43  # Excel's default column width when unset


def fit_image_dimensions(image_path: str, max_width: int, max_height: int) -> tuple[int, int]:
    """
    Returns (width, height) in pixels that fit the image within the given
    max box while preserving its original aspect ratio ("contain" scaling).

    Forcing a fixed width/height regardless of the source aspect ratio
    stretches/squashes the image non-uniformly - this keeps the crop's true
    proportions intact when it's inserted into an Excel cell.
    """
    try:
        with PILImage.open(image_path) as img:
            src_w, src_h = img.size
    except Exception:
        return max_width, max_height

    if src_w <= 0 or src_h <= 0:
        return max_width, max_height

    scale = min(max_width / src_w, max_height / src_h)
    return max(1, round(src_w * scale)), max(1, round(src_h * scale))


def trim_whitespace(img: PILImage.Image, padding: int = 10, bg_threshold: int = 245) -> PILImage.Image:
    """
    Crops a PIL image down to the bounding box of its non-background content
    plus a small padding, discarding blank margin.

    A manually-dragged crop selection is rarely pixel-tight around the
    target - a loosely-drawn box leaves blank space on one or more sides
    baked into the image itself. Centering *that* image in a cell still
    looks off-center, since centering the outer bounds doesn't center the
    ink within them - trimming first is what actually fixes it.
    """
    gray = img.convert("L")
    content_mask = gray.point(lambda p: 255 if p < bg_threshold else 0)
    bbox = content_mask.getbbox()
    if not bbox:
        return img
    left, top, right, bottom = bbox
    left = max(0, left - padding)
    top = max(0, top - padding)
    right = min(img.width, right + padding)
    bottom = min(img.height, bottom + padding)
    return img.crop((left, top, right, bottom))


def _col_width_to_px(width_chars: float) -> float:
    # Standard Excel approximation for the default Calibri 11 font (~7px/char + 5px padding)
    return width_chars * 7 + 5


def _row_height_to_px(height_pt: float) -> float:
    return height_pt * 96 / 72


def place_image_centered(ws, image_path: str, cell_ref: str, max_width: int = None, max_height: int = None) -> ExcelImage:
    """
    Insert `image_path` into `cell_ref`, scaled to fit (aspect preserved) and
    truly centered within the cell's actual pixel bounds.

    Plain `ws.add_image(img, cell_ref)` anchors the image flush to the cell's
    top-left corner at whatever pixel width/height was set on it - it never
    centers or accounts for the cell's real rendered size. Whenever the fitted
    image height didn't exactly equal the row's actual pixel height (which is
    the common case), the image visibly drifted upward into the row above,
    which is what the "images placed above the cell" reports were seeing.

    This instead computes the cell's real size from its row height / column
    width, fits the image within a small margin of that box, and builds an
    explicit OneCellAnchor with the offsets needed to center it - so the
    result is correct regardless of what row height / column width the
    template ends up with, with no manual nudging required.

    If `cell_ref` is the top-left of a merged range, the full merged span's
    width/height is used instead of just the anchor cell's own - otherwise a
    wide merge (e.g. a full-width reference-photo row) would be sized as if
    it were a single narrow column.
    """
    col_letter, row_num = coordinate_from_string(cell_ref)
    col_idx0 = column_index_from_string(col_letter) - 1
    row_idx0 = row_num - 1

    span_cols, span_rows = 1, 1
    for m in ws.merged_cells.ranges:
        if m.min_row == row_num and m.min_col == col_idx0 + 1:
            span_cols = m.max_col - m.min_col + 1
            span_rows = m.max_row - m.min_row + 1
            break

    cell_w_px = 0.0
    for c in range(col_idx0 + 1, col_idx0 + 1 + span_cols):
        col_width_chars = ws.column_dimensions[get_column_letter(c)].width or DEFAULT_COL_WIDTH_CHARS
        cell_w_px += _col_width_to_px(col_width_chars)

    cell_h_px = 0.0
    for r in range(row_num, row_num + span_rows):
        row_height_pt = ws.row_dimensions[r].height or DEFAULT_ROW_HEIGHT_PT
        cell_h_px += _row_height_to_px(row_height_pt)

    # Fit within ~94% of the cell so the image doesn't touch the borders
    box_w = min(max_width, cell_w_px * 0.94) if max_width else cell_w_px * 0.94
    box_h = min(max_height, cell_h_px * 0.94) if max_height else cell_h_px * 0.94

    img_w, img_h = fit_image_dimensions(image_path, int(box_w), int(box_h))

    img = ExcelImage(image_path)
    img.width, img.height = img_w, img_h

    x_off_px = max(0.0, (cell_w_px - img_w) / 2)
    y_off_px = max(0.0, (cell_h_px - img_h) / 2)

    img.anchor = OneCellAnchor(
        _from=AnchorMarker(
            col=col_idx0, colOff=round(x_off_px * EMU_PER_PIXEL),
            row=row_idx0, rowOff=round(y_off_px * EMU_PER_PIXEL),
        ),
        ext=XDRPositiveSize2D(cx=round(img_w * EMU_PER_PIXEL), cy=round(img_h * EMU_PER_PIXEL)),
    )
    ws.add_image(img)
    return img
