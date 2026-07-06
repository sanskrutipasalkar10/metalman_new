from PIL import Image as PILImage


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
