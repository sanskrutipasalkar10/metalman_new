from fastapi import APIRouter, UploadFile, File, BackgroundTasks, Form, HTTPException
import os
import shutil
import uuid
import json
import re
import glob
import asyncio
import pandas as pd
from pydantic import BaseModel
from typing import List
from services.pfd_renderer import render_sub_assy_index, render_sub_assy_sheet, render_sheetmetal_sheet, render_bop_sheet
from services.cad_service import export_step_to_stl
from services.bom_generator import generate_bom_excel, extract_isometric_view, isolate_part_and_remove_text
from services.tooling_generator import generate_tooling_list
from services.fitment_generator import generate_fitment_check_sheet
from services.pfmea_generator import generate_pfmea_excel
from services.control_plan_generator import generate_control_plan_excel
from services.fixture_plan_generator import generate_fixture_plan
from services.correction_service import apply_bulk_correction
from utils.part_info import get_canonical_part_info
from xlsx2html import xlsx2html
from fastapi.responses import HTMLResponse, FileResponse
import io
try:
    import pythoncom
    import win32com.client
    HAS_WIN32 = True
except ImportError:
    pythoncom = None
    HAS_WIN32 = False

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from utils.image_utils import place_image_centered, trim_whitespace
from PIL import Image
import base64
from io import BytesIO
router = APIRouter()

@router.get("/demo-folders")
async def get_demo_folders():
    # Path to the internal stitch_suite folder
    demo_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "stitch_suite"))
    folders = []
    if os.path.exists(demo_path):
        for item in os.listdir(demo_path):
            if os.path.isdir(os.path.join(demo_path, item)) and not item.startswith("."):
                # Filter out landing and login pages as requested
                if "landing" in item or "login" in item:
                    continue
                # Check if content_only.html or code.html exists
                if os.path.exists(os.path.join(demo_path, item, "content_only.html")) or \
                   os.path.exists(os.path.join(demo_path, item, "code.html")):
                    folders.append(item)
    return sorted(folders)

class CorrectionItem(BaseModel):
    document: str
    column: str
    cell_no: str
    replacement_content: str

class BulkCorrectionRequest(BaseModel):
    task_id: str
    corrections: List[CorrectionItem]


DATA_DIR = "/data" if os.path.exists("/data") else os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
UPLOAD_DIR = os.path.abspath(os.path.join(DATA_DIR, "uploads"))
OUTPUT_DIR = os.path.abspath(os.path.join(DATA_DIR, "outputs"))
TASKS_FILE = os.path.join(DATA_DIR, "tasks_db.json")
WELDING_EXCEL_FILENAME = os.path.join(OUTPUT_DIR, "welding_crops.xlsx")
WELDING_TEMPLATE_PATH  = os.path.join("assets", "Welding Check Sheet-92187158.xlsx")
WELDING_DATA_START_ROW = 10  # first Sl. No row below the "Sr.No." sub-header
WELDING_DATA_END_ROW   = 36  # last data row before the "REMARK" footer (row 37)
PFD_TEMPLATE_PATH     = os.path.join("assets", "pfd template.xlsx")
BOM_TEMPLATE_PATH     = os.path.join("assets", "BLANK_BOM_TEMPLATE.xlsx")
TOOLING_TEMPLATE_PATH = os.path.join("assets", "tooling list temp.xlsx")
FITMENT_TEMPLATE_PATH = os.path.join("assets", "fitment_checksheet_template.xlsx")
PFMEA_TEMPLATE_PATH   = os.path.join("assets", "PFMEA_TEMPLATE.xlsx")
PFMEA_DICT_DIR             = os.path.join("assets", "pfmea_dicts")
CONTROL_PLAN_TEMPLATE_PATH = os.path.join("assets", "CONTROL_PLAN_TEMPLATE.xlsx")
CONTROL_PLAN_DICT_DIR           = os.path.join("assets", "control_plan_dicts")
FIXTURE_PLAN_TEMPLATE_PATH     = os.path.join("assets", "fixture_plan_template.xlsx")

# Task tracking
TASKS = {} # task_id -> {status, progress, stl_url, files: [], message, error}

def load_tasks():
    global TASKS
    if os.path.exists(TASKS_FILE):
        try:
            with open(TASKS_FILE, "r") as f:
                TASKS = json.load(f)
        except:
            TASKS = {}

def save_tasks():
    try:
        with open(TASKS_FILE, "w") as f:
            json.dump(TASKS, f, indent=2)
    except:
        pass

# Upload retention: keep source files (CAD/feasibility/drawings/temp images)
# only for the most recently finished N tasks; older ones are deleted once a
# task completes or errors out, since nothing reads uploads after that point.
UPLOAD_RETENTION_FILE = os.path.join(DATA_DIR, "upload_retention.json")
MAX_RETAINED_UPLOADS = 3

def load_retention():
    if os.path.exists(UPLOAD_RETENTION_FILE):
        try:
            with open(UPLOAD_RETENTION_FILE, "r") as f:
                return json.load(f)
        except:
            return []
    return []

def save_retention(retained):
    try:
        with open(UPLOAD_RETENTION_FILE, "w") as f:
            json.dump(retained, f, indent=2)
    except:
        pass

def cleanup_uploads_for_task(task_id: str):
    for pattern in (f"{task_id}_*", f"tmp_imgs_{task_id}", f"tool_imgs_{task_id}"):
        for path in glob.glob(os.path.join(UPLOAD_DIR, pattern)):
            try:
                if os.path.isdir(path):
                    shutil.rmtree(path, ignore_errors=True)
                else:
                    os.remove(path)
            except Exception as e:
                print(f"[CLEANUP] Failed to remove {path}: {e}")

def register_task_completion_and_cleanup(task_id: str):
    """Track finished tasks and purge uploads for any beyond the last MAX_RETAINED_UPLOADS."""
    retained = load_retention()
    if task_id in retained:
        retained.remove(task_id)
    retained.append(task_id)
    while len(retained) > MAX_RETAINED_UPLOADS:
        old_task_id = retained.pop(0)
        print(f"[CLEANUP] Retention limit reached, removing uploads for task {old_task_id}")
        cleanup_uploads_for_task(old_task_id)
    save_retention(retained)

# Ensure directories exist
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Welding Cropper State
current_welding_row = WELDING_DATA_START_ROW

def init_welding_excel():
    """
    Working copy starts as a fresh copy of the real "Welding Check Sheet"
    template (header, Sr.No. sub-header, footer/signature rows all already
    formatted) rather than a bare "ID | Cropped Image" sheet. Each manual
    crop from the Cropper page then lands in the template's own "Required
    Welding Specifications" column (B) on the next open Sl. No row - the
    other columns (Welding Type, Measuring Instruments, etc.) are left for a
    separate fill step, not touched here.
    """
    global current_welding_row
    try:
        if not os.path.exists(WELDING_EXCEL_FILENAME):
            shutil.copy(WELDING_TEMPLATE_PATH, WELDING_EXCEL_FILENAME)

        wb = load_workbook(WELDING_EXCEL_FILENAME)
        ws = wb.active
        # Resume at the first data row that doesn't have a Sl. No yet, so a
        # server restart (or repeated /save-crop calls) doesn't overwrite
        # crops already saved in this working copy.
        row = WELDING_DATA_START_ROW
        while row <= WELDING_DATA_END_ROW and ws.cell(row=row, column=1).value not in (None, ""):
            row += 1
        current_welding_row = row
    except Exception as e:
        print(f"[WELDING] Init error: {e}")
        current_welding_row = WELDING_DATA_START_ROW

init_welding_excel()

load_tasks()

# Ensure directories exist
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

def process_engineering_task(task_id: str, cad_path: str, feasibility_path: str, drawings_dir: str, requested_outputs: dict):
    """Full pipeline: 3D -> BOM -> PFD -> Tooling."""
    try:
        print(f"\n[TASK {task_id}] Starting engineering pipeline...")
        TASKS[task_id]["status"] = "processing"
        TASKS[task_id]["files"] = []
        save_tasks()
        
        # 0. Extract Part Number from Feasibility for proper naming
        print(f"[TASK {task_id}] Extracting Part Number for naming convention...")
        part_no = "UNKNOWN"
        try:
            # 🛡️ PRIORITY 1: Filename (usually contains the Master Part Number)
            match = re.search(r'\d{8,9}', os.path.basename(feasibility_path))
            if match:
                part_no = match.group(0)
            else:
                # PRIORITY 2: Inside Excel (might be a sub-part, so secondary)
                df_temp = pd.read_excel(feasibility_path, sheet_name=0, skiprows=3)
                col_match = [c for c in df_temp.columns if 'PART NO' in str(c).upper()]
                if col_match:
                    val = df_temp[col_match[0]].dropna().iloc[0]
                    part_no = str(val).replace('.0', '').strip()
                else:
                    # PRIORITY 3: Any digits in filename
                    match = re.search(r'\d+', os.path.basename(feasibility_path))
                    if match: part_no = match.group(0)
        except Exception as e:
            print(f"   [WARN] Could not extract part number: {e}")
            
        print(f"[TASK {task_id}] Target Part Number: {part_no}")

        # 0b. Canonical Part Name/Description - extracted ONCE, reused by every
        # generator below, so all documents show the same identity instead of
        # each independently re-deriving (and potentially disagreeing on) it.
        try:
            _canon_no, part_desc = get_canonical_part_info(feasibility_path)
            if part_no == "UNKNOWN" and _canon_no != "UNKNOWN":
                part_no = _canon_no
        except Exception as e:
            print(f"   [WARN] Could not extract canonical part description: {e}")
            part_desc = "ASSEMBLY"
        print(f"[TASK {task_id}] Canonical Part Name/Description: {part_desc}")

        # 1. Export 3D STL
        TASKS[task_id]["progress"] = 5
        TASKS[task_id]["message"] = "Initializing CAD geometry..."
        print(f"[TASK {task_id}] Exporting STEP to STL: {cad_path}")
        stl_filename = f"model_{part_no}.stl"
        stl_path = os.path.join(OUTPUT_DIR, stl_filename)
        if export_step_to_stl(cad_path, stl_path):
            TASKS[task_id]["stl_url"] = f"/outputs/{stl_filename}"
            TASKS[task_id]["progress"] = 15
            save_tasks()
            print(f"[TASK {task_id}] STL export successful.")
        
        # 2. BOM Generation & Image Extraction
        # Note: Fitment Checksheet also relies on images extracted during this phase.
        if requested_outputs.get("bom") or requested_outputs.get("fitment") or requested_outputs.get("control_plan"):
            print(f"[TASK {task_id}] Starting BOM/Image generation...")
            TASKS[task_id]["message"] = "Generating Bill of Materials & Extracting Images..."
            TASKS[task_id]["progress"] = 20
            
            bom_filename = f"BOM_Result_{part_no}.xlsx"
            bom_path = os.path.join(OUTPUT_DIR, bom_filename)
            
            tmp_img_dir = os.path.join(UPLOAD_DIR, f"tmp_imgs_{task_id}")
            os.makedirs(tmp_img_dir, exist_ok=True)
            
            generate_bom_excel(
                feasibility_path,
                BOM_TEMPLATE_PATH,
                drawings_dir,
                tmp_img_dir,
                bom_path,
                part_no=part_no,
                part_desc=part_desc,
            )
            
            if requested_outputs.get("bom"):
                TASKS[task_id]["files"].append({"name": "Bill of Materials", "url": f"/outputs/{bom_filename}", "type": "xlsx"})
            
            TASKS[task_id]["progress"] = 45
            save_tasks()
            print(f"[TASK {task_id}] BOM/Image generation complete.")

        # 3. PFD Generation
        if requested_outputs.get("pfd"):
            print(f"[TASK {task_id}] Starting PFD generation pipeline...")
            TASKS[task_id]["message"] = "Generating Process Flow Diagrams..."
            excel_filename = f"PFD_Result_{part_no}.xlsx"
            excel_path = os.path.join(OUTPUT_DIR, excel_filename)
            
            shutil.copy(PFD_TEMPLATE_PATH, excel_path)
            
            print(f"[TASK {task_id}] Rendering PFD Sheet 1 (Sub-Assy Index)...")
            render_sub_assy_index(feasibility_path, cad_path, excel_path, excel_path, OUTPUT_DIR)
            TASKS[task_id]["progress"] = 60
            
            print(f"[TASK {task_id}] Rendering PFD Sheet 2 (Sub-Assy Flow)...")
            render_sub_assy_sheet(feasibility_path, excel_path, excel_path, part_no=part_no, part_desc=part_desc)
            TASKS[task_id]["progress"] = 70

            print(f"[TASK {task_id}] Rendering PFD Sheet 3 (Sheetmetal Flow)...")
            render_sheetmetal_sheet(feasibility_path, excel_path, excel_path, part_no=part_no, part_desc=part_desc)
            TASKS[task_id]["progress"] = 80

            print(f"[TASK {task_id}] Rendering PFD Sheet 4 (BOP List)...")
            render_bop_sheet(feasibility_path, excel_path, excel_path, part_no=part_no, part_desc=part_desc)
            
            TASKS[task_id]["files"].append({"name": "Engineering PFD (Excel)", "url": f"/outputs/{excel_filename}", "type": "xlsx"})
            TASKS[task_id]["progress"] = 90
            print(f"[TASK {task_id}] PFD generation complete.")

        # 4. Tooling Analysis (Fixture Master List)
        if requested_outputs.get("tooling"):
            print(f"[TASK {task_id}] Starting Tooling List generation...")
            TASKS[task_id]["message"] = "Generating Tooling Master List..."
            TASKS[task_id]["progress"] = 92
            save_tasks()
            
            tool_filename = f"Tooling_List_{part_no}.xlsx"
            tool_path = os.path.join(OUTPUT_DIR, tool_filename)
            tool_img_dir = os.path.join(UPLOAD_DIR, f"tool_imgs_{task_id}")
            
            if generate_tooling_list(feasibility_path, TOOLING_TEMPLATE_PATH, tool_path, tool_img_dir, part_no=part_no, part_desc=part_desc):
                TASKS[task_id]["files"].append({"name": "Tooling Master List", "url": f"/outputs/{tool_filename}", "type": "xlsx"})
            
            TASKS[task_id]["progress"] = 95
            save_tasks()
            print(f"[TASK {task_id}] Tooling List generation complete.")

        # 5. Fitment Checksheet
        if requested_outputs.get("fitment"):
            print(f"[TASK {task_id}] Starting Fitment Checksheet generation...")
            TASKS[task_id]["message"] = "Generating Fitment Checksheet..."
            TASKS[task_id]["progress"] = 97
            save_tasks()

            fitment_filename = f"Fitment_Checksheet_{part_no}.xlsx"
            fitment_path = os.path.join(OUTPUT_DIR, fitment_filename)

            # Use the shared image directory (ensured to be populated by Step 2)
            tmp_img_dir = os.path.join(UPLOAD_DIR, f"tmp_imgs_{task_id}")
            
            if generate_fitment_check_sheet(feasibility_path, FITMENT_TEMPLATE_PATH, fitment_path, tmp_img_dir, part_no=part_no, part_desc=part_desc):
                TASKS[task_id]["files"].append({"name": "Fitment Checksheet", "url": f"/outputs/{fitment_filename}", "type": "xlsx"})

            save_tasks()
            print(f"[TASK {task_id}] Fitment Checksheet generation complete.")

        # 6. PFMEA Generation
        if requested_outputs.get("pfmea"):
            print(f"[TASK {task_id}] Starting PFMEA generation...")
            TASKS[task_id]["message"] = "Generating PFMEA Document..."
            TASKS[task_id]["progress"] = 98
            save_tasks()

            pfmea_filename = f"PFMEA_Result_{part_no}.xlsx"
            pfmea_path = os.path.join(OUTPUT_DIR, pfmea_filename)

            if generate_pfmea_excel(feasibility_path, PFMEA_TEMPLATE_PATH, pfmea_path, PFMEA_DICT_DIR, part_no=part_no, part_desc=part_desc):
                TASKS[task_id]["files"].append({"name": "Process FMEA", "url": f"/outputs/{pfmea_filename}", "type": "xlsx"})
            
            save_tasks()
            print(f"[TASK {task_id}] PFMEA generation complete.")

        # 7. Control Plan Generation
        if requested_outputs.get("control_plan"):
            print(f"[TASK {task_id}] Starting Control Plan generation...")
            TASKS[task_id]["message"] = "Generating Control Plan..."
            TASKS[task_id]["progress"] = 99
            save_tasks()

            cp_filename = f"Control_Plan_{part_no}.xlsx"
            cp_path     = os.path.join(OUTPUT_DIR, cp_filename)

            # Resolve paths needed by the generator
            tmp_img_dir  = os.path.join(UPLOAD_DIR, f"tmp_imgs_{task_id}")
            pfd_filename = f"PFD_Result_{part_no}.xlsx"
            pfd_out_path = os.path.join(OUTPUT_DIR, pfd_filename)

            # Build sub-assy index data from feasibility (reuse same extraction logic)
            try:
                from services.pfd_renderer import extract_feasibility_data
                _, assy_data = extract_feasibility_data(feasibility_path)
            except Exception as e:
                print(f"[CP] Failed to extract sub assy data: {e}")
                assy_data = []

            generated = generate_control_plan_excel(
                feasibility_path=feasibility_path,
                template_path=CONTROL_PLAN_TEMPLATE_PATH,
                output_path=cp_path,
                dict_dir=CONTROL_PLAN_DICT_DIR,
                pfd_output_path=pfd_out_path,
                img_dir=tmp_img_dir,
                assy_data=assy_data,
                pfd_img_dir=OUTPUT_DIR,
                part_no=part_no,
                part_desc=part_desc,
            )
            if generated:
                TASKS[task_id]["files"].append({"name": "Control Plan", "url": f"/outputs/{cp_filename}", "type": "xlsx"})

            save_tasks()
            print(f"[TASK {task_id}] Control Plan generation complete.")

        # 8. Fixture PM Master Plan
        if requested_outputs.get("fixture_plan"):
            print(f"[TASK {task_id}] Starting Fixture Plan generation...")
            TASKS[task_id]["message"] = "Generating Fixture PM Master Plan..."
            TASKS[task_id]["progress"] = 99.5
            save_tasks()

            fp_filename = f"Fixture_Plan_{part_no}.xlsx"
            fp_path     = os.path.join(OUTPUT_DIR, fp_filename)

            if generate_fixture_plan(feasibility_path, FIXTURE_PLAN_TEMPLATE_PATH, fp_path, part_no=part_no, part_desc=part_desc):
                TASKS[task_id]["files"].append({"name": "Fixture PM Master Plan", "url": f"/outputs/{fp_filename}", "type": "xlsx"})

            save_tasks()
            print(f"[TASK {task_id}] Fixture Plan generation complete.")

        TASKS[task_id]["progress"] = 100
        TASKS[task_id]["status"] = "completed"
        TASKS[task_id]["message"] = "Generation complete"
        save_tasks()
        print(f"[TASK {task_id}] ALL TASKS COMPLETED SUCCESSFULLY.\n")
        register_task_completion_and_cleanup(task_id)

    except Exception as e:
        print(f"[TASK {task_id}] FATAL ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        TASKS[task_id]["status"] = "error"
        TASKS[task_id]["error"] = str(e)
        TASKS[task_id]["message"] = f"Error: {str(e)}"
        save_tasks()
        register_task_completion_and_cleanup(task_id)

@router.post("/analyze")
async def start_analysis(
    background_tasks: BackgroundTasks,
    cad_file: UploadFile = File(...),
    feasibility_file: UploadFile = File(...),
    drawings: list[UploadFile] = File(default=[]),
    outputs: str = Form("{}")
):
    task_id = str(uuid.uuid4())[:8]
    requested_outputs = json.loads(outputs)
    
    # Initialize task tracking
    TASKS[task_id] = {
        "status": "queued",
        "progress": 0,
        "message": "Task queued...",
        "stl_url": None,
        "files": [],
        "error": None
    }
    save_tasks()
    
    # Save files
    cad_path = os.path.join(UPLOAD_DIR, f"{task_id}_{cad_file.filename}")
    feas_path = os.path.join(UPLOAD_DIR, f"{task_id}_{feasibility_file.filename}")
    
    with open(cad_path, "wb") as f:
        shutil.copyfileobj(cad_file.file, f)
    with open(feas_path, "wb") as f:
        shutil.copyfileobj(feasibility_file.file, f)
        
    # Save drawings if any
    drawings_dir = os.path.join(UPLOAD_DIR, f"{task_id}_drawings")
    os.makedirs(drawings_dir, exist_ok=True)
    for drawing in drawings:
        d_path = os.path.join(drawings_dir, drawing.filename)
        with open(d_path, "wb") as f:
            shutil.copyfileobj(drawing.file, f)
            
    # Start background task
    background_tasks.add_task(process_engineering_task, task_id, cad_path, feas_path, drawings_dir, requested_outputs)
    
    return {
        "message": "Analysis started",
        "task_id": task_id
    }

@router.get("/status/{task_id}")
async def get_analysis_status(task_id: str):
    if task_id in TASKS:
        return TASKS[task_id]
    return {"status": "not_found"}

@router.post("/correct")
async def correct_document(req: BulkCorrectionRequest):
    return apply_bulk_correction(
        task_id=req.task_id,
        corrections=[c.dict() for c in req.corrections]
    )

def _compute_true_extents(excel_path: str) -> dict:
    """
    Excel's own COM bookkeeping (UsedRange, Cells.Find) is unreliable here: these
    templates pre-apply borders/fills to thousands of rows for future-proofing,
    so UsedRange balloons far past the real data (observed: a BOM sheet reported
    1003 rows via UsedRange when only 23 held any value; Cells.Find, meanwhile,
    reported as little as 1 row on sheets with 70+ populated rows, since it
    searches relative to whatever cell was active when the file was last saved).
    Scanning actual cell values with openpyxl gives the true printable extent,
    confirmed against the real generated files, per sheet.
    """
    extents = {}
    wb = load_workbook(excel_path, data_only=True)
    for ws in wb.worksheets:
        max_row = max_col = 0
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None and str(cell.value).strip() != "":
                    max_row = max(max_row, cell.row)
                    max_col = max(max_col, cell.column)
        extents[ws.title] = (max_row, max_col)
    return extents


def _export_excel_to_pdf_sync(excel_path: str, pdf_path: str) -> bool:
    """
    Blocking Excel COM automation — must run in a worker thread (via
    asyncio.to_thread), never directly on the event loop, since a single
    conversion can take 10s+ and would otherwise freeze every other request
    (downloads included) on this single-process server for that whole time.
    """
    try:
        extents = _compute_true_extents(excel_path)
    except Exception as e:
        print(f"[PREVIEW] Extent scan failed: {e}")
        extents = {}

    pythoncom.CoInitialize()
    try:
        # Use DispatchEx for a new instance and better isolation
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(excel_path)
            # Ensure all sheets are processed for full visibility
            for sheet in wb.Sheets:
                try:
                    max_row, max_col = extents.get(sheet.Name, (0, 0))
                    if max_row and max_col:
                        print_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(max_row, max_col))
                        sheet.PageSetup.PrintArea = print_range.Address
                    else:
                        # Couldn't compute it (e.g. sheet has only images, no cell values) —
                        # fall back to Excel's own bookkeeping rather than leaving it unset.
                        sheet.PageSetup.PrintArea = sheet.UsedRange.Address

                    # Ensure everything fits on one page width but can span multiple pages down
                    sheet.PageSetup.Zoom = False
                    sheet.PageSetup.FitToPagesWide = 1
                    sheet.PageSetup.FitToPagesTall = False
                except Exception as sheet_err:
                    print(f"[PREVIEW] Sheet setup error: {sheet_err}")

            # Export the entire workbook to PDF
            # 0 = xlTypePDF, 1 = xlQualityStandard, IncludeDocProperties=True, IgnorePrintAreas=False (Crucial to respect the PrintArea we set)
            wb.ExportAsFixedFormat(0, pdf_path, Quality=0, IncludeDocProperties=True, IgnorePrintAreas=False)
            wb.Close(False)
        finally:
            excel.Quit()

        return os.path.exists(pdf_path)
    except Exception as e:
        print(f"[PREVIEW ERROR] Native conversion failed: {e}")
        return False
    finally:
        try: pythoncom.CoUninitialize()
        except: pass


@router.get("/preview/{filename}")
async def preview_excel(filename: str):
    """
    High-fidelity 'Sure Shot' preview solution.
    """
    excel_path = os.path.join(OUTPUT_DIR, filename)
    pdf_filename = filename.replace(".xlsx", ".pdf")
    pdf_path = os.path.join(OUTPUT_DIR, pdf_filename)

    if not os.path.exists(excel_path):
        raise HTTPException(status_code=404, detail=f"File not found: {filename}")

    # 1. IMMEDIATE RETURN: serve the cached PDF only if it's not older than the
    # source xlsx. Each generation run overwrites the .xlsx with fresh data, so
    # a PDF cached from a previous run (or from before a bugfix) would otherwise
    # be served forever and look "stuck" no matter how many times the document
    # gets regenerated.
    if os.path.exists(pdf_path) and os.path.getmtime(pdf_path) >= os.path.getmtime(excel_path):
        return FileResponse(pdf_path, media_type="application/pdf")

    # 2. GENERATE PDF: Try native Excel conversion (offloaded to a thread so it
    # can't block the event loop and stall unrelated requests like downloads)
    if HAS_WIN32:
        success = await asyncio.to_thread(_export_excel_to_pdf_sync, excel_path, pdf_path)
        if success:
            return FileResponse(pdf_path, media_type="application/pdf")

    # 3. FALLBACK: Data View (if PDF fails)
    try:
        # Use pandas for a clean data-only fallback instead of xlsx2html which is buggy with images
        # Our templates lead with a logo/title block before the real header row, so a
        # fixed header=0 picks up "Unnamed: N" columns and NaN-filled title rows. Instead,
        # auto-detect the header row as the one with the most populated cells (title rows
        # in these templates only fill a cell or two; the real header fills every column).
        raw = pd.read_excel(excel_path, sheet_name=0, header=None)
        scan_limit = min(10, len(raw))
        header_row_idx = max(range(scan_limit), key=lambda i: raw.iloc[i].notna().sum(), default=0)

        df = pd.read_excel(excel_path, sheet_name=0, header=header_row_idx)
        df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
        df.columns = ["" if str(c).startswith("Unnamed") else str(c) for c in df.columns]
        html_content = df.to_html(classes="table table-striped", index=False, na_rep="")
        
        styled_html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: 'Segoe UI', sans-serif; padding: 20px; background: #f0f2f5; }}
                .container {{ background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); margin: auto; }}
                table {{ border-collapse: collapse; width: 100%; font-size: 13px; }}
                td, th {{ border: 1px solid #dfe3e8; padding: 10px; text-align: left; }}
                th {{ background-color: #f4f6f8; font-weight: 600; color: #454f5b; }}
            </style>
        </head>
        <body>
            <div class="container">
                <h3 style="margin-top:0">Data Preview (Low Fidelity)</h3>
                {html_content}
            </div>
        </body>
        </html>
        """
        return HTMLResponse(content=styled_html)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Preview failed: {str(e)}")


@router.post("/edit/{filename}")
async def edit_in_excel(filename: str):
    """
    Launches the real desktop Excel on the actual output file (not a copy), so
    edits save straight back to the file the app already tracks. This only
    makes sense because the backend and the user's browser are on the same
    Windows machine right now — os.startfile() opens Excel on whatever machine
    this process runs on, so this would silently open Excel on the *server*
    instead of the user's desktop if this ever runs as a remote/multi-user
    deployment.
    """
    excel_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(excel_path):
        raise HTTPException(status_code=404, detail=f"File not found: {filename}")

    if not hasattr(os, "startfile"):
        raise HTTPException(status_code=501, detail="Edit-in-Excel is only supported when the backend runs on Windows.")

    try:
        os.startfile(excel_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Could not launch Excel: {str(e)}")

    return {"status": "opened", "filename": filename}


@router.post("/save-crop")
async def save_crop(image_data: str = Form(...)):
    global current_welding_row
    try:
        if current_welding_row > WELDING_DATA_END_ROW:
            return {"status": "error", "message": "Welding check sheet is full - no more rows available. Reset to start a new sheet."}

        # Decode base64 image
        header, encoded = image_data.split(",", 1)
        data = base64.b64decode(encoded)
        img_buffer = BytesIO(data)
        image = Image.open(img_buffer)
        # A manually-dragged selection box is rarely pixel-tight around the
        # symbol - trim any blank margin so centering it in the cell below
        # actually centers the visible content, not just the outer crop.
        image = trim_whitespace(image)

        # Save image to outputs
        crop_id = str(uuid.uuid4())[:8]
        crop_filename = f"crop_{crop_id}.png"
        crop_path = os.path.join(OUTPUT_DIR, crop_filename)
        image.save(crop_path)

        # Load the working copy of the Welding Check Sheet template
        wb = load_workbook(WELDING_EXCEL_FILENAME)
        ws = wb.active

        sl_no = current_welding_row - WELDING_DATA_START_ROW + 1
        ws.cell(row=current_welding_row, column=1, value=sl_no)

        # Fit the crop into column B's own template cell size (already
        # formatted for this sheet) instead of a fixed pixel box, so it's
        # centered rather than pinned to the cell's top-left corner.
        place_image_centered(ws, crop_path, f"B{current_welding_row}", 230, 150)

        wb.save(WELDING_EXCEL_FILENAME)
        saved_row = current_welding_row
        current_welding_row += 1

        return {"status": "success", "message": f"Image saved as Sl. No {sl_no}", "row": saved_row, "sl_no": sl_no}
    except Exception as e:
        print(f"[WELDING] Save error: {e}")
        return {"status": "error", "message": str(e)}

@router.post("/welding/set-reference-image")
async def set_welding_reference_image(drawing_pdf: UploadFile = File(...)):
    """
    Fills the blank assembly-reference photo row (A7:J7) of the working
    welding check sheet from a part's own drawing PDF, using the same
    "REFERENCE ISOMETRIC VIEW" extraction pipeline BOM generation uses -
    but standalone, so this works even when no BOM has been generated for
    the part yet (the whole point of the Cropper being a separate, manual
    tool from the main NPD pipeline).
    """
    tmp_pdf_path = raw_path = clean_path = None
    try:
        tmp_pdf_path = os.path.join(UPLOAD_DIR, f"weldref_{uuid.uuid4().hex[:8]}.pdf")
        with open(tmp_pdf_path, "wb") as f:
            shutil.copyfileobj(drawing_pdf.file, f)

        raw_path = os.path.join(OUTPUT_DIR, f"weldref_raw_{uuid.uuid4().hex[:8]}.png")
        clean_path = os.path.join(OUTPUT_DIR, f"weldref_clean_{uuid.uuid4().hex[:8]}.png")

        target = extract_isometric_view(tmp_pdf_path, raw_path)
        if not target or not os.path.exists(raw_path):
            raise HTTPException(status_code=422, detail="Could not locate a reference isometric view in this drawing.")
        if not isolate_part_and_remove_text(raw_path, clean_path, target[0], target[1]):
            raise HTTPException(status_code=422, detail="Could not isolate the part from the drawing.")

        wb = load_workbook(WELDING_EXCEL_FILENAME)
        ws = wb.active

        # Replace, don't duplicate, if a reference image was already set
        for img in [i for i in ws._images if i.anchor._from.row == 6]:
            ws._images.remove(img)

        place_image_centered(ws, clean_path, "A7", 900, 220)
        wb.save(WELDING_EXCEL_FILENAME)

        return {"status": "success", "message": "Reference image set."}
    except HTTPException:
        raise
    except Exception as e:
        print(f"[WELDING] Reference image error: {e}")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        for p in (tmp_pdf_path, raw_path, clean_path):
            if p and os.path.exists(p):
                try: os.remove(p)
                except Exception: pass


@router.get("/download-excel")
async def download_welding_excel():
    if os.path.exists(WELDING_EXCEL_FILENAME):
        return FileResponse(WELDING_EXCEL_FILENAME, filename="welding_report.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    raise HTTPException(status_code=404, detail="File not found")

@router.post("/reset-excel")
async def reset_welding_excel():
    global current_welding_row
    if os.path.exists(WELDING_EXCEL_FILENAME):
        os.remove(WELDING_EXCEL_FILENAME)
    init_welding_excel()
    return {"status": "success", "message": "Excel has been reset"}
