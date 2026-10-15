"""Canonical part-number/description extraction, shared across every document generator.

Each generator (BOM, PFD, Tooling, Fitment, PFMEA, Control Plan, Fixture Plan) used to
independently re-derive the part number/description from the feasibility file, which is
why the same job could show different Part Name/Description text in different output
documents. routes.py now extracts these values once via get_canonical_part_info() and
passes them into every generator, which use them in place of their own extraction when
provided.
"""
import os
import re
import pandas as pd


def get_canonical_part_info(feasibility_path: str):
    """Returns a single (part_no, part_desc) pair to be reused across all generated documents."""
    part_no = "UNKNOWN"
    part_desc = "ASSEMBLY"
    try:
        filename_match = re.search(r"\d{8,9}", os.path.basename(feasibility_path))
        if filename_match:
            part_no = filename_match.group(0)

        df = pd.read_excel(feasibility_path, sheet_name=0, skiprows=3, engine="openpyxl")
        df.columns = df.columns.astype(str).str.strip().str.replace("\n", "")

        if part_no == "UNKNOWN":
            part_no_cols = [c for c in df.columns if "PART NO" in str(c).upper()]
            if part_no_cols:
                vals = df[part_no_cols[0]].dropna()
                if not vals.empty:
                    part_no = str(vals.iloc[0]).replace(".0", "").strip()

        # "Part Name" holds the top-level assembly label (e.g. "DOOR, ASSY, LH, 952NX"),
        # populated on the last data row - more reliable across real files than the
        # BOM-Level-0-row lookup some generators used, which is often absent.
        part_name_cols = [c for c in df.columns if str(c).upper().strip() == "PART NAME"]
        if part_name_cols:
            vals = [str(v).strip() for v in df[part_name_cols[0]].dropna()
                    if str(v).strip().lower() not in ("nan", "")]
            if vals:
                part_desc = vals[-1]
    except Exception as e:
        print(f"[part_info] Could not extract canonical part info: {e}")

    return part_no, part_desc


def extract_program_code(part_desc: str) -> str:
    """Pulls the trailing program/customer code out of an assembly description.

    e.g. "DOOR, ASSY, LH, 952NX" -> "952-NX"
    """
    if not part_desc:
        return ""
    last_token = part_desc.split(",")[-1].strip()
    match = re.match(r"^(\d+)([A-Za-z]+)$", last_token)
    if match:
        return f"{match.group(1)}-{match.group(2).upper()}"
    return last_token
