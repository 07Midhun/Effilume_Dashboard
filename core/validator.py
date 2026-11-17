import os
import openpyxl
from typing import Tuple, Dict, Any, Optional

def validate_workbook(file_path: str) -> Dict[str, Any]:
    """
    Validates an uploaded Excel workbook:
    - Checks file extension (.xlsx, .xlsm)
    - Verifies file readability with openpyxl
    - Identifies the Master Sheet automatically
    - Verifies the existence of the Page Link column
    Returns a dict with validation status, master sheet details, and sheet list.
    """
    if not os.path.exists(file_path):
        raise ValueError(f"File not found: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    if ext not in ['.xlsx', '.xlsm']:
        raise ValueError(f"Unsupported file format '{ext}'. Please upload an Excel (.xlsx or .xlsm) file.")

    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    except Exception as e:
        raise ValueError(f"Failed to open Excel workbook: {str(e)}")

    sheet_names = wb.sheetnames
    if not sheet_names:
        wb.close()
        raise ValueError("The uploaded workbook contains no worksheets.")

    # 1. Identify Master Sheet
    master_sheet_name = None
    # Check for exact or case-insensitive 'MASTER'
    for s in sheet_names:
        if s.strip().upper() == 'MASTER':
            master_sheet_name = s
            break

    # If no sheet explicitly named 'MASTER', search for sheet containing 'Page Link'
    wb_full = openpyxl.load_workbook(file_path, data_only=True)
    
    if not master_sheet_name:
        for s in sheet_names:
            ws = wb_full[s]
            for r in range(1, min(5, ws.max_row + 1)):
                row_vals = [str(ws.cell(row=r, column=c).value or '').strip().lower() for c in range(1, ws.max_column + 1)]
                if any('page link' in v or 'pagelink' in v for v in row_vals):
                    master_sheet_name = s
                    break
            if master_sheet_name:
                break

    if not master_sheet_name:
        wb.close()
        wb_full.close()
        raise ValueError("Master Sheet could not be identified. No worksheet named 'MASTER' or containing 'Page Link' column was found.")

    # 2. Check for Page Link column in Master Sheet
    ws_master = wb_full[master_sheet_name]
    page_link_col = None
    location_col = None
    qty_col = None
    category_col = None
    s_no_col = None
    header_row_idx = None

    for r in range(1, min(6, ws_master.max_row + 1)):
        for c in range(1, ws_master.max_column + 1):
            val = ws_master.cell(row=r, column=c).value
            if val:
                v_clean = str(val).strip().lower()
                if 'page' in v_clean and 'link' in v_clean:
                    page_link_col = c
                    header_row_idx = r
                elif 'location' in v_clean:
                    location_col = c
                elif 'qty' in v_clean or 'quantity' in v_clean:
                    qty_col = c
                elif 'category' in v_clean:
                    category_col = c
                elif 's' in v_clean and 'no' in v_clean:
                    s_no_col = c
        if page_link_col:
            break

    wb.close()
    wb_full.close()

    if not page_link_col:
        raise ValueError(f"The identified Master Sheet ('{master_sheet_name}') does not contain a 'Page Link' column.")

    return {
        'valid': True,
        'filename': os.path.basename(file_path),
        'master_sheet_name': master_sheet_name,
        'header_row': header_row_idx or 1,
        'page_link_col': page_link_col,
        'location_col': location_col or 3,
        'qty_col': qty_col or 4,
        'category_col': category_col or 2,
        's_no_col': s_no_col or 1,
        'total_sheets': len(sheet_names),
        'sheet_names': sheet_names
    }

