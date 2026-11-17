import re
import openpyxl
from datetime import date
from typing import List, Dict, Any, Tuple
from .normalizer import parse_date_value

HYPERLINK_START_DATE = date(2026, 3, 23)

def parse_master_sheet(wb: openpyxl.Workbook, master_info: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    Parses the Master Sheet:
    - Identifies dates and their corresponding Page Link hyperlinks.
    - Extracts internal worksheet destinations.
    - Strictly filters for EFFILUME records (excluding other locations).
    - Returns:
        - effilume_links: List of valid EFFILUME hyperlink targets
        - skipped_entries: List of skipped/invalid/non-EFFILUME entries for the log
    """
    master_sheet_name = master_info['master_sheet_name']
    ws = wb[master_sheet_name]
    
    loc_col = master_info['location_col']
    page_col = master_info['page_link_col']
    qty_col = master_info.get('qty_col', 4)
    cat_col = master_info.get('category_col', 2)
    s_no_col = master_info.get('s_no_col', 1)
    header_row = master_info.get('header_row', 1)

    effilume_links = []
    skipped_entries = []

    for r in range(header_row + 1, ws.max_row + 1):
        loc_val = ws.cell(row=r, column=loc_col).value
        pcell = ws.cell(row=r, column=page_col)
        display_val = pcell.value
        qty_val = ws.cell(row=r, column=qty_col).value
        s_no_val = ws.cell(row=r, column=s_no_col).value
        cat_val = ws.cell(row=r, column=cat_col).value

        # Skip rows that are completely empty
        if not any([loc_val, display_val, pcell.hyperlink, qty_val, s_no_val]):
            continue

        loc_str = str(loc_val).strip() if loc_val is not None else ''
        display_str = str(display_val).strip() if display_val is not None else ''

        # Strict check: EFFILUME Only
        # Conversion entries like 'Conversion(Effilume to Vincy)' or other locations are excluded
        is_effilume = (loc_str.upper() == 'EFFILUME')

        if not is_effilume:
            skipped_entries.append({
                'row': r,
                'location': loc_str or 'Unknown',
                'date_str': display_str or 'N/A',
                'reason': f"Excluded non-EFFILUME location: '{loc_str}'",
                'status': 'Skipped (Non-EFFILUME)'
            })
            continue

        # Extract hyperlink destination
        hl_location = pcell.hyperlink.location if pcell.hyperlink else None
        target_sheet = None
        if hl_location:
            # Handle formats like "'11.04.2026'!A1" or "23.03.2026!A1"
            m = re.match(r"^'?([^'!]+)'?!", str(hl_location).strip())
            if m:
                target_sheet = m.group(1).strip()
            else:
                target_sheet = str(hl_location).strip()

        dt, date_str = parse_date_value(display_str)
        if not dt and target_sheet:
            dt, date_str = parse_date_value(target_sheet)

        if not target_sheet:
            skipped_entries.append({
                'row': r,
                'location': 'EFFILUME',
                'date_str': display_str or 'Unknown',
                'sheet_name': 'None',
                'reason': 'Missing Page Link hyperlink in Master Sheet',
                'status': 'Skipped'
            })
            continue

        effilume_links.append({
            'row': r,
            'location': 'EFFILUME',
            'display_date': display_str,
            'parsed_date': dt,
            'date_str': date_str or display_str or target_sheet,
            'target_sheet': target_sheet,
            'hyperlink_location': hl_location,
            'expected_qty': qty_val
        })

    return effilume_links, skipped_entries

