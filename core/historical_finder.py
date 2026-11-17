import re
from datetime import date
from typing import List, Dict, Any
from .normalizer import parse_date_value

# Fixed business rules as per Section 2.B and Section 5.B
HISTORICAL_START = date(2026, 2, 23)
HISTORICAL_END = date(2026, 3, 17)

def find_historical_worksheets(sheet_names: List[str], master_sheet_name: str = 'MASTER') -> List[Dict[str, Any]]:
    """
    Discovers historical worksheets covering 23.02.2026 to 17.03.2026.
    These worksheets are identified directly from the workbook's existing sheet names,
    even though their hyperlinks are missing from the Master Sheet.
    
    Rules:
    - Extract only worksheets that actually exist.
    - Do not generate records for dates without corresponding source worksheets.
    - Preserve original dates associated with each worksheet.
    - Assign EFFILUME as the Location.
    """
    historical_sheets = []
    
    # Find boundary: index of the first sheet that begins hyperlink processing (e.g. 23.03.2026)
    cutoff_index = len(sheet_names)
    for idx, s in enumerate(sheet_names):
        s_clean = s.strip()
        if s_clean.startswith('23.03.2026'):
            cutoff_index = idx
            break

    for idx, sname in enumerate(sheet_names):
        if idx >= cutoff_index:
            # Beyond historical tab boundary
            break

        clean_name = sname.strip()
        if clean_name.upper() == master_sheet_name.upper() or clean_name.lower().startswith('sheet'):
            continue

        # Parse date strictly as DD.MM.YYYY
        m = re.match(r'^\s*(\d{1,2})[\.\/-](\d{1,2})[\.\/-](\d{4})\s*$', clean_name)
        if not m:
            continue

        d, mth, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if yr < 100:
            yr += 2000
        elif yr > 2099:
            yr = 2026

        is_historical = False
        dt = None
        
        # Primary check: DD.MM.YYYY in range [HISTORICAL_START, HISTORICAL_END]
        try:
            primary_dt = date(yr, mth, d)
            if HISTORICAL_START <= primary_dt <= HISTORICAL_END:
                is_historical = True
                dt = primary_dt
        except ValueError:
            pass

        # Secondary check for sheets like '03.02.2026' placed between 28.02 and 05.03
        if not is_historical:
            try:
                alt_dt = date(yr, d, mth)
                if HISTORICAL_START <= alt_dt <= HISTORICAL_END:
                    is_historical = True
                    dt = alt_dt
            except ValueError:
                pass

        if is_historical:
            historical_sheets.append({
                'sheet_name': sname,
                'parsed_date': dt,
                'date_str': clean_name,
                'source_type': 'Historical Worksheet',
                'tab_index': idx
            })

    # Sort chronologically by parsed_date; if dates match, by tab_index
    historical_sheets.sort(key=lambda x: (x['parsed_date'] or date.min, x['tab_index']))
    return historical_sheets
