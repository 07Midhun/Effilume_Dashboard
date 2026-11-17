import re
from datetime import datetime, date
from typing import Optional, Tuple, Any

STANDARD_COLUMNS = [
    'Date',
    'Location',
    'S No',
    'Panel UID',
    'Meter Serial Number',
    'Phase',
    'Mode',
    'FW Ver',
    'Installable Since',
    'Remarks'
]

def normalize_column_name(raw_name: Any) -> Optional[str]:
    """
    Normalizes a column header from source sheets to standard consolidated columns.
    If the column does not match standard columns, preserves the original stripped name.
    """
    if raw_name is None:
        return None
        
    clean = str(raw_name).strip()
    if not clean:
        return None
        
    lower = clean.lower()

    # S No
    if re.match(r'^s[\.\s_-]*no\.?$', lower) or lower in ['sl no', 'slno', 'serial no', 'sno']:
        return 'S No'

    # Panel UID / Panel ID
    if 'panel' in lower and ('uid' in lower or 'id' in lower):
        return 'Panel UID'
    if lower in ['uid', 'panel uid', 'panel id']:
        return 'Panel UID'

    # Meter Serial Number
    if 'meter' in lower and ('serial' in lower or 'sl' in lower or 'number' in lower or 'no' in lower):
        return 'Meter Serial Number'

    # Phase
    if lower == 'phase':
        return 'Phase'

    # Mode
    if lower == 'mode':
        return 'Mode'

    # FW Ver
    if 'fw' in lower and ('ver' in lower or 'version' in lower):
        return 'FW Ver'

    # Installable Since
    if 'installable' in lower and 'since' in lower:
        return 'Installable Since'

    # Remarks
    if 'remark' in lower:
        return 'Remarks'

    # Return original cleaned name for additional custom columns
    return clean

def parse_date_value(val: Any) -> Tuple[Optional[date], str]:
    """
    Parses a date from string, datetime, or int/float representation.
    Returns (datetime.date object, formatted display string 'DD.MM.YYYY').
    """
    if val is None:
        return None, ''

    if isinstance(val, (datetime, date)):
        dt = val.date() if isinstance(val, datetime) else val
        return dt, dt.strftime('%d.%m.%Y')

    s = str(val).strip()
    # Match standard patterns like 23.02.2026, 23/02/2026, 23-02-2026
    m = re.match(r'^(\d{1,2})[\.\/-](\d{1,2})[\.\/-](\d{2,4})$', s)
    if m:
        d = int(m.group(1))
        mth = int(m.group(2))
        yr = int(m.group(3))
        if yr < 100:
            yr += 2000
        elif 200 <= yr <= 299: # Typo like 226 -> 2026
            yr = 2000 + (yr % 100)
        elif yr > 2099: # Typo like 20226 -> 2026
            yr = 2026
            
        # Try DD.MM.YYYY first (standard Indian format)
        try:
            dt = date(yr, mth, d)
            return dt, f"{d:02d}.{mth:02d}.{yr}"
        except ValueError:
            pass

        # Try MM.DD.YYYY if DD.MM was invalid
        try:
            dt = date(yr, d, mth)
            return dt, f"{mth:02d}.{d:02d}.{yr}"
        except ValueError:
            pass

    return None, s

def clean_cell_value(val: Any, col_name: str) -> Any:
    """
    Cleans cell values according to column type:
    - Text-based UID and serial values preserved as exact strings
    - S No converted to integer if valid number
    - Datetime formatted cleanly
    - None/empty left as None or clean string
    """
    if val is None:
        return None

    if isinstance(val, str):
        val_str = val.strip()
        if not val_str:
            return None
        if col_name in ['Panel UID', 'Meter Serial Number']:
            return val_str
        return val_str

    if col_name in ['Panel UID', 'Meter Serial Number']:
        return str(val).strip()

    if col_name == 'S No':
        try:
            return int(val)
        except (ValueError, TypeError):
            return val

    if col_name == 'Installable Since':
        if isinstance(val, datetime):
            return val.strftime('%d/%m/%Y, %H:%M:%S')
        elif isinstance(val, date):
            return val.strftime('%d/%m/%Y')
        return str(val).strip()

    if isinstance(val, datetime):
        return val.strftime('%d/%m/%Y, %H:%M:%S')

    return val
