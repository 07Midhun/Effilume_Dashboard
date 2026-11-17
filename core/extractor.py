import re
import openpyxl
from datetime import datetime, date
from typing import List, Dict, Any, Tuple, Set, Optional
from .normalizer import (
    STANDARD_COLUMNS,
    normalize_column_name,
    parse_date_value,
    clean_cell_value
)

class DataExtractor:
    def __init__(self, wb: openpyxl.Workbook):
        self.wb = wb
        self.processed_sheets: Set[str] = set()
        self.processing_log: List[Dict[str, Any]] = []
        self.additional_columns: List[str] = []
        
        # Build whitespace & case-insensitive sheet map
        self.sheet_map: Dict[str, str] = {}
        for s in self.wb.sheetnames:
            self.sheet_map[s.strip()] = s
            self.sheet_map[s.strip().lower()] = s

    def get_sheet(self, target_name: Optional[str]) -> Optional[openpyxl.worksheet.worksheet.Worksheet]:
        if not target_name:
            return None
        if target_name in self.wb.sheetnames:
            return self.wb[target_name]
        clean = target_name.strip()
        if clean in self.wb.sheetnames:
            return self.wb[clean]
        if clean in self.sheet_map:
            return self.wb[self.sheet_map[clean]]
        if clean.lower() in self.sheet_map:
            return self.wb[self.sheet_map[clean.lower()]]
        return None

    def extract_from_sheet(
        self,
        sheet_target: str,
        date_str: str,
        parsed_date: Optional[date],
        source_type: str,
        expected_qty: Optional[int] = None
    ) -> List[Dict[str, Any]]:
        """
        Extracts records from a single worksheet:
        - Prevents duplicate worksheet processing.
        - Identifies header row automatically.
        - Maps equivalent columns into standard columns.
        - Preserves additional source columns.
        - Preserves exact source values for UIDs, serials, and dates.
        - Records processing status in the processing log.
        """
        clean_target = sheet_target.strip()
        ws = self.get_sheet(clean_target)

        if ws is None:
            self.processing_log.append({
                'date_processed': date_str,
                'source_sheet': sheet_target,
                'records_extracted': 0,
                'location': 'EFFILUME',
                'status': 'Error / Not Found',
                'reason': f"Worksheet '{sheet_target}' does not exist in workbook",
                'source_type': source_type
            })
            return []

        # Duplicate Prevention Rule
        # Avoid processing the same worksheet multiple times if multiple hyperlinks point to it
        actual_name = ws.title
        if actual_name in self.processed_sheets:
            self.processing_log.append({
                'date_processed': date_str,
                'source_sheet': actual_name,
                'records_extracted': 0,
                'location': 'EFFILUME',
                'status': 'Skipped (Duplicate Reference)',
                'reason': f"Worksheet '{actual_name}' was already processed",
                'source_type': source_type
            })
            return []

        self.processed_sheets.add(actual_name)

        # Detect Header Row (search top 6 rows)
        header_row_idx = None
        headers_map: Dict[int, str] = {} # col_idx -> normalized_col_name

        for r in range(1, min(7, ws.max_row + 1)):
            row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
            str_vals = [str(v).strip().lower() for v in row_vals if v is not None]
            
            # Check if this row looks like a header
            has_uid_or_serial = any('uid' in v or 'meter' in v or 'serial' in v for v in str_vals)
            has_sno_or_phase = any('s no' in v or 's.no' in v or 'phase' in v for v in str_vals)
            
            if has_uid_or_serial or (has_sno_or_phase and len(str_vals) >= 3):
                header_row_idx = r
                for c_idx, raw_val in enumerate(row_vals, start=1):
                    if raw_val is not None:
                        norm = normalize_column_name(raw_val)
                        if norm:
                            headers_map[c_idx] = norm
                            if norm not in STANDARD_COLUMNS and norm not in self.additional_columns:
                                self.additional_columns.append(norm)
                break

        if not header_row_idx or not headers_map:
            self.processing_log.append({
                'date_processed': date_str,
                'source_sheet': actual_name,
                'records_extracted': 0,
                'location': 'EFFILUME',
                'status': 'Skipped (Empty / No Headers)',
                'reason': f"Worksheet '{actual_name}' has no recognizable table headers or records",
                'source_type': source_type
            })
            return []

        # Extract Records
        sheet_records = []
        for r in range(header_row_idx + 1, ws.max_row + 1):
            row_dict = {}
            has_data = False
            
            for c_idx, col_name in headers_map.items():
                val = ws.cell(row=r, column=c_idx).value
                if val is not None:
                    cleaned_val = clean_cell_value(val, col_name)
                    if cleaned_val is not None and str(cleaned_val).strip() != '':
                        row_dict[col_name] = cleaned_val
                        has_data = True

            if not has_data:
                continue

            # Ensure minimum validity: must have UID or Meter Serial or S No
            uid = row_dict.get('Panel UID')
            meter = row_dict.get('Meter Serial Number')
            sno = row_dict.get('S No')
            if not uid and not meter and sno is None:
                continue

            # Build record adhering to consolidation requirements
            record: Dict[str, Any] = {
                'Date': date_str,
                '_parsed_date': parsed_date,
                'Location': 'EFFILUME',
                'S No': row_dict.get('S No'),
                'Panel UID': row_dict.get('Panel UID'),
                'Meter Serial Number': row_dict.get('Meter Serial Number'),
                'Phase': row_dict.get('Phase'),
                'Mode': row_dict.get('Mode'),
                'FW Ver': row_dict.get('FW Ver'),
                'Installable Since': row_dict.get('Installable Since'),
                'Remarks': row_dict.get('Remarks'),
                '_source_sheet': actual_name,
                '_source_row': r
            }

            # Add any additional source columns
            for extra_col in self.additional_columns:
                if extra_col in row_dict:
                    record[extra_col] = row_dict[extra_col]
                else:
                    record[extra_col] = None

            sheet_records.append(record)

        self.processing_log.append({
            'date_processed': date_str,
            'source_sheet': actual_name,
            'records_extracted': len(sheet_records),
            'location': 'EFFILUME',
            'status': 'Processed',
            'reason': f"Extracted {len(sheet_records)} records ({source_type})",
            'source_type': source_type
        })

        return sheet_records

