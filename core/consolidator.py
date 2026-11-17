import openpyxl
from datetime import date
from typing import Dict, Any, List
from .validator import validate_workbook
from .master_parser import parse_master_sheet
from .historical_finder import find_historical_worksheets
from .extractor import DataExtractor
from .normalizer import STANDARD_COLUMNS

def consolidate_workbook(file_path: str) -> Dict[str, Any]:
    """
    Main consolidation engine:
    1. Validates the uploaded workbook.
    2. Identifies the Master Sheet and Page Link column.
    3. Finds historical worksheets (23.02.2026 to 17.03.2026).
    4. Parses hyperlinks from Master Sheet for EFFILUME entries (23.03.2026 onwards).
    5. Extracts records without duplicate worksheet processing.
    6. Sorts all records chronologically.
    7. Compiles processing metrics and detailed processing log.
    """
    # 1. Validation & Master Identification
    val_res = validate_workbook(file_path)

    wb = openpyxl.load_workbook(file_path, data_only=True)
    extractor = DataExtractor(wb)

    # 2. Historical Worksheet Discovery
    historical_worksheets = find_historical_worksheets(
        sheet_names=val_res['sheet_names'],
        master_sheet_name=val_res['master_sheet_name']
    )

    # 3. Master Hyperlink Parsing
    eff_hyperlinks, skipped_master_entries = parse_master_sheet(wb, val_res)

    # Add skipped master entries to extractor log
    for skip in skipped_master_entries:
        if skip.get('status') != 'Skipped (Non-EFFILUME)':
            # Log missing hyperlinks or broken refs for EFFILUME
            extractor.processing_log.append({
                'date_processed': skip.get('date_str', 'N/A'),
                'source_sheet': skip.get('sheet_name', 'None'),
                'records_extracted': 0,
                'location': 'EFFILUME',
                'status': skip.get('status', 'Skipped'),
                'reason': skip.get('reason', 'Skipped'),
                'source_type': 'Master Sheet Reference'
            })

    all_records: List[Dict[str, Any]] = []

    # 4. Phase 1: Process Historical Worksheets
    for h in historical_worksheets:
        recs = extractor.extract_from_sheet(
            sheet_target=h['sheet_name'],
            date_str=h['date_str'],
            parsed_date=h['parsed_date'],
            source_type='Historical Sheet'
        )
        all_records.extend(recs)

    # 5. Phase 2: Process Hyperlink-based Worksheets (EFFILUME only)
    for link in eff_hyperlinks:
        recs = extractor.extract_from_sheet(
            sheet_target=link['target_sheet'],
            date_str=link['date_str'],
            parsed_date=link['parsed_date'],
            source_type='Hyperlinked Sheet',
            expected_qty=link.get('expected_qty')
        )
        all_records.extend(recs)

    wb.close()

    # 6. Chronological Sorting
    # Arrange all records in chronological order while maintaining original serial numbering
    all_records.sort(
        key=lambda r: (
            r.get('_parsed_date') or date.min,
            r.get('_source_sheet', ''),
            r.get('_source_row', 0)
        )
    )

    # Update columns list
    final_columns = list(STANDARD_COLUMNS)
    for extra in extractor.additional_columns:
        if extra not in final_columns:
            final_columns.append(extra)

    # Fill missing extra column keys in each record with None
    for r in all_records:
        for col in final_columns:
            if col not in r:
                r[col] = None

    # Calculate summary metrics
    total_processed = sum(1 for entry in extractor.processing_log if entry['status'] == 'Processed')
    total_skipped = sum(1 for entry in extractor.processing_log if 'Skipped' in entry['status'])
    total_errors = sum(1 for entry in extractor.processing_log if 'Error' in entry['status'])

    summary = {
        'filename': val_res['filename'],
        'total_sheets_in_workbook': val_res['total_sheets'],
        'master_sheet': val_res['master_sheet_name'],
        'historical_sheets_count': len(historical_worksheets),
        'hyperlink_sheets_count': len(eff_hyperlinks),
        'worksheets_processed': total_processed,
        'worksheets_skipped': total_skipped,
        'processing_errors': total_errors,
        'total_records_extracted': len(all_records),
        'columns': final_columns,
        'additional_columns': extractor.additional_columns
    }

    return {
        'summary': summary,
        'records': all_records,
        'processing_log': extractor.processing_log,
        'columns': final_columns
    }

