import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from typing import List, Dict, Any, Optional

def generate_consolidated_excel(
    records: List[Dict[str, Any]],
    columns: List[str],
    processing_log: Optional[List[Dict[str, Any]]] = None,
    output_path: Optional[str] = None
) -> bytes:
    """
    Generates a professionally styled Excel workbook containing:
    - Primary worksheet: 'EFFILUME_Consolidated_Data'
    - Audit worksheet: 'Processing_Log'
    
    Styling:
    - Bold headers with corporate deep-blue fill (#1E3A8A) and white text
    - Frozen header row
    - Active auto-filters
    - Alternating row striping (#F8FAFC)
    - Auto-fitted column widths
    - Explicit text format for UIDs and Meter Serials
    """
    wb = openpyxl.Workbook()
    
    # 1. Primary Sheet: EFFILUME_Consolidated_Data
    ws = wb.active
    ws.title = "EFFILUME_Consolidated_Data"

    # Styling definitions
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    zebra_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    regular_font = Font(name="Calibri", size=10, color="000000")
    
    thin_border_side = Side(border_style="thin", color="E2E8F0")
    cell_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side
    )
    
    header_border = Border(
        left=Side(border_style="thin", color="334155"),
        right=Side(border_style="thin", color="334155"),
        top=Side(border_style="medium", color="0F172A"),
        bottom=Side(border_style="medium", color="0F172A")
    )

    # Filter out internal keys starting with '_'
    export_columns = [c for c in columns if not c.startswith('_')]

    # Write Headers
    for col_idx, col_name in enumerate(export_columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = header_border

    ws.row_dimensions[1].height = 28

    # Write Records
    for row_idx, record in enumerate(records, start=2):
        is_even = (row_idx % 2 == 0)
        row_fill = None if is_even else zebra_fill

        for col_idx, col_name in enumerate(export_columns, start=1):
            val = record.get(col_name)
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = regular_font
            cell.border = cell_border
            if row_fill:
                cell.fill = row_fill

            # Format cell according to data type and business rules
            if col_name in ['Panel UID', 'Meter Serial Number']:
                cell.number_format = '@' # Text format to prevent scientific notation
                cell.value = str(val) if val is not None else ''
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_name == 'S No':
                cell.value = val
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name in ['Date', 'Phase', 'Mode', 'FW Ver']:
                cell.value = val
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == 'Location':
                cell.value = val or 'EFFILUME'
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.value = val
                cell.alignment = Alignment(horizontal="left", vertical="center")

        ws.row_dimensions[row_idx].height = 20

    # Auto-fit Column Widths
    for col_idx, col_name in enumerate(export_columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(col_name)
        # sample up to 100 rows for length calculation
        for r in range(2, min(100, len(records) + 2)):
            c_val = ws.cell(row=r, column=col_idx).value
            if c_val:
                max_len = max(max_len, len(str(c_val)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Freeze Header Row (row 2 top)
    ws.freeze_panes = "A2"

    # Enable Excel Auto-Filter
    if records:
        last_col_letter = get_column_letter(len(export_columns))
        last_row = len(records) + 1
        ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    # 2. Audit Sheet: Processing_Log
    if processing_log:
        ws_log = wb.create_sheet(title="Processing_Log")
        log_headers = [
            'Date Processed',
            'Source Worksheet',
            'Source Type',
            'Location',
            'Records Extracted',
            'Status',
            'Details / Reason'
        ]
        log_header_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
        
        for col_idx, lh in enumerate(log_headers, start=1):
            cell = ws_log.cell(row=1, column=col_idx, value=lh)
            cell.font = header_font
            cell.fill = log_header_fill
            cell.alignment = header_alignment
            cell.border = header_border

        ws_log.row_dimensions[1].height = 26

        for r_idx, entry in enumerate(processing_log, start=2):
            ws_log.cell(row=r_idx, column=1, value=entry.get('date_processed', ''))
            ws_log.cell(row=r_idx, column=2, value=entry.get('source_sheet', ''))
            ws_log.cell(row=r_idx, column=3, value=entry.get('source_type', ''))
            ws_log.cell(row=r_idx, column=4, value=entry.get('location', 'EFFILUME'))
            ws_log.cell(row=r_idx, column=5, value=entry.get('records_extracted', 0))
            
            st_cell = ws_log.cell(row=r_idx, column=6, value=entry.get('status', ''))
            # Color code status
            if entry.get('status') == 'Processed':
                st_cell.font = Font(name="Calibri", size=10, bold=True, color="166534") # Green
            elif 'Error' in entry.get('status', ''):
                st_cell.font = Font(name="Calibri", size=10, bold=True, color="991B1B") # Red
            else:
                st_cell.font = Font(name="Calibri", size=10, bold=True, color="9A3412") # Amber

            ws_log.cell(row=r_idx, column=7, value=entry.get('reason', ''))

            for c in range(1, 8):
                ws_log.cell(row=r_idx, column=c).border = cell_border

            ws_log.row_dimensions[r_idx].height = 20

        # Auto-fit log columns
        for c in range(1, 8):
            col_letter = get_column_letter(c)
            ws_log.column_dimensions[col_letter].width = 22

        ws_log.freeze_panes = "A2"
        ws_log.auto_filter.ref = f"A1:G{len(processing_log) + 1}"

    # Output to file or memory
    if output_path:
        wb.save(output_path)
        wb.close()
        return b''

    bio = io.BytesIO()
    wb.save(bio)
    wb.close()
    bio.seek(0)
    return bio.getvalue()

