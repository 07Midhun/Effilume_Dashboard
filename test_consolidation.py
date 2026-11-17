import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from core import consolidate_workbook, generate_consolidated_excel

test_file = "BBMP smart meter UID Details New (2).xlsx"
print(f"Testing consolidation on: {test_file}")

result = consolidate_workbook(test_file)
summary = result['summary']
records = result['records']
log = result['processing_log']
columns = result['columns']

print("\n--- Summary Metrics ---")
for k, v in summary.items():
    print(f"  {k}: {v}")

print(f"\nTotal records: {len(records)}")
assert len(records) > 0, "No records extracted!"

# Check first 3 and last 3 records
print("\nFirst 3 records:")
for r in records[:3]:
    print(" ", {k: v for k, v in r.items() if not k.startswith('_') and v is not None})

print("\nLast 3 records:")
for r in records[-3:]:
    print(" ", {k: v for k, v in r.items() if not k.startswith('_') and v is not None})

# Verify chronological ordering
dates_seen = [r.get('_parsed_date') for r in records if r.get('_parsed_date')]
for i in range(len(dates_seen) - 1):
    assert dates_seen[i] <= dates_seen[i+1], f"Chronological order violation: {dates_seen[i]} > {dates_seen[i+1]}"
print("Chronological order verified: all dates ascending!")

# Verify EFFILUME only
for r in records:
    assert r['Location'] == 'EFFILUME', f"Non-EFFILUME record found: {r['Location']}"
print("EFFILUME filtering verified: 100% of records belong to EFFILUME!")

# Test Excel Generation
output_excel = "output_test_EFFILUME_Consolidated.xlsx"
generate_consolidated_excel(
    records=records,
    columns=columns,
    processing_log=log,
    output_path=output_excel
)
assert os.path.exists(output_excel), "Generated Excel file not found!"
print(f"\nGenerated Excel workbook saved to: {output_excel}")

# Inspect Generated Excel
wb_gen = openpyxl.load_workbook(output_excel, data_only=True)
print("Generated sheets:", wb_gen.sheetnames)
assert "EFFILUME_Consolidated_Data" in wb_gen.sheetnames, "Primary sheet name missing!"
assert "Processing_Log" in wb_gen.sheetnames, "Processing_Log sheet missing!"

ws_gen = wb_gen["EFFILUME_Consolidated_Data"]
print(f"Primary sheet rows: {ws_gen.max_row}, cols: {ws_gen.max_column}")
headers = [ws_gen.cell(row=1, column=c).value for c in range(1, ws_gen.max_column + 1)]
print("Primary headers:", headers)
assert "Date" in headers
assert "Location" in headers
assert "Panel UID" in headers
assert "Meter Serial Number" in headers

# Check freeze panes and filter
assert ws_gen.freeze_panes == "A2", f"Freeze panes not set properly: {ws_gen.freeze_panes}"
assert ws_gen.auto_filter.ref is not None, "Auto filter not enabled!"
print(f"Freeze panes: {ws_gen.freeze_panes}, Auto-filter ref: {ws_gen.auto_filter.ref}")

wb_gen.close()
print("\nALL CONSOLIDATION AND EXCEL GENERATION TESTS PASSED!")
