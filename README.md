# EFFILUME Automated Excel Consolidation Dashboard

Automated date-wise smart meter UID data processing and consolidation system for the **SCHNELL BBMP Project**.

---

## 1. Project Overview

The dashboard automates the extraction and consolidation of date-wise smart meter records exclusively for the **EFFILUME** location from an uploaded master workbook.

- **Primary Consolidated Excel Deliverable**: [`EFFILUME_Consolidated_BBMP_smart_meter_UID_Details.xlsx`](file:///c:/MIDHUN/SCHNELL/Effilume_sheet_automation/EFFILUME_Consolidated_BBMP_smart_meter_UID_Details.xlsx)
- **Primary Worksheet**: `EFFILUME_Consolidated_Data`
- **Total EFFILUME Records Extracted**: **6,231 records**
- **Date Range Covered**: **23.02.2026 to 31.08.2026** (arranged in strict chronological order)
- **Duplicate Prevention**: 100% verified, avoiding duplicate processing of multi-referenced worksheets while preserving all legitimate records.
- **Audit Sheet**: `Processing_Log` worksheet included in the Excel deliverable detailing every worksheet, status, and reason.

---

## 2. Fixed Business Rules Implemented

1. **Target Location**: Exclusively `EFFILUME` (other locations such as VINCY GLOBAL, TRANSGLOBAL, SWAMY ELECTRICALS, etc., are completely excluded).
2. **Historical Date Extraction**:
   - Covers unlinked worksheets from **23.02.2026 to 17.03.2026** directly from workbook sheet names.
   - Assigned Location: `EFFILUME`.
   - Extracted 12 historical worksheets (698 records).
3. **Hyperlink-Based Extraction (23.03.2026 Onwards)**:
   - Reads the Master Sheet `Page Link` column hyperlinks.
   - Resolves the internal worksheet destination.
   - Verifies EFFILUME data and extracts records.
   - Processed 54 hyperlinked worksheets (5,533 records).
4. **Data Integrity & Column Handling**:
   - Normalized columns: `Date`, `Location`, `S No`, `Panel UID`, `Meter Serial Number`, `Phase`, `Mode`, `FW Ver`, `Installable Since`, `Remarks`.
   - Preserved additional source columns (`Critical Points checking`, `FOR UID 2.8MM NUT IS ASSEMBED IN DOOR SENSOR`).
   - Panel UID and Meter Serial values preserved strictly as text format (`@`) without numeric or scientific formatting.
   - Original serial numbers retained for each date.

---

## 3. Project Structure

```
Effilume_sheet_automation/
├── app.py                                                   # FastAPI web server and REST API
├── run_server.py                                            # Dashboard startup script
├── test_consolidation.py                                    # Automated consolidation verification test suite
├── BBMP smart meter UID Details New (2).xlsx               # Original source workbook
├── EFFILUME_Consolidated_BBMP_smart_meter_UID_Details.xlsx # Final consolidated Excel deliverable
├── core/
│   ├── __init__.py
│   ├── validator.py                                         # Excel upload and Master Sheet validation
│   ├── normalizer.py                                        # Column normalization and date formatting
│   ├── historical_finder.py                                 # Historical worksheet discovery (23.02.2026 to 17.03.2026)
│   ├── master_parser.py                                     # Master Sheet hyperlink parsing & EFFILUME filtering
│   ├── extractor.py                                         # Worksheet record extractor & duplicate prevention
│   ├── consolidator.py                                      # Chronological consolidation engine
│   └── excel_generator.py                                   # Professional openpyxl Excel generator
├── static/
│   ├── css/styles.css                                       # Responsive dashboard styling
│   └── js/app.js                                            # Drag & drop upload, search, filters, pagination, API calls
├── templates/
│   └── index.html                                           # Dashboard web interface
├── outputs/                                                 # Generated consolidated Excel files for download
└── uploads/                                                 # Temporary uploaded workbooks
```

---

## 4. How to Launch and Use the Dashboard

### Step 1: Run the Server
Open a terminal in `c:\MIDHUN\SCHNELL\Effilume_sheet_automation` and run:
```bash
python run_server.py
```

### Step 2: Open Dashboard in Browser
Navigate to:
```
http://127.0.0.1:8000
```

### Step 3: Upload or Select Workbook
- Click **"Use This File"** to immediately process the existing `BBMP smart meter UID Details New (2).xlsx`.
- Or drag-and-drop any new `.xlsx` / `.xlsm` workbook into the upload drop zone.

### Step 4: Click "Start Consolidation"
- The dashboard extracts and consolidates the records chronologically.
- View live progress, metrics, and processing counts.
- Search and filter records by Panel UID, Meter Serial Number, or Date.
- Inspect the interactive **Processing Audit Log** tab.

### Step 5: Download Consolidated Excel
- Click **"Download EFFILUME Consolidated Excel"** to download the final formatted workbook.

---

## 5. Verification & Test Suite

Run the automated test suite at any time:
```bash
python test_consolidation.py
```
This tests:
- Master sheet identification
- Historical sheet discovery
- Hyperlink extraction
- EFFILUME filtering
- Chronological sorting
- Professional Excel styling and auto-filter checks

