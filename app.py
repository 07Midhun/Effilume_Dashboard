import os
import shutil
import uuid
from typing import Optional
from fastapi import FastAPI, File, UploadFile, HTTPException, Query
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from core import (
    validate_workbook,
    consolidate_workbook,
    generate_consolidated_excel
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
STATIC_DIR = os.path.join(BASE_DIR, "static")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

app = FastAPI(
    title="EFFILUME Data Consolidation API",
    description="Automated Date-Wise Smart Meter UID Data Processing Engine"
)

# Mount static files
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

class ProcessRequest(BaseModel):
    filepath: str

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_path = os.path.join(TEMPLATES_DIR, "index.html")
    if not os.path.exists(index_path):
        raise HTTPException(status_code=404, detail="Dashboard template not found.")
    with open(index_path, "r", encoding="utf-8") as f:
        return f.read()

@app.get("/api/default-file")
async def get_default_file():
    """Checks if a sample or default workbook exists in the root directory."""
    default_names = [
        "BBMP smart meter UID Details New (2).xlsx",
        "BBMP smart meter UID Details New.xlsx"
    ]
    for name in default_names:
        full_path = os.path.join(BASE_DIR, name)
        if os.path.exists(full_path):
            try:
                val = validate_workbook(full_path)
                return {
                    "available": True,
                    "filename": name,
                    "filepath": full_path,
                    "total_sheets": val["total_sheets"],
                    "master_sheet": val["master_sheet_name"]
                }
            except Exception:
                pass
    return {"available": False}

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    """Uploads and validates an Excel workbook (.xlsx, .xlsm)."""
    if not file.filename:
        raise HTTPException(status_code=400, detail="No file uploaded.")

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in [".xlsx", ".xlsm"]:
        raise HTTPException(status_code=400, detail="Invalid file format. Please upload an Excel (.xlsx or .xlsm) file.")

    # Unique upload filename
    unique_id = uuid.uuid4().hex[:8]
    safe_filename = f"{unique_id}_{file.filename}"
    saved_path = os.path.join(UPLOAD_DIR, safe_filename)

    try:
        with open(saved_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to save uploaded file: {str(e)}")

    # Validate workbook
    try:
        val_result = validate_workbook(saved_path)
        return {
            "success": True,
            "filename": file.filename,
            "filepath": saved_path,
            "master_sheet_name": val_result["master_sheet_name"],
            "total_sheets": val_result["total_sheets"],
            "page_link_col": val_result["page_link_col"]
        }
    except Exception as e:
        # Cleanup invalid file
        if os.path.exists(saved_path):
            os.remove(saved_path)
        raise HTTPException(status_code=400, detail=str(e))

@app.post("/api/process")
async def process_consolidation(req: ProcessRequest):
    """Processes consolidation for the specified workbook and generates the final Excel file."""
    filepath = req.filepath
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="Workbook file not found on server.")

    try:
        # Consolidate
        result = consolidate_workbook(filepath)
        records = result["records"]
        summary = result["summary"]
        columns = result["columns"]
        processing_log = result["processing_log"]

        # Generate final output Excel
        base_name = os.path.splitext(os.path.basename(filepath))[0]
        # Remove any leading uuid prefix if present
        if "_" in base_name and len(base_name.split("_")[0]) == 8:
            base_name = "_".join(base_name.split("_")[1:])
            
        output_filename = f"EFFILUME_Consolidated_{base_name}.xlsx"
        output_filepath = os.path.join(OUTPUT_DIR, output_filename)

        generate_consolidated_excel(
            records=records,
            columns=columns,
            processing_log=processing_log,
            output_path=output_filepath
        )

        return {
            "success": True,
            "summary": summary,
            "columns": columns,
            "output_file": output_filename,
            "records": records,
            "processing_log": processing_log
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Consolidation error: {str(e)}")

@app.get("/api/download")
async def download_file(filepath: str = Query(...)):
    """Downloads the generated consolidated Excel file."""
    # Ensure filepath is just the filename or inside OUTPUT_DIR
    filename = os.path.basename(filepath)
    full_path = os.path.join(OUTPUT_DIR, filename)

    if not os.path.exists(full_path):
        raise HTTPException(status_code=404, detail="Requested file not found.")

    return FileResponse(
        path=full_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename
    )

@app.get("/health")
async def health_check():
    return {"status": "ok", "app": "EFFILUME Consolidation Dashboard"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="127.0.0.1", port=8000, reload=True)

