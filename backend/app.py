import os
import asyncio
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from credit_report import extract_credit_report_data, extract_credit_report_data_async, append_to_excel

from dotenv import load_dotenv
load_dotenv()

app = FastAPI(title="Credit Report Extractor")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BASE_DIR.parent
FRONTEND_DIR = PROJECT_ROOT / "frontend"
EXCEL_PATH = os.getenv("EXCEL_PATH", str(BASE_DIR / "reports.xlsx"))

SUPPORTED_EXTS = {".pdf", ".htm", ".html", ".png", ".jpg", ".jpeg", ".tiff", ".tif", ".bmp"}

# Serialize Excel writes; limit concurrent Azure API calls to avoid rate limits
excel_lock = asyncio.Lock()
azure_semaphore = asyncio.Semaphore(5)


@app.post("/api/upload")
async def upload_report(file: UploadFile = File(...)):
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in SUPPORTED_EXTS:
        raise HTTPException(400, f"Unsupported file type: {ext}")

    content = await file.read()
    try:
        # Azure OCR + GPT: up to 3 files processed concurrently
        async with azure_semaphore:
            extracted, flat_row = await extract_credit_report_data_async(content, file.filename)

        # Excel write: strictly serialized to prevent file corruption
        async with excel_lock:
            loop = asyncio.get_event_loop()
            await loop.run_in_executor(None, append_to_excel, extracted, flat_row, EXCEL_PATH)

        return {"success": True, "data": flat_row}
    except Exception as e:
        raise HTTPException(500, str(e))


from openpyxl import load_workbook

@app.get("/api/data")
def get_data():
    if not os.path.exists(EXCEL_PATH):
        return {"headers": [], "rows": []}
    wb = load_workbook(EXCEL_PATH, read_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    if not all_rows:
        return {"headers": [], "rows": []}
    headers = list(all_rows[0])
    data = [{"row_idx": i + 2, "cells": list(r)} for i, r in enumerate(all_rows[1:])]
    data.reverse()
    return {"headers": headers, "rows": data}


@app.delete("/api/data/{row_idx}")
def delete_row(row_idx: int):
    if not os.path.exists(EXCEL_PATH):
        raise HTTPException(404, "Excel file not found")
    if row_idx < 2:
        raise HTTPException(400, "Cannot delete header row")
    with excel_lock:
        wb = load_workbook(EXCEL_PATH)
        ws = wb.active
        if row_idx > ws.max_row:
            raise HTTPException(404, "Row not found")
        ws.delete_rows(row_idx)
        wb.save(EXCEL_PATH)
    return {"success": True}


@app.get("/api/download")
def download_excel():
    if not os.path.exists(EXCEL_PATH):
        raise HTTPException(404, "No reports processed yet")
    return FileResponse(EXCEL_PATH, filename="credit_reports.xlsx")


app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")