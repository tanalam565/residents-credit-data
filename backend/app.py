import os
import asyncio
import io
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import Workbook

from credit_report import extract_credit_report_data_async
from db import init_db, insert_report, get_all_reports, delete_report, export_all_flat_rows

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

SUPPORTED_EXTS = {".pdf", ".htm", ".html", ".png", ".jpg", ".jpeg", ".tiff", ".tif", ".bmp"}

azure_semaphore = asyncio.Semaphore(5)


@app.on_event("startup")
async def startup_event():
    init_db()


@app.post("/api/upload")
async def upload_report(file: UploadFile = File(...)):
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in SUPPORTED_EXTS:
        raise HTTPException(400, f"Unsupported file type: {ext}")

    content = await file.read()
    try:
        async with azure_semaphore:
            extracted, flat_row = await extract_credit_report_data_async(content, file.filename)

        loop = asyncio.get_event_loop()
        row_id = await loop.run_in_executor(None, insert_report, flat_row)

        return {"success": True, "data": flat_row, "id": row_id}
    except Exception as e:
        raise HTTPException(500, str(e))


@app.get("/api/data")
def get_data():
    headers, rows = get_all_reports()
    return {"headers": headers, "rows": rows}


@app.delete("/api/data")
def delete_all_rows():
    from db import get_conn
    with get_conn() as conn:
        conn.execute("DELETE FROM reports")
    return {"success": True}


@app.delete("/api/data/{row_idx}")
def delete_row(row_idx: int):
    deleted = delete_report(row_idx)
    if not deleted:
        raise HTTPException(404, "Row not found")
    return {"success": True}


@app.get("/api/download")
def download_excel():
    headers, value_lists = export_all_flat_rows()
    if not headers:
        raise HTTPException(404, "No reports yet")

    wb = Workbook()
    ws = wb.active
    ws.title = "Credit Reports"
    ws.append(headers)
    for row_vals in value_lists:
        ws.append([
            v if isinstance(v, (int, float, bool, type(None))) else str(v)
            for v in row_vals
        ])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=credit_reports.xlsx"},
    )


app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")
