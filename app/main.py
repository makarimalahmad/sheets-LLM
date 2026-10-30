import os
import re
import csv
import io
import uuid
import httpx
from typing import Any, Dict, List
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from app.models import (
    SingleSheetPayload,
    MultiSheetConsolidateRequest,
    PullGoogleSheetsRequest,
    ConsolidatedDataset
)
from app.mapper import harmonize_sheets_schema
from app.database import (
    save_consolidated_records,
    get_latest_consolidated_records,
    clear_all_records
)

app = FastAPI(
    title="Sheets LLM Backend API",
    description="Otomatisasi standarisasi nama kolom dan penggabungan data multi-sheet spreadsheet.",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Ubah teks angka atau mata uang menjadi float
def parse_cell_value(val: Any) -> Any:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return None
    clean_s = re.sub(r"[^\d.,\-]", "", s)
    if clean_s and any(c.isdigit() for c in clean_s):
        if "," in clean_s and "." in clean_s:
            if clean_s.rfind(",") > clean_s.rfind("."):
                clean_s = clean_s.replace(".", "").replace(",", ".")
            else:
                clean_s = clean_s.replace(",", "")
        elif "," in clean_s:
            clean_s = clean_s.replace(",", ".")
        try:
            return float(clean_s)
        except ValueError:
            return s
    return s

# Tampilkan dashboard HTML
@app.get("/")
def root():
    index_file = os.path.join("dashboard", "index.html")
    if os.path.exists(index_file):
        return FileResponse(index_file)
    return {
        "status": "online",
        "service": "Sheets LLM Backend",
        "docs_url": "/docs"
    }

@app.get("/dashboard")
def view_dashboard():
    index_file = os.path.join("dashboard", "index.html")
    if os.path.exists(index_file):
        return FileResponse(index_file)
    raise HTTPException(status_code=404, detail="Dashboard file not found")

# Gabungkan data dari banyak sheet menjadi 1 tabel standar
@app.post("/api/consolidate", response_model=ConsolidatedDataset)
def consolidate_multi_sheets(payload: MultiSheetConsolidateRequest):
    if not payload.sheets:
        raise HTTPException(status_code=400, detail="Daftar sheet tidak boleh kosong")

    sheets_data = [s.model_dump() for s in payload.sheets]

    try:
        harmonized = harmonize_sheets_schema(sheets_data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal memproses data: {str(e)}")

    canonical_fields = harmonized.get("canonical_fields", [])
    sheet_mappings = harmonized.get("sheet_mappings", {})

    all_consolidated_records: List[Dict[str, Any]] = []
    batch_id = f"batch_{uuid.uuid4().hex[:8]}"

    for sheet in payload.sheets:
        s_name = sheet.sheet_name
        headers = sheet.headers
        rows = sheet.rows

        s_mapping = sheet_mappings.get(s_name, {})
        index_to_canonical: Dict[int, str] = {}
        for idx, h in enumerate(headers):
            canonical = s_mapping.get(h)
            if canonical:
                index_to_canonical[idx] = canonical

        for row in rows:
            record: Dict[str, Any] = {
                "_source_sheet": s_name
            }
            for c_field in canonical_fields:
                record[c_field] = None

            for idx, c_field in index_to_canonical.items():
                if idx < len(row):
                    record[c_field] = parse_cell_value(row[idx])

            all_consolidated_records.append(record)

    save_consolidated_records(batch_id, all_consolidated_records, canonical_fields, sheet_mappings)

    return ConsolidatedDataset(
        status="success",
        total_sheets=len(payload.sheets),
        total_records=len(all_consolidated_records),
        canonical_fields=canonical_fields,
        sheet_mappings=sheet_mappings,
        records=all_consolidated_records
    )

# Tarik data langsung dari link Google Sheets tanpa Google Apps Script
@app.post("/api/pull-google-sheets", response_model=ConsolidatedDataset)
async def pull_google_sheets(req: PullGoogleSheetsRequest):
    raw_url = req.sheet_url.strip()
    if "/d/" not in raw_url:
        raise HTTPException(status_code=400, detail="Link Google Sheets tidak valid. Harus mengandung /d/...")

    try:
        sheet_id = raw_url.split("/d/")[1].split("/")[0]
    except Exception:
        raise HTTPException(status_code=400, detail="Gagal mengambil ID dari link Google Sheets")

    # Daftar tab yang akan dicoba tarik
    target_tabs = req.tab_names or ["Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"]
    sheets_payload: List[SingleSheetPayload] = []

    async with httpx.AsyncClient(timeout=15.0, follow_redirects=True) as client:
        for tab in target_tabs:
            export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={tab}"
            try:
                res = await client.get(export_url)
                if res.status_code == 200 and res.text.strip():
                    csv_text = res.text.strip()
                    # Cek apakah response berupa HTML error dari Google login
                    if "<html" in csv_text.lower():
                        continue
                    reader = list(csv.reader(io.StringIO(csv_text)))
                    if len(reader) >= 2:
                        headers = [str(col).strip() for col in reader[0]]
                        rows = reader[1:]
                        sheets_payload.append(SingleSheetPayload(sheet_name=tab, headers=headers, rows=rows))
            except Exception:
                continue

    if not sheets_payload:
        raise HTTPException(
            status_code=400,
            detail="Gagal membaca data dari link. Pastikan akses Google Sheets diatur ke 'Siapa saja yang memiliki link dapat melihat'."
        )

    # Proses langsung dengan fungsi konsolidasi AI
    return consolidate_multi_sheets(MultiSheetConsolidateRequest(sheets=sheets_payload))

# Ambil riwayat data gabungan dari database
@app.get("/api/consolidated-records")
def list_consolidated_records(limit: int = 100):
    return get_latest_consolidated_records(limit=limit)

# Hapus semua riwayat data di database
@app.delete("/api/clear-data")
def clear_data():
    clear_all_records()
    return {"status": "success", "message": "Riwayat data berhasil dibersihkan"}
