import os
import re
import csv
import io
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

app = FastAPI(
    title="Sheets LLM Backend API",
    description="Sistem AI standarisasi nama kolom dan penggabungan data multi-sheet spreadsheet.",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Simpan data sementara di memori RAM server
LATEST_MEMORY_DATA: Dict[str, Any] = {}

# Ubah teks angka atau mata uang menjadi float
def parse_cell_value(val: Any) -> Any:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    text_value = str(val).strip()
    if not text_value:
        return None

    # Hapus awalan mata uang seperti Rp, Rp., IDR
    text_value = re.sub(r"^(rp\.?|idr)\s*", "", text_value, flags=re.IGNORECASE).strip()

    # Saring karakter angka, koma, titik, dan minus
    cleaned_text = re.sub(r"[^\d.,\-]", "", text_value)
    if not cleaned_text or not any(c.isdigit() for c in cleaned_text):
        return text_value

    # Jika ada titik di depan (sisa Rp.), hilangkan
    if cleaned_text.startswith("."):
        cleaned_text = cleaned_text.lstrip(".")

    # Tangani format dengan koma dan titik sekaligus (contoh: 1.500.000,00)
    if "," in cleaned_text and "." in cleaned_text:
        if cleaned_text.rfind(",") > cleaned_text.rfind("."):
            cleaned_text = cleaned_text.replace(".", "").replace(",", ".")
        else:
            cleaned_text = cleaned_text.replace(",", "")
    elif "." in cleaned_text:
        # Titik sebagai pemisah ribuan (contoh: 350.000 atau 1.200.000)
        parts = cleaned_text.split(".")
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3 and int(parts[0]) > 0):
            cleaned_text = cleaned_text.replace(".", "")
    elif "," in cleaned_text:
        # Koma sebagai pemisah ribuan atau desimal
        parts = cleaned_text.split(",")
        if len(parts) == 2 and len(parts[1]) == 3 and int(parts[0]) > 0:
            cleaned_text = cleaned_text.replace(",", "")
        else:
            cleaned_text = cleaned_text.replace(",", ".")

    try:
        return float(cleaned_text)
    except ValueError:
        return text_value
    return text_value

# Tampilkan halaman dashboard HTML
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
    raise HTTPException(status_code=404, detail="File dashboard tidak ditemukan")

# Gabungkan data dari banyak sheet menjadi 1 tabel standar via AI
@app.post("/api/consolidate", response_model=ConsolidatedDataset)
def consolidate_multi_sheets(payload: MultiSheetConsolidateRequest):
    global LATEST_MEMORY_DATA

    if not payload.sheets:
        raise HTTPException(status_code=400, detail="Daftar sheet tidak boleh kosong")

    sheets_data = [s.model_dump() for s in payload.sheets]

    try:
        harmonized = harmonize_sheets_schema(sheets_data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal memproses AI: {str(e)}")

    canonical_fields = harmonized.get("canonical_fields", [])
    sheet_mappings = harmonized.get("sheet_mappings", {})

    all_consolidated_records: List[Dict[str, Any]] = []

    for sheet in payload.sheets:
        sheet_name = sheet.sheet_name
        headers = sheet.headers
        rows = sheet.rows

        sheet_mapping = sheet_mappings.get(sheet_name, {})
        index_to_canonical: Dict[int, str] = {}
        for idx, header_name in enumerate(headers):
            canonical = sheet_mapping.get(header_name)
            if canonical:
                index_to_canonical[idx] = canonical

        for row in rows:
            record: Dict[str, Any] = {
                "_source_sheet": sheet_name
            }
            for col_field in canonical_fields:
                record[col_field] = None

            for idx, col_field in index_to_canonical.items():
                if idx < len(row):
                    record[col_field] = parse_cell_value(row[idx])

            all_consolidated_records.append(record)

    result = ConsolidatedDataset(
        status="success",
        total_sheets=len(payload.sheets),
        total_records=len(all_consolidated_records),
        canonical_fields=canonical_fields,
        sheet_mappings=sheet_mappings,
        records=all_consolidated_records
    )

    # Simpan hasil terbaru di RAM server
    LATEST_MEMORY_DATA = result.model_dump()

    return result

# Ambil data terbaru yang tersimpan di RAM server
@app.get("/api/latest-data")
def get_latest_data():
    return LATEST_MEMORY_DATA

# Tarik data dari link Google Sheets dan gabungkan langsung via AI
@app.post("/api/pull-google-sheets", response_model=ConsolidatedDataset)
async def pull_google_sheets(req: PullGoogleSheetsRequest):
    raw_url = req.sheet_url.strip()
    if "/d/" not in raw_url:
        raise HTTPException(status_code=400, detail="Link Google Sheets tidak valid. Harus mengandung /d/...")

    try:
        sheet_id = raw_url.split("/d/")[1].split("/")[0]
    except Exception:
        raise HTTPException(status_code=400, detail="Gagal mengambil ID dari link Google Sheets")

    target_tabs = req.tab_names or ["Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"]
    sheets_payload: List[SingleSheetPayload] = []

    async with httpx.AsyncClient(timeout=15.0, follow_redirects=True) as client:
        for tab in target_tabs:
            export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={tab}"
            try:
                res = await client.get(export_url)
                if res.status_code == 200 and res.text.strip():
                    csv_text = res.text.strip()
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

    return consolidate_multi_sheets(MultiSheetConsolidateRequest(sheets=sheets_payload))
