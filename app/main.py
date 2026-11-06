import os
import re
import csv
import io
import httpx
from typing import Any, Dict, List, Optional
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse

from app.models import (
    SingleSheetPayload,
    MultiSheetConsolidateRequest,
    PullGoogleSheetsRequest,
    ConsolidatedDataset,
    NormalizedTable,
    NormalizedResponse,
    NormalizeRequest
)
from app.mapper import harmonize_sheets_schema
from app.normalizer import normalize_dataset, export_normalized_to_excel_buffer

app = FastAPI(
    title="Sheets LLM Backend API",
    description="Sistem AI standarisasi nama kolom, penggabungan data multi-sheet, dan normalisasi relasional (ERD).",
    version="2.1.0"
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
LATEST_NORMALIZED_DATA: Dict[str, Any] = {}

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
    if not cleaned_text or not any(char.isdigit() for char in cleaned_text):
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

    sheets_data = [sheet_item.model_dump() for sheet_item in payload.sheets]

    try:
        harmonized = harmonize_sheets_schema(sheets_data)
    except Exception as error_detail:
        raise HTTPException(status_code=500, detail=f"Gagal memproses AI: {str(error_detail)}")

    canonical_fields = harmonized.get("canonical_fields", [])
    sheet_mappings = harmonized.get("sheet_mappings", {})

    all_consolidated_records: List[Dict[str, Any]] = []

    for sheet in payload.sheets:
        sheet_name = sheet.sheet_name
        headers = sheet.headers
        rows = sheet.rows

        sheet_mapping = sheet_mappings.get(sheet_name, {})
        index_to_canonical: Dict[int, str] = {}
        for col_index, header_name in enumerate(headers):
            canonical = sheet_mapping.get(header_name)
            if canonical:
                index_to_canonical[col_index] = canonical

        for row in rows:
            record: Dict[str, Any] = {
                "_source_sheet": sheet_name
            }
            for col_field in canonical_fields:
                record[col_field] = None

            for col_index, col_field in index_to_canonical.items():
                if col_index < len(row):
                    record[col_field] = parse_cell_value(row[col_index])

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

    sheets_payload: List[SingleSheetPayload] = []
    seen_sheet_contents = set()

    async with httpx.AsyncClient(timeout=15.0, follow_redirects=True) as client:
        detected_tabs: List[str] = []
        try:
            res_page = await client.get(f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit")
            if res_page.status_code == 200:
                detected_tabs = re.findall(r'docs-sheet-tab-caption">([^<]+)</div>', res_page.text)
        except Exception:
            pass

        default_fallback_tabs = [f"Sheet{i}" for i in range(1, 21)]
        target_tabs = req.tab_names or detected_tabs or default_fallback_tabs

        for tab in target_tabs:
            export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={tab}"
            try:
                res = await client.get(export_url)
                if res.status_code == 200 and res.text.strip():
                    csv_text = res.text.strip()
                    if "<html" in csv_text.lower():
                        continue

                    if csv_text in seen_sheet_contents:
                        continue
                    seen_sheet_contents.add(csv_text)

                    reader = list(csv.reader(io.StringIO(csv_text)))
                    if len(reader) >= 2:
                        headers = [str(column_name).strip() for column_name in reader[0]]
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

# Normalisasikan data tabel menjadi tabel-tabel relasional (ERD) via AI
@app.post("/api/normalize", response_model=NormalizedResponse)
async def normalize_table_data(req: Optional[NormalizeRequest] = None):
    global LATEST_NORMALIZED_DATA

    records_to_normalize: List[Dict[str, Any]] = []

    if req and req.records:
        records_to_normalize = req.records
    elif req and req.sheet_url:
        consolidated = await pull_google_sheets(PullGoogleSheetsRequest(sheet_url=req.sheet_url))
        records_to_normalize = consolidated.records
    elif LATEST_MEMORY_DATA and LATEST_MEMORY_DATA.get("records"):
        records_to_normalize = LATEST_MEMORY_DATA.get("records", [])

    if not records_to_normalize:
        raise HTTPException(status_code=400, detail="Tempel link Google Sheets terlebih dahulu.")

    try:
        normalized_result = normalize_dataset(records_to_normalize)
    except Exception as error_detail:
        raise HTTPException(status_code=500, detail=f"Gagal melakukan normalisasi AI: {str(error_detail)}")

    LATEST_NORMALIZED_DATA = normalized_result
    return NormalizedResponse(
        status="success",
        erd_summary=normalized_result.get("erd_summary", "Relasi normalisasi berhasil dibentuk"),
        tables=normalized_result.get("tables", [])
    )

# Ambil hasil normalisasi terakhir dari memori RAM
@app.get("/api/latest-normalized", response_model=NormalizedResponse)
def get_latest_normalized():
    if not LATEST_NORMALIZED_DATA or not LATEST_NORMALIZED_DATA.get("tables"):
        raise HTTPException(status_code=404, detail="Belum ada data normalisasi tersimpan")
    return NormalizedResponse(
        status="success",
        erd_summary=LATEST_NORMALIZED_DATA.get("erd_summary", ""),
        tables=LATEST_NORMALIZED_DATA.get("tables", [])
    )

# Export seluruh tabel relasional hasil normalisasi ke 1 file Excel multi-tab
@app.get("/api/export-normalized-excel")
def export_normalized_excel():
    if not LATEST_NORMALIZED_DATA or not LATEST_NORMALIZED_DATA.get("tables"):
        raise HTTPException(status_code=404, detail="Belum ada data normalisasi untuk diekspor")

    tables = LATEST_NORMALIZED_DATA.get("tables", [])
    stream = export_normalized_to_excel_buffer(tables)

    return StreamingResponse(
        stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=data_ternormalisasi_erd.xlsx"}
    )
