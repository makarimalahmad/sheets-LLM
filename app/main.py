import os
import re
import uuid
from typing import Any, Dict, List, Optional
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from app.models import (
    SheetIngestRequest,
    IngestResponse,
    MultiSheetConsolidateRequest,
    ConsolidatedDataset,
    HarmonizationResult,
    StandardProductItem,
    ProductAnalytics
)
from app.mapper import harmonize_sheets_schema, map_single_sheet_headers
from app.database import (
    save_products,
    get_all_products,
    get_highest_eup,
    get_analytics,
    save_consolidated_records,
    get_latest_consolidated_records
)

app = FastAPI(
    title="Sheets LLM Backend API",
    description="Otomatisasi pembaca Google Sheets, zero-shot LLM schema harmonization, dan konsolidasi data multi-sheet.",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def parse_cell_value(val: Any) -> Any:
    """Konversi cerdas: jika angka/mata uang ubah ke float, jika teks pertahankan string."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return None
    # Cek apakah format angka atau mata uang
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

@app.get("/")
def root():
    index_file = os.path.join("dashboard", "index.html")
    if os.path.exists(index_file):
        return FileResponse(index_file)
    return {
        "status": "online",
        "service": "Sheets LLM Backend v2 (Multi-Sheet Harmonizer)",
        "docs_url": "/docs"
    }

@app.get("/dashboard")
def view_dashboard():
    index_file = os.path.join("dashboard", "index.html")
    if os.path.exists(index_file):
        return FileResponse(index_file)
    raise HTTPException(status_code=404, detail="Dashboard file not found")

@app.post("/api/harmonize-schema", response_model=HarmonizationResult)
def preview_harmonization(payload: MultiSheetConsolidateRequest):
    """
    Zero-Shot Schema Discovery via LLM.
    Menerima N sheet dengan nama kolom acak, mengembalikan skema kanonikal & pemetaan tiap sheet.
    Tanpa kamus manual.
    """
    if not payload.sheets:
        raise HTTPException(status_code=400, detail="Daftar sheet tidak boleh kosong")

    sheets_data = [s.model_dump() for s in payload.sheets]
    try:
        harmonized = harmonize_sheets_schema(sheets_data)
        return HarmonizationResult(
            canonical_fields=harmonized["canonical_fields"],
            sheet_mappings=harmonized["sheet_mappings"]
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal melakukan harmonisasi LLM: {str(e)}")

@app.post("/api/consolidate", response_model=ConsolidatedDataset)
def consolidate_multi_sheets(payload: MultiSheetConsolidateRequest):
    """
    Konsolidasi Multi-Sheet ke 1 Tabel Baku Terstruktur.
    1. Kirim semua sheet ke LLM untuk harmonisasi skema serentak.
    2. Transformasi semua baris data dari masing-masing sheet ke skema kanonikal.
    3. Simpan ke database & kembalikan 1 dataset utuh yang seragam.
    """
    if not payload.sheets:
        raise HTTPException(status_code=400, detail="Daftar sheet tidak boleh kosong")

    sheets_data = [s.model_dump() for s in payload.sheets]

    try:
        harmonized = harmonize_sheets_schema(sheets_data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal harmonisasi via LLM: {str(e)}")

    canonical_fields = harmonized.get("canonical_fields", [])
    sheet_mappings = harmonized.get("sheet_mappings", {})

    all_consolidated_records: List[Dict[str, Any]] = []
    batch_id = f"batch_{uuid.uuid4().hex[:8]}"

    for sheet in payload.sheets:
        s_name = sheet.sheet_name
        headers = sheet.headers
        rows = sheet.rows

        # Dapatkan mapping untuk sheet ini
        s_mapping = sheet_mappings.get(s_name, {})
        # Hubungkan index kolom asli ke nama field kanonikal
        index_to_canonical: Dict[int, str] = {}
        for idx, h in enumerate(headers):
            canonical = s_mapping.get(h)
            if canonical:
                index_to_canonical[idx] = canonical

        for row in rows:
            record: Dict[str, Any] = {
                "_source_sheet": s_name
            }
            # Isi nilai untuk semua canonical field
            for c_field in canonical_fields:
                record[c_field] = None

            # Petakan nilai baris asli ke field kanonikal
            for idx, c_field in index_to_canonical.items():
                if idx < len(row):
                    record[c_field] = parse_cell_value(row[idx])

            all_consolidated_records.append(record)

    # Simpan ke database
    save_consolidated_records(batch_id, all_consolidated_records)

    return ConsolidatedDataset(
        status="success",
        total_sheets=len(payload.sheets),
        total_records=len(all_consolidated_records),
        canonical_fields=canonical_fields,
        sheet_mappings=sheet_mappings,
        records=all_consolidated_records
    )

@app.get("/api/consolidated-records")
def list_consolidated_records(limit: int = 100):
    """Melihat data hasil konsolidasi multi-sheet yang tersimpan di database."""
    return get_latest_consolidated_records(limit=limit)

# --- Endpoint Single Sheet Ingest ---
@app.post("/api/ingest", response_model=IngestResponse)
def ingest_sheet_data(payload: SheetIngestRequest):
    if not payload.headers or not payload.rows:
        raise HTTPException(status_code=400, detail="Headers and rows cannot be empty")

    column_mapping = map_single_sheet_headers(payload.headers, payload.rows[:2])

    header_indices = {}
    for idx, header in enumerate(payload.headers):
        field_name = column_mapping.get(header)
        if field_name:
            header_indices[field_name] = idx

    standard_items: List[StandardProductItem] = []

    for row in payload.rows:
        item_data: Dict[str, Any] = {"source": payload.sheet_name}
        for field_name, idx in header_indices.items():
            if idx < len(row):
                raw_value = row[idx]
                item_data[field_name] = parse_cell_value(raw_value) or 0.0

        item = StandardProductItem(**item_data)
        standard_items.append(item)

    saved_count = save_products(standard_items)

    return IngestResponse(
        status="success",
        total_rows_received=len(payload.rows),
        total_rows_saved=saved_count,
        column_mapping=column_mapping
    )

@app.get("/api/products")
def list_products(limit: int = 100):
    return get_all_products(limit=limit)

@app.get("/api/products/highest-eup")
def highest_eup_product():
    product = get_highest_eup()
    if not product:
        raise HTTPException(status_code=404, detail="No product data found")
    return product

@app.get("/api/analytics", response_model=ProductAnalytics)
def analytics():
    return get_analytics()
