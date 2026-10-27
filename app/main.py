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
    get_latest_consolidated_records,
    clear_all_records
)

app = FastAPI(
    title="Sheets LLM Backend API",
    description="Otomatisasi pembacaan Google Sheets, standarisasi nama kolom, dan penggabungan data spreadsheet.",
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

# Preview pemetaan kolom standar dari beberapa sheet
@app.post("/api/harmonize-schema", response_model=HarmonizationResult)
def preview_harmonization(payload: MultiSheetConsolidateRequest):
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
        raise HTTPException(status_code=500, detail=f"Gagal memproses nama kolom: {str(e)}")

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

# Ambil riwayat data gabungan dari database
@app.get("/api/consolidated-records")
def list_consolidated_records(limit: int = 100):
    return get_latest_consolidated_records(limit=limit)

# Hapus semua riwayat data di database
@app.delete("/api/clear-data")
def clear_data():
    clear_all_records()
    return {"status": "success", "message": "Riwayat data berhasil dibersihkan"}

# Terima dan simpan data sheet tunggal
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
                val = parse_cell_value(raw_value) or 0.0
                fn = field_name.lower()
                if "incl" in fn or ("vat" in fn and "excl" not in fn and "without" not in fn and "tanpa" not in fn) or "ppn" in fn:
                    item_data["eup_inc_vat"] = val
                elif "excl" in fn or "without" in fn or "tanpa" in fn:
                    item_data["eup_ex_vat"] = val
                elif "goods" in fn or "cogs" in fn or "hpp" in fn or "modal" in fn:
                    item_data["cogs"] = val
                elif "margin" in fn or "profit" in fn or "untung" in fn or "laba" in fn:
                    item_data["margin"] = val
                elif "rate" in fn or "persen" in fn:
                    item_data["payment_rate"] = val
                elif "cost" in fn or "fee" in fn or "biaya" in fn or "ongkos" in fn:
                    item_data["payment_cost"] = val
                elif "unit" in fn or "satuan" in fn or "price" in fn:
                    item_data["unit_price"] = val
                else:
                    item_data[field_name] = val

        item = StandardProductItem(**item_data)
        standard_items.append(item)

    saved_count = save_products(standard_items)

    return IngestResponse(
        status="success",
        total_rows_received=len(payload.rows),
        total_rows_saved=saved_count,
        column_mapping=column_mapping
    )

# Ambil daftar produk yang tersimpan
@app.get("/api/products")
def list_products(limit: int = 100):
    return get_all_products(limit=limit)

# Ambil produk dengan harga EUP termahal
@app.get("/api/products/highest-eup")
def highest_eup_product():
    product = get_highest_eup()
    if not product:
        raise HTTPException(status_code=404, detail="No product data found")
    return product

# Hitung ringkasan statistik produk untuk dashboard
@app.get("/api/analytics", response_model=ProductAnalytics)
def analytics():
    return get_analytics()
