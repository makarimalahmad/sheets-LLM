import os
import re
from typing import Any, Dict, List
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from app.models import (
    SheetIngestRequest,
    IngestResponse,
    MappingResponse,
    StandardProductItem,
    ProductAnalytics
)
from app.mapper import map_headers
from app.database import (
    save_products,
    get_all_products,
    get_highest_eup,
    get_analytics
)

app = FastAPI(
    title="Sheets LLM Backend API",
    description="API otomatisasi pembaca Google Sheets, pemetaan nama kolom, dan analitik harga.",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def parse_number(val: Any) -> float:
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return 0.0
    s = re.sub(r"[^\d.,\-]", "", s)
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0

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

@app.post("/api/preview-mapping", response_model=MappingResponse)
def preview_mapping(headers: List[str]):
    mapping = map_headers(headers)
    unmapped = [h for h, field in mapping.items() if field is None]
    return MappingResponse(
        mapped_columns=mapping,
        unmapped_columns=unmapped
    )

@app.post("/api/ingest", response_model=IngestResponse)
def ingest_sheet_data(payload: SheetIngestRequest):
    if not payload.headers or not payload.rows:
        raise HTTPException(status_code=400, detail="Headers and rows cannot be empty")

    column_mapping = map_headers(payload.headers)
    
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
                item_data[field_name] = parse_number(raw_value)
            else:
                item_data[field_name] = 0.0

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
