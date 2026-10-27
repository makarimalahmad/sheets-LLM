from pydantic import BaseModel, Field
from typing import Optional, List, Dict, Any

class SingleSheetPayload(BaseModel):
    sheet_name: str = Field(default="Sheet1", description="Nama lembar / asal sheet")
    headers: List[str] = Field(description="Daftar nama kolom asli dari sheet")
    rows: List[List[Any]] = Field(default_factory=list, description="Baris data dari sheet")

class MultiSheetConsolidateRequest(BaseModel):
    sheets: List[SingleSheetPayload]

class HarmonizationResult(BaseModel):
    canonical_fields: List[str]
    sheet_mappings: Dict[str, Dict[str, Optional[str]]]

class ConsolidatedDataset(BaseModel):
    status: str
    total_sheets: int
    total_records: int
    canonical_fields: List[str]
    sheet_mappings: Dict[str, Dict[str, Optional[str]]]
    records: List[Dict[str, Any]]

class StandardProductItem(BaseModel):
    id: Optional[int] = None
    source: Optional[str] = "google-sheets"
    eup_inc_vat: Optional[float] = Field(default=0.0, description="End user price include VAT")
    eup_ex_vat: Optional[float] = Field(default=0.0, description="End user price exclude VAT")
    payment_rate: Optional[float] = Field(default=0.0, description="Payment rate")
    payment_cost: Optional[float] = Field(default=0.0, description="Payment cost")
    unit_price: Optional[float] = Field(default=0.0, description="Unit price")
    cogs: Optional[float] = Field(default=0.0, description="Cost of goods sold")
    margin: Optional[float] = Field(default=0.0, description="Margin")

# Model untuk kompatibilitas single sheet
class SheetIngestRequest(BaseModel):
    sheet_name: Optional[str] = "Sheet1"
    headers: List[str]
    rows: List[List[Any]]

class IngestResponse(BaseModel):
    status: str
    total_rows_received: int
    total_rows_saved: int
    column_mapping: Dict[str, Optional[str]]

class ProductAnalytics(BaseModel):
    total_products: int
    highest_eup_inc_vat: Optional[float]
    lowest_eup_inc_vat: Optional[float]
    average_margin: Optional[float]
    top_product_by_eup: Optional[Dict[str, Any]]
