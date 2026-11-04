from pydantic import BaseModel, Field
from typing import Optional, List, Dict, Any

class SingleSheetPayload(BaseModel):
    sheet_name: str = Field(default="Sheet1", description="Nama lembar / asal sheet")
    headers: List[str] = Field(description="Daftar nama kolom asli dari sheet")
    rows: List[List[Any]] = Field(default_factory=list, description="Baris data dari sheet")

class MultiSheetConsolidateRequest(BaseModel):
    sheets: List[SingleSheetPayload]

class PullGoogleSheetsRequest(BaseModel):
    sheet_url: str = Field(description="URL Google Sheets publik/viewable")
    tab_names: Optional[List[str]] = Field(default=None, description="Daftar tab yang ingin ditarik (opsional)")

class ConsolidatedDataset(BaseModel):
    status: str
    total_sheets: int
    total_records: int
    canonical_fields: List[str]
    sheet_mappings: Dict[str, Dict[str, Optional[str]]]
    records: List[Dict[str, Any]]

class NormalizedTable(BaseModel):
    table_name: str
    columns: List[str]
    primary_key: Optional[str] = None
    foreign_keys: Optional[List[Dict[str, Any]]] = None
    rows: List[List[Any]]

class NormalizedResponse(BaseModel):
    status: str = "success"
    erd_summary: str
    tables: List[NormalizedTable]

class NormalizeRequest(BaseModel):
    records: Optional[List[Dict[str, Any]]] = Field(default=None, description="Daftar record flat untuk dinormalisasi")
    sheet_url: Optional[str] = Field(default=None, description="Atau link Google Sheets yang ingin dinormalisasi langsung")
