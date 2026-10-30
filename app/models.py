from pydantic import BaseModel, Field
from typing import Optional, List, Dict, Any

class SingleSheetPayload(BaseModel):
    sheet_name: str = Field(default="Sheet1", description="Nama lembar / asal sheet")
    headers: List[str] = Field(description="Daftar nama kolom asli dari sheet")
    rows: List[List[Any]] = Field(default_factory=list, description="Baris data dari sheet")

class MultiSheetConsolidateRequest(BaseModel):
    sheets: List[SingleSheetPayload]

class ConsolidatedDataset(BaseModel):
    status: str
    total_sheets: int
    total_records: int
    canonical_fields: List[str]
    sheet_mappings: Dict[str, Dict[str, Optional[str]]]
    records: List[Dict[str, Any]]
