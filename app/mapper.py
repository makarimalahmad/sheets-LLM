import os
import json
from typing import Dict, List, Optional, Any
from dotenv import load_dotenv
from groq import Groq

load_dotenv()

# Ambil koneksi client Groq API
def get_groq_client() -> Optional[Groq]:
    api_key = os.getenv("GROQ_API_KEY", "").strip()
    if not api_key:
        return None
    return Groq(api_key=api_key)

# Minta LLM menyatukan nama-nama kolom dari berbagai sheet menjadi 1 format standar
def harmonize_sheets_schema(sheets: List[Dict[str, Any]]) -> Dict[str, Any]:
    client = get_groq_client()
    if not client:
        raise ValueError("GROQ_API_KEY belum disetel di file .env")

    # Ambil nama kolom dan 2 contoh baris isi data untuk dianalisis LLM
    simplified_sheets = []
    for s in sheets:
        simplified_sheets.append({
            "sheet_name": s.get("sheet_name", "Sheet"),
            "headers": s.get("headers", []),
            "sample_rows": s.get("rows", [])[:2]
        })

    prompt = f"""You are an automated schema harmonization engine.
You are given headers and sample row values from multiple different spreadsheets created by various users or systems.
Columns representing the exact same conceptual information across different sheets may have different wording, abbreviations, languages, or casing.

Your task:
1. Deduce the semantic meaning of each column purely from its header string and its sample cell values.
2. Unify columns across different sheets that represent the exact same concept under ONE single canonical field name (formatted in descriptive, clean snake_case).
3. STRICT CONSTRAINT - 1-to-1 Mapping Per Sheet: Two distinct columns within the SAME sheet must NEVER map to the same canonical field. If a sheet has two separate columns (such as end-user price excluding tax vs unit price, or payment rate vs payment cost), treat them as separate distinct attributes and give each its own unique canonical field name.
4. Provide a mapping from every raw header of each sheet to its corresponding canonical field name.

Input Sheets Data:
{json.dumps(simplified_sheets, indent=2)}

Output strictly a JSON object with this exact structure:
{{
  "canonical_fields": ["field_name_1", "field_name_2"],
  "sheet_mappings": {{
    "Sheet_Name": {{
      "original_header_name": "field_name_1"
    }}
  }}
}}"""

    response = client.chat.completions.create(
        model="openai/gpt-oss-120b",
        messages=[
            {"role": "system", "content": "You are a database architect AI. Output valid JSON only."},
            {"role": "user", "content": prompt}
        ],
        response_format={"type": "json_object"},
        temperature=0.0
    )

    content = response.choices[0].message.content
    data = json.loads(content)

    canonical_fields = data.get("canonical_fields", [])
    sheet_mappings = data.get("sheet_mappings", {})

    return {
        "canonical_fields": canonical_fields,
        "sheet_mappings": sheet_mappings
    }

# Pemetaan kolom untuk 1 sheet saja
def map_single_sheet_headers(headers: List[str], sample_rows: Optional[List[List[Any]]] = None) -> Dict[str, Optional[str]]:
    sheet_payload = [{
        "sheet_name": "InputSheet",
        "headers": headers,
        "rows": sample_rows or []
    }]
    harmonized = harmonize_sheets_schema(sheet_payload)
    mapping = harmonized.get("sheet_mappings", {}).get("InputSheet", {})
    return mapping
