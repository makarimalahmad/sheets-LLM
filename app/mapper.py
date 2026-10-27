import os
import json
from typing import Dict, List, Optional, Any
from dotenv import load_dotenv
from groq import Groq

load_dotenv()

def get_groq_client() -> Optional[Groq]:
    api_key = os.getenv("GROQ_API_KEY", "").strip()
    if not api_key:
        return None
    return Groq(api_key=api_key)

def harmonize_sheets_schema(sheets: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Zero-Shot Schema Harmonization via LLM (openai/gpt-oss-120b).
    Tanpa kamus kata kunci manual (Zero Hardcoded Dictionary).
    Menerima N sheet dengan nama kolom bebas/singkatan,
    LLM mengenali arti semantik & menghasilkan 1 skema kanonikal baku.
    """
    client = get_groq_client()
    if not client:
        raise ValueError("GROQ_API_KEY belum disetel di file .env")

    # Siapkan data ringkas (header + maksimal 2 baris sampel data untuk konteks nilai)
    simplified_sheets = []
    for s in sheets:
        simplified_sheets.append({
            "sheet_name": s.get("sheet_name", "Sheet"),
            "headers": s.get("headers", []),
            "sample_rows": s.get("rows", [])[:2]
        })

    prompt = f"""You are an automated schema harmonization engine.
You are given headers and sample row values from multiple different spreadsheets created by various users or systems.
Columns representing the exact same conceptual information may have different wording, abbreviations, languages, or casing.

Your task:
1. Deduce the semantic meaning of each column purely from its header string and its sample cell values.
2. Unify columns across all sheets that represent the exact same concept under ONE single canonical field name (formatted in descriptive, clean snake_case).
3. If a column represents a unique concept distinct from all others, assign it a clear canonical snake_case name.
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

def map_single_sheet_headers(headers: List[str], sample_rows: Optional[List[List[Any]]] = None) -> Dict[str, Optional[str]]:
    """
    Fungsi pembantu untuk memetakan kolom dari 1 sheet saja tanpa kamus manual.
    """
    sheet_payload = [{
        "sheet_name": "InputSheet",
        "headers": headers,
        "rows": sample_rows or []
    }]
    harmonized = harmonize_sheets_schema(sheet_payload)
    mapping = harmonized.get("sheet_mappings", {}).get("InputSheet", {})
    return mapping
