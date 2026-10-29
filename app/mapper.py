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

    prompt = f"""Kamu adalah sistem AI untuk standarisasi nama kolom data spreadsheet.
Kamu diberikan daftar nama kolom dan contoh nilai baris dari beberapa sheet berbeda.
Setiap sheet bisa menggunakan kata, singkatan, bahasa, atau penulisan yang berbeda untuk informasi yang sebenarnya sama.

Tugas kamu:
1. Pahami arti setiap kolom dari teks nama kolom dan contoh isi nilainya.
2. Satukan kolom dari berbagai sheet yang memiliki makna sama ke dalam SATU nama kolom standar (ditulis dalam format snake_case bahasa Inggris, contoh: end_user_price_incl_vat, payment_rate, unit_price, cost_of_goods_sold, profit_margin).
3. ATURAN KETAT - 1-ke-1 per Sheet: Dua kolom yang berbeda dalam SATU sheet yang sama TIDAK BOLEH disamakan ke satu nama standar yang sama. Jika ada sheet yang memiliki dua kolom terpisah (misalnya harga sebelum pajak dan harga per unit, atau rate pembayaran dan biaya pembayaran), perlakukan sebagai dua hal terpisah dan berikan masing-masing nama standar yang berbeda.
4. Buat pemetaan dari setiap nama kolom asli ke nama kolom standar yang sesuai.

Data Sheet:
{json.dumps(simplified_sheets, indent=2)}

Keluarkan format HANYA objek JSON dengan struktur persis seperti ini:
{{
  "standard_columns": ["nama_standar_1", "nama_standar_2"],
  "sheet_mappings": {{
    "Nama_Sheet": {{
      "nama_kolom_asli": "nama_standar_1"
    }}
  }}
}}"""

    response = client.chat.completions.create(
        model="openai/gpt-oss-120b",
        messages=[
            {"role": "system", "content": "Kamu adalah sistem standarisasi data. Hanya balas dalam format JSON yang valid."},
            {"role": "user", "content": prompt}
        ],
        response_format={"type": "json_object"},
        temperature=0.0
    )

    content = response.choices[0].message.content
    data = json.loads(content)

    standard_columns = data.get("standard_columns") or data.get("canonical_fields") or []
    sheet_mappings = data.get("sheet_mappings", {})

    return {
        "canonical_fields": standard_columns,
        "standard_columns": standard_columns,
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
