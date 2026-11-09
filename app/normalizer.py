import os
import json
import io
from typing import Dict, List, Any, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from app.mapper import get_llm_client, get_llm_model

# Normalisasi data flat menjadi tabel-tabel relasional murni via LLM
def normalize_dataset(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    client = get_llm_client()
    if not client:
        raise ValueError("Variabel LLM_API_KEY belum disetel di file .env")

    model_name = get_llm_model()

    # Ambil maksimal 50 baris untuk dianalisis LLM agar hemat token 
    sample_records = records[:50]

    prompt = f"""Kamu adalah AI arsitek database untuk normalisasi tabel spreadsheet menjadi skema relasional (1NF, 2NF, 3NF / ERD).
Diberikan baris data denormalisasi (flat table):
{json.dumps(sample_records, indent=2)}

Tugas kamu:
Pecah data denormalisasi di atas menjadi 4 tabel relasional terstruktur:
1. Tabel Game Title (kolom: Title, ID) -> ID berupa kode akronim huruf besar (misal 'Free Fire' -> 'FF', 'Mobile Legends' -> 'ML', 'Roblox' -> 'RB').
2. Tabel Game Denom (kolom: ID Title, Denomination, ID DENOM) -> ID Title adalah FK ke Game Title. Nilai Denomination harus sesuai mata uang game aslinya (misal '10 diamond', '80 robux'). ID DENOM adalah PK gabungan ID Title, nominal angka, dan singkatan mata uang (gunakan 'DM' untuk diamond seperti 'FF_10_DM', 'ML_50_DM'; dan 'RBX' untuk robux seperti 'RB_80_RBX').
3. Tabel Channel (kolom: Channel) -> daftar nama channel/toko unik (misal: Upoint, Dunia Games, Itemku).
4. Tabel Game Dist Channel (kolom: ID DENOM, Channel, Harga) -> tabel relasi harga dengan Foreign Key 'ID DENOM' dan Foreign Key 'Channel'.

Keluarkan format HANYA objek JSON dengan struktur persis seperti ini:
{{
  "erd_summary": "Game Title (ID) -> Game Denom (ID DENOM) -> Game Dist Channel <- Channel (Channel)",
  "tables": [
    {{
      "table_name": "Game Title",
      "columns": ["Title", "ID"],
      "primary_key": "ID",
      "foreign_keys": [],
      "rows": [
        ["Free Fire", "FF"]
      ]
    }},
    {{
      "table_name": "Game Denom",
      "columns": ["ID Title", "Denomination", "ID DENOM"],
      "primary_key": "ID DENOM",
      "foreign_keys": [
        {{
          "column": "ID Title",
          "references_table": "Game Title",
          "references_column": "ID"
        }}
      ],
      "rows": [
        ["FF", "10 diamond", "FF_10_DM"]
      ]
    }},
    {{
      "table_name": "Channel",
      "columns": ["Channel"],
      "primary_key": "Channel",
      "foreign_keys": [],
      "rows": [
        ["Upoint"]
      ]
    }},
    {{
      "table_name": "Game Dist Channel",
      "columns": ["ID DENOM", "Channel", "Harga"],
      "primary_key": "ID DENOM + Channel",
      "foreign_keys": [
        {{
          "column": "ID DENOM",
          "references_table": "Game Denom",
          "references_column": "ID DENOM"
        }},
        {{
          "column": "Channel",
          "references_table": "Channel",
          "references_column": "Channel"
        }}
      ],
      "rows": [
        ["FF_10_DM", "Upoint", 12000]
      ]
    }}
  ]
}}"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": "Kamu adalah arsitek normalisasi database. Hanya balas dalam format JSON yang valid."},
            {"role": "user", "content": prompt}
        ],
        response_format={"type": "json_object"},
        temperature=0.0
    )

    content = response.choices[0].message.content
    result = json.loads(content)
    return result

# Buat file Excel multi-tab dari hasil normalisasi
def export_normalized_to_excel_buffer(tables_data: List[Dict[str, Any]]) -> io.BytesIO:
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    cell_font = Font(name="Calibri", size=11)
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    for table_item in tables_data:
        raw_name = str(table_item.get("table_name", "Sheet"))
        safe_sheet_name = raw_name[:31].replace(":", "").replace("/", "").replace("\\", "").replace("?", "").replace("*", "")
        worksheet = workbook.create_sheet(title=safe_sheet_name)

        columns = table_item.get("columns", [])
        rows = table_item.get("rows", [])

        # Tulis Header
        for col_index, col_name in enumerate(columns, start=1):
            cell = worksheet.cell(row=1, column=col_index, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        # Tulis Data
        for row_index, row_data in enumerate(rows, start=2):
            for col_index, cell_value in enumerate(row_data, start=1):
                cell = worksheet.cell(row=row_index, column=col_index, value=cell_value)
                cell.font = cell_font
                cell.border = thin_border
                if isinstance(cell_value, (int, float)):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        # Sesuaikan lebar kolom otomatis
        for column_cells in worksheet.columns:
            max_length = max(len(str(cell.value or "")) for cell in column_cells)
            col_letter = column_cells[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_length + 4, 12)

    output_stream = io.BytesIO()
    workbook.save(output_stream)
    output_stream.seek(0)
    return output_stream
