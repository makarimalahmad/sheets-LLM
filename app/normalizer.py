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

    # Ambil maksimal 50 baris untuk dianalisis LLM agar hemat token namun representatif
    sample_records = records[:50]

    prompt = f"""Kamu adalah AI arsitek database untuk normalisasi tabel spreadsheet menjadi skema relasional (1NF, 2NF, 3NF / ERD).
Diberikan baris data denormalisasi (flat table):
{json.dumps(sample_records, indent=2)}

Tugas kamu:
1. Analisis entitas independen (master/dimensi) dan tabel relasi/transaksi (fakta/junction).
2. Normalisasikan data di atas menjadi tabel-tabel relasional terpisah (misalnya tabel master judul game, tabel denominasi dengan relasi ke game, tabel channel penjualan, dan tabel relasi distribusi harga/fakta).
3. Buat ID atau kode unik yang ringkas, deskriptif, dan konsisten (misal 'Free Fire' -> 'FF', '10 diamond' -> 'FF_10_D').
4. Hasilkan baris data unik dan bersih untuk setiap tabel tanpa ada baris duplikat.
5. Hubungkan tabel-tabel tersebut menggunakan Foreign Key ID yang telah dibuat.

Keluarkan format HANYA objek JSON dengan struktur persis seperti ini:
{{
  "erd_summary": "Deskripsi alur relasi tabel (contoh: Game Title -> Game Denom -> Game Dist Channel <- Channel)",
  "tables": [
    {{
      "table_name": "Nama Tabel",
      "columns": ["Kolom1", "Kolom2"],
      "primary_key": "Nama Kolom PK",
      "foreign_keys": [
        {{
          "column": "Nama Kolom FK",
          "references_table": "Tabel Tujuan",
          "references_column": "Kolom Tujuan"
        }}
      ],
      "rows": [
        ["Nilai1", "Nilai2"]
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
