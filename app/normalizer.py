import os
import json
import re
import io
from typing import Dict, List, Any, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from app.mapper import get_llm_client, get_llm_model

# Cek apakah dataset berupa katalog item (Content, Item, Description, Supplier, ID)
def is_catalog_dataset(sample_record: Dict[str, Any]) -> bool:
    keys_lower = [str(k).lower().strip() for k in sample_record.keys() if not str(k).startswith("_")]
    has_item = any("item" in k for k in keys_lower)
    has_supplier = any("supplier" in k or "vendor" in k for k in keys_lower)
    return has_item and has_supplier

# Ambil angka nominal pecahan item (contoh: "freefire_140" -> "140", "mkios_5_diamonds" -> "5")
def get_item_nominal(item_text: Any) -> str:
    numbers = re.findall(r"\d+", str(item_text or ""))
    if numbers:
        return numbers[0]
    return ""

# Pilih baris item terbaik dari kumpulan kandidat varian yang sama
def choose_best_catalog_item(candidate_rows: List[Dict[str, Any]], game_name: str, item_key: str, desc_key: str) -> Dict[str, Any]:
    best_item = candidate_rows[0]

    for current_row in candidate_rows:
        current_desc = str(current_row.get(desc_key, "") or "").strip()
        best_desc = str(best_item.get(desc_key, "") or "").strip()

        current_name = str(current_row.get(item_key, "") or "").lower().strip()
        best_name = str(best_item.get(item_key, "") or "").lower().strip()

        # 1. Prioritaskan baris yang memiliki deskripsi lengkap (bukan kosong/tanda strip)
        has_current_desc = bool(current_desc and current_desc != "-")
        has_best_desc = bool(best_desc and best_desc != "-")

        if has_current_desc and not has_best_desc:
            best_item = current_row
            continue
        elif not has_current_desc and has_best_desc:
            continue

        # 2. Prioritaskan nama item yang mengandung nama game (contoh: "freefire_5" lebih utama dari "mkios_5")
        clean_game = game_name.replace("_", "").lower()
        current_has_game = clean_game in current_name.replace("_", "")
        best_has_game = clean_game in best_name.replace("_", "")

        if current_has_game and not best_has_game:
            best_item = current_row
            continue
        elif not current_has_game and best_has_game:
            continue

        # 3. Pilih kode item yang lebih pendek dan bersih
        if len(current_name) < len(best_name):
            best_item = current_row

    return best_item

# Deduplikasi varian item per supplier untuk seluruh baris data (Opsi A)
def deduplicate_catalog_records(records: List[Dict[str, Any]]):
    from collections import defaultdict

    original_cols = [str(k) for k in records[0].keys() if not str(k).startswith("_")]

    # Standarisasi nama kolom header agar selalu rapi: ID, Content, Item, Description, Supplier
    display_cols = []
    for col_name in original_cols:
        col_lower = col_name.lower().replace("_", " ").strip()
        if "id" in col_lower and ("tim" in col_lower or "team" in col_lower or col_lower == "id"):
            display_cols.append("ID")
        elif "content" in col_lower or "game" in col_lower:
            display_cols.append("Content")
        elif "item" in col_lower:
            display_cols.append("Item")
        elif "desc" in col_lower:
            display_cols.append("Description")
        elif "supplier" in col_lower or "vendor" in col_lower:
            display_cols.append("Supplier")
        else:
            display_cols.append(col_name)

    item_key = next((k for k in original_cols if "item" in str(k).lower().strip()), original_cols[0])
    supplier_key = next((k for k in original_cols if "supplier" in str(k).lower().strip() or "vendor" in str(k).lower().strip()), None)
    content_key = next((k for k in original_cols if "content" in str(k).lower().strip() or "game" in str(k).lower().strip()), None)
    desc_key = next((k for k in original_cols if "description" in str(k).lower().strip() or "desc" in str(k).lower().strip()), None)

    # Kelompokkan per (Supplier, Content) agar supplier dan game berbeda tidak pernah tercampur
    by_supp_content = defaultdict(list)
    for row_item in records:
        supp_val = str(row_item.get(supplier_key, "") or "").strip().lower() if supplier_key else ""
        cont_val = str(row_item.get(content_key, "") or "").strip().lower() if content_key else ""
        by_supp_content[(supp_val, cont_val)].append(row_item)

    cleaned_rows = []

    for (supp_val, cont_val), group_rows in by_supp_content.items():
        # Kelompokkan item berdasarkan angka nominal pecahan
        nominal_groups = defaultdict(list)

        for row in group_rows:
            item_text = str(row.get(item_key, "") or "").lower().strip()
            nominal = get_item_nominal(item_text)
            if nominal:
                nominal_groups[nominal].append(row)
            else:
                nominal_groups[item_text].append(row)

        # Pilih 1 baris terbaik yang bersih untuk setiap nominal pecahan
        for nominal_key, candidate_rows in nominal_groups.items():
            best_row = choose_best_catalog_item(candidate_rows, cont_val, item_key, desc_key or "")
            cleaned_rows.append([best_row.get(col_name) for col_name in original_cols])

    return display_cols, cleaned_rows

# Normalisasi data flat menjadi tabel terstruktur via LLM
def normalize_dataset(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    client = get_llm_client()
    if not client:
        raise ValueError("Variabel LLM_API_KEY belum disetel di file .env")

    model_name = get_llm_model()
    sample_records = records[:50]

    # KASUS 1: Dataset Katalog Varian Item
    if records and is_catalog_dataset(records[0]):
        original_cols = [str(k) for k in records[0].keys() if not str(k).startswith("_")]

        prompt = f"""Kamu adalah AI pembersih dan penormalisasi data katalog item spreadsheet.
Diberikan baris data denormalisasi (banyak duplikat varian dan suffix promo):
{json.dumps(sample_records, indent=2)}

Tugas kamu (Normalisasi Varian Katalog):
1. Hilangkan redundansi dan duplikasi varian dalam supplier yang sama:
   - Jika ada beberapa baris dengan Content, Description, dan Supplier yang sama (misal 'AOV 1430 Vouchers' di supplier 'garena' yang memiliki banyak kode varian seperti 'aov_1430', 'aov_1430_dg_gpy', 'aov_1430_tact'), satukan menjadi SATU baris bersih menggunakan kode Item dasar tanpa suffix/ekor promo tambahan (contoh: 'aov_1430').
   - Gunakan ID yang sesuai dengan baris item dasar tersebut.
2. Pembeda antar-supplier:
   - Jika item yang sama dijual oleh Supplier berbeda (misal 'garena' dan 'coda'), TETAP catat sebagai baris terpisah untuk masing-masing supplier.
3. Hasilkan ringkasan hasil normalisasi.

Keluarkan format HANYA objek JSON valid:
{{
  "erd_summary": "Deduplikasi varian item per supplier: varian dengan suffix disatukan ke item dasar, supplier berbeda tetap dipisah.",
  "clean_table_name": "Tabel Bersih"
}}"""

        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": "Kamu adalah analis data katalog. Hanya balas dalam format JSON yang valid."},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"},
            temperature=0.0
        )
        content_data = json.loads(response.choices[0].message.content)
        summary_text = content_data.get("erd_summary", "Deduplikasi varian item per supplier berhasil.")

        # Eksekusi pembersihan menyeluruh untuk seluruh baris data
        clean_columns, cleaned_table_rows = deduplicate_catalog_records(records)

        return {
            "erd_summary": summary_text,
            "tables": [
                {
                    "table_name": "Tabel Bersih",
                    "columns": clean_columns,
                    "primary_key": next((k for k in clean_columns if str(k).lower().strip() == "id"), None),
                    "foreign_keys": [],
                    "rows": cleaned_table_rows
                }
            ]
        }

    # KASUS 2: Dataset Game, Denom, Channel, Harga (Normalisasi ERD 4 Tabel)
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
