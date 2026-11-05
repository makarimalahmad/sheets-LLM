import os
import json
import re
import io
from typing import Dict, List, Any, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from app.mapper import get_llm_client, get_llm_model

# Fungsi pembantu untuk membuat kode ID unik yang rapi dan ringkas
def generate_code_id(text_input: Any, prefix_code: str = "") -> str:
    raw_text = str(text_input).strip()
    words = re.findall(r"[a-zA-Z0-9]+", raw_text)
    if not words:
        return f"{prefix_code}_1" if prefix_code else "1"

    if len(words) == 1 and len(words[0]) <= 3:
        code_result = words[0].upper()
    else:
        acronym = "".join(word[0].upper() for word in words if not word.isdigit())
        numbers = [word for word in words if word.isdigit()]
        if numbers:
            code_result = f"{acronym}_{'_'.join(numbers)}" if acronym else "_".join(numbers)
        else:
            code_result = acronym if len(acronym) >= 2 else raw_text[:3].upper()

    if prefix_code and not code_result.startswith(prefix_code):
        return f"{prefix_code}_{code_result}"
    return code_result

# Eksekusi pembuatan data tabel relasional oleh Python berdasarkan cetak biru skema dari LLM
def execute_normalization_engine(records: List[Dict[str, Any]], blueprint: Dict[str, Any]) -> Dict[str, Any]:
    tables_definition = blueprint.get("tables", [])
    if not tables_definition:
        raise ValueError("Cetak biru tabel kosong dari LLM")

    master_ids: Dict[str, Dict[str, str]] = {}
    child_ids: Dict[str, Dict[tuple, str]] = {}
    tables_output = []

    for table_def in tables_definition:
        table_name = table_def["table_name"]
        table_type = table_def.get("table_type", "master")
        columns = table_def.get("columns", [])
        primary_key = table_def.get("primary_key", "")
        foreign_keys = table_def.get("foreign_keys", [])
        source_column = table_def.get("source_column")

        table_rows = []

        if table_type == "master":
            master_ids[table_name] = {}
            seen_values = set()
            for row_item in records:
                cell_value = row_item.get(source_column)
                if cell_value is not None and cell_value not in seen_values:
                    seen_values.add(cell_value)
                    if len(columns) >= 2 and any("id" in str(col_name).lower() for col_name in columns):
                        unique_id = generate_code_id(cell_value)
                        master_ids[table_name][cell_value] = unique_id
                        row_values = []
                        for col_name in columns:
                            if "id" in str(col_name).lower():
                                row_values.append(unique_id)
                            else:
                                row_values.append(cell_value)
                        table_rows.append(row_values)
                    else:
                        master_ids[table_name][cell_value] = str(cell_value)
                        table_rows.append([cell_value])

        elif table_type == "child":
            child_ids[table_name] = {}
            parent_table = table_def.get("parent_table")
            parent_source_col = next(
                (target_def.get("source_column") for target_def in tables_definition if target_def.get("table_name") == parent_table),
                None
            )

            seen_combinations = set()
            for row_item in records:
                parent_value = row_item.get(parent_source_col)
                child_value = row_item.get(source_column)
                parent_id = master_ids.get(parent_table, {}).get(parent_value, str(parent_value))

                combo_key = (parent_id, child_value)
                if combo_key not in seen_combinations and child_value is not None:
                    seen_combinations.add(combo_key)
                    child_id = generate_code_id(child_value, prefix_code=str(parent_id))
                    child_ids[table_name][combo_key] = child_id

                    row_values = []
                    for col_name in columns:
                        if col_name == table_def.get("parent_fk_col") or (isinstance(col_name, str) and "parent" in col_name.lower()):
                            row_values.append(parent_id)
                        elif col_name == table_def.get("primary_key") or (isinstance(col_name, str) and col_name.lower().endswith("id")):
                            row_values.append(child_id)
                        else:
                            row_values.append(child_value)
                    table_rows.append(row_values)

        elif table_type == "fact":
            for row_item in records:
                row_values = []
                for col_name in columns:
                    matched = False
                    for foreign_key in foreign_keys:
                        if foreign_key.get("column") == col_name:
                            ref_table = foreign_key.get("references_table")
                            if ref_table in child_ids:
                                p_table = next((target_def.get("parent_table") for target_def in tables_definition if target_def.get("table_name") == ref_table), None)
                                p_source = next((target_def.get("source_column") for target_def in tables_definition if target_def.get("table_name") == p_table), None)
                                c_source = next((target_def.get("source_column") for target_def in tables_definition if target_def.get("table_name") == ref_table), None)
                                parent_pk = master_ids.get(p_table, {}).get(row_item.get(p_source))
                                child_fk = child_ids.get(ref_table, {}).get((parent_pk, row_item.get(c_source)))
                                row_values.append(child_fk)
                                matched = True
                                break
                            elif ref_table in master_ids:
                                m_source = next((target_def.get("source_column") for target_def in tables_definition if target_def.get("table_name") == ref_table), None)
                                master_fk = master_ids.get(ref_table, {}).get(row_item.get(m_source))
                                row_values.append(master_fk)
                                matched = True
                                break
                    if not matched:
                        metric_value = None
                        for metric in table_def.get("metric_columns", []):
                            target_col = metric.get("target") or metric.get("target_column")
                            source_col = metric.get("source") or metric.get("source_column")
                            if target_col == col_name:
                                metric_value = row_item.get(source_col)
                                break
                        row_values.append(metric_value if metric_value is not None else row_item.get(col_name))
                table_rows.append(row_values)

        tables_output.append({
            "table_name": table_name,
            "columns": columns,
            "primary_key": primary_key,
            "foreign_keys": foreign_keys,
            "rows": table_rows
        })

    return {
        "erd_summary": blueprint.get("erd_summary", "Relasi normalisasi berhasil dibentuk"),
        "tables": tables_output
    }

# Normalisasi data flat menjadi tabel-tabel relasional (LLM merancang arsitektur, Python mengeksekusi seluruh baris)
def normalize_dataset(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    client = get_llm_client()
    if not client:
        raise ValueError("Variabel LLM_API_KEY belum disetel di file .env")

    model_name = get_llm_model()

    # Ambil maksimal 4 baris sampel agar AI fokus memahami pola hubungan tanpa boros token
    sample_records = records[:4]

    prompt = f"""Kamu adalah AI arsitek database untuk normalisasi tabel flat menjadi skema relasional (1NF, 2NF, 3NF / ERD).
Diberikan contoh kolom dan baris data:
{json.dumps(sample_records, indent=2)}

Tugas kamu:
Rancang struktur tabel relasional (tanpa perlu mengetik semua baris data).
Identifikasi:
1. Tabel master / dimensi independen (misal: Game, Channel).
2. Tabel turunan bertingkat jika ada (misal: Denomination yang bergantung pada Game).
3. Tabel fakta / transaksi penghubung (misal: Harga per Channel & Denom).
4. Kolom sumber asli yang dipetakan ke tiap tabel.

Keluarkan format HANYA JSON valid dengan struktur:
{{
  "erd_summary": "Deskripsi alur relasi (misal Game Title -> Game Denom -> Game Dist Channel <- Channel)",
  "tables": [
    {{
      "table_name": "Game Title",
      "table_type": "master",
      "primary_key": "ID",
      "source_column": "kolom_sumber_di_data",
      "columns": ["Title", "ID"]
    }},
    {{
      "table_name": "Game Denom",
      "table_type": "child",
      "parent_table": "Game Title",
      "parent_fk_col": "ID Title",
      "primary_key": "ID DENOM",
      "source_column": "kolom_sumber_di_data",
      "columns": ["ID Title", "Denomination", "ID DENOM"],
      "foreign_keys": [{{"column": "ID Title", "references_table": "Game Title", "references_column": "ID"}}]
    }},
    {{
      "table_name": "Channel",
      "table_type": "master",
      "primary_key": "Channel",
      "source_column": "kolom_sumber_di_data",
      "columns": ["Channel"]
    }},
    {{
      "table_name": "Game Dist Channel",
      "table_type": "fact",
      "primary_key": "ID DENOM + Channel",
      "columns": ["ID DENOM", "Channel", "Harga"],
      "foreign_keys": [
        {{"column": "ID DENOM", "references_table": "Game Denom", "references_column": "ID DENOM"}},
        {{"column": "Channel", "references_table": "Channel", "references_column": "Channel"}}
      ],
      "metric_columns": [{{"source": "kolom_harga_sumber", "target": "Harga"}}]
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
    blueprint = json.loads(content)

    # Eksekusi seluruh baris data asli menggunakan mesin Python
    try:
        return execute_normalization_engine(records, blueprint)
    except Exception:
        if "tables" in blueprint and all("rows" in target_table for target_table in blueprint["tables"]):
            return blueprint
        raise

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
