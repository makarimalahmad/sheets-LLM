import os
import json
import re
import io
from typing import Dict, List, Any, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from app.mapper import get_llm_client, get_llm_model

# Helper membuat kode ID unik yang rapi dan ringkas
def generate_code_id(text: Any, prefix: str = "") -> str:
    s = str(text).strip()
    words = re.findall(r"[a-zA-Z0-9]+", s)
    if not words:
        return f"{prefix}_1" if prefix else "1"

    if len(words) == 1 and len(words[0]) <= 3:
        code = words[0].upper()
    else:
        acronym = "".join(w[0].upper() for w in words if not w.isdigit())
        nums = [w for w in words if w.isdigit()]
        if nums:
            code = f"{acronym}_{'_'.join(nums)}" if acronym else "_".join(nums)
        else:
            code = acronym if len(acronym) >= 2 else s[:3].upper()

    return f"{prefix}_{code}" if prefix and not code.startswith(prefix) else code

# Eksekusi pembuatan data tabel relasional oleh Python berdasarkan cetak biru skema dari LLM
def execute_normalization_engine(records: List[Dict[str, Any]], blueprint: Dict[str, Any]) -> Dict[str, Any]:
    tables_def = blueprint.get("tables", [])
    if not tables_def:
        raise ValueError("Cetak biru tabel kosong dari LLM")

    master_ids: Dict[str, Dict[str, str]] = {}
    child_ids: Dict[str, Dict[tuple, str]] = {}
    tables_output = []

    for t_def in tables_def:
        t_name = t_def["table_name"]
        t_type = t_def.get("table_type", "master")
        cols = t_def.get("columns", [])
        p_key = t_def.get("primary_key", "")
        f_keys = t_def.get("foreign_keys", [])
        src_col = t_def.get("source_column")

        table_rows = []

        if t_type == "master":
            master_ids[t_name] = {}
            seen = set()
            for row_item in records:
                val = row_item.get(src_col)
                if val is not None and val not in seen:
                    seen.add(val)
                    if len(cols) >= 2 and any("id" in str(c).lower() for c in cols):
                        uid = generate_code_id(val)
                        master_ids[t_name][val] = uid
                        row_vals = []
                        for c in cols:
                            if "id" in str(c).lower():
                                row_vals.append(uid)
                            else:
                                row_vals.append(val)
                        table_rows.append(row_vals)
                    else:
                        master_ids[t_name][val] = str(val)
                        table_rows.append([val])

        elif t_type == "child":
            child_ids[t_name] = {}
            parent_table = t_def.get("parent_table")
            parent_src_col = next(
                (td.get("source_column") for td in tables_def if td.get("table_name") == parent_table),
                None
            )

            seen = set()
            for row_item in records:
                parent_val = row_item.get(parent_src_col)
                child_val = row_item.get(src_col)
                parent_id = master_ids.get(parent_table, {}).get(parent_val, str(parent_val))

                combo_key = (parent_id, child_val)
                if combo_key not in seen and child_val is not None:
                    seen.add(combo_key)
                    child_uid = generate_code_id(child_val, prefix=str(parent_id))
                    child_ids[t_name][combo_key] = child_uid

                    row_vals = []
                    for c in cols:
                        if c == t_def.get("parent_fk_col") or (isinstance(c, str) and "parent" in c.lower()):
                            row_vals.append(parent_id)
                        elif c == t_def.get("primary_key") or (isinstance(c, str) and c.lower().endswith("id")):
                            row_vals.append(child_uid)
                        else:
                            row_vals.append(child_val)
                    table_rows.append(row_vals)

        elif t_type == "fact":
            for row_item in records:
                row_vals = []
                for c in cols:
                    matched = False
                    for fk in f_keys:
                        if fk.get("column") == c:
                            ref_tbl = fk.get("references_table")
                            if ref_tbl in child_ids:
                                p_tbl = next((td.get("parent_table") for td in tables_def if td.get("table_name") == ref_tbl), None)
                                p_src = next((td.get("source_column") for td in tables_def if td.get("table_name") == p_tbl), None)
                                c_src = next((td.get("source_column") for td in tables_def if td.get("table_name") == ref_tbl), None)
                                p_id = master_ids.get(p_tbl, {}).get(row_item.get(p_src))
                                c_id = child_ids.get(ref_tbl, {}).get((p_id, row_item.get(c_src)))
                                row_vals.append(c_id)
                                matched = True
                                break
                            elif ref_tbl in master_ids:
                                m_src = next((td.get("source_column") for td in tables_def if td.get("table_name") == ref_tbl), None)
                                m_id = master_ids.get(ref_tbl, {}).get(row_item.get(m_src))
                                row_vals.append(m_id)
                                matched = True
                                break
                    if not matched:
                        metric_val = None
                        for mc in t_def.get("metric_columns", []):
                            target_c = mc.get("target") or mc.get("target_column")
                            source_c = mc.get("source") or mc.get("source_column")
                            if target_c == c:
                                metric_val = row_item.get(source_c)
                                break
                        row_vals.append(metric_val if metric_val is not None else row_item.get(c))
                table_rows.append(row_vals)

        tables_output.append({
            "table_name": t_name,
            "columns": cols,
            "primary_key": p_key,
            "foreign_keys": f_keys,
            "rows": table_rows
        })

    return {
        "erd_summary": blueprint.get("erd_summary", "Relasi normalisasi berhasil dibentuk"),
        "tables": tables_output
    }

# Normalisasi data flat menjadi tabel-tabel relasional (Best Practice: LLM merancang arsitektur, Python mengeksekusi seluruh baris)
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

    # Eksekusi seluruh baris data asli (mau ribuan baris) menggunakan mesin Python
    try:
        return execute_normalization_engine(records, blueprint)
    except Exception:
        # Fallback jika bentuk blueprint berbeda dari standar
        if "tables" in blueprint and all("rows" in t for t in blueprint["tables"]):
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

    for table in tables_data:
        raw_name = str(table.get("table_name", "Sheet"))
        safe_sheet_name = raw_name[:31].replace(":", "").replace("/", "").replace("\\", "").replace("?", "").replace("*", "")
        worksheet = workbook.create_sheet(title=safe_sheet_name)

        columns = table.get("columns", [])
        rows = table.get("rows", [])

        # Tulis Header
        for col_index, col_name in enumerate(columns, start=1):
            cell = worksheet.cell(row=1, column=col_index, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        # Tulis Data
        for row_index, row_data in enumerate(rows, start=2):
            for col_index, val in enumerate(row_data, start=1):
                cell = worksheet.cell(row=row_index, column=col_index, value=val)
                cell.font = cell_font
                cell.border = thin_border
                if isinstance(val, (int, float)):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        # Sesuaikan lebar kolom otomatis
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    output_stream = io.BytesIO()
    workbook.save(output_stream)
    output_stream.seek(0)
    return output_stream
