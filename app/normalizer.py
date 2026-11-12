import os
import json
import re
import io
from typing import Dict, List, Any, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from app.mapper import get_llm_client, get_llm_model

# 1. Cek apakah item adalah data uji coba / sampah (contoh: rom_test, laplace_test, cod_test)
def is_test_item(item_text: Any, desc_text: Any) -> bool:
    item = str(item_text or "").lower().strip()
    desc = str(desc_text or "").lower().strip()
    if "test" in item or "test" in desc:
        if desc == "" or desc == "-" or desc == "none" or "test" in desc:
            return True
    return False

# 2. Ekstraksi identitas produk/denominasi kanonikal secara universal
def get_product_key(item_text: Any, desc_text: Any) -> str:
    item = str(item_text or "").lower().strip()
    desc = str(desc_text or "").lower().strip()

    # Cek pass / membership / paket khusus
    special_passes = [
        "twilight", "weekly", "starlight", "monthly epic", "weekly elite", 
        "diamond pass", "aurum", "companionship", "supply pass", "grow pass", 
        "abundance", "bounty", "fountain", "rainin", "chest o", "battle pass", 
        "peluru", "komposit", "pemula"
    ]
    for pattern in special_passes:
        if pattern in desc or pattern in item:
            return pattern

    # Ambil angka nominal dari Description jika ada (Description adalah nama produk resmi)
    if desc and desc not in ["-", "none", ""]:
        d_clean = re.sub(r"\(.*?\)", "", desc)
        # Cek jika ada format penjumlahan (misal: 1580 + 110) yang hasilnya cocok dengan nama Item
        plus_match = re.search(r"(\d+)\s*\+\s*(\d+)", d_clean)
        if plus_match:
            total_sum = int(plus_match.group(1)) + int(plus_match.group(2))
            if str(total_sum) in item:
                return f"denom_{total_sum}"

        # Buang tanda + dan semua teks bonus promo di belakangnya agar melebur ke base item
        d_clean = re.sub(r"\+.*", "", d_clean)
        d_clean = re.sub(r"\b(funpack|langganan|fita|promo|blanja|vcr|voucher|disc)\b", "", d_clean)
        nums = re.findall(r"\d+", d_clean)
        if nums:
            num_val = nums[0]
            # Standarisasi alias nominal kembar (169 -> 170, 282 -> 296, 56 -> 59)
            if num_val == "169": num_val = "170"
            elif num_val == "282": num_val = "296"
            elif num_val == "56": num_val = "59"
            return f"denom_{num_val}"
        else:
            return d_clean.strip()

    # Jika Description kosong atau tidak ada angka, cari angka dari nama Item
    clean_item = re.sub(r"^(dgd_|v\d+_|v\d+-|np_|mkios_|nontelco_|dtu_|mcash_p_|promobl_|blanja-v\d+-|diamond_telkomselidr_|star_telkomselidr_)", "", item)
    clean_item = re.sub(r"^(mlbb_|freefire_|aov_|cod_|rom_)", "", clean_item)
    clean_item = re.sub(r"(_diamonds|_diamond|_voucher|_vcr|_disc.*|_promo.*|_dg.*|_ndione|_ndispecial|_new\d*|_tact|_restac|_fntch|_rv_s|_ioh|_idv|_mytsel|_bns.*|_promoupoint|_funpack.*|_fita.*|_data|_digipos.*)$", "", clean_item)

    # Deteksi combo item: DTU_MLBB_11_1 -> 12 (angka kedua adalah bonus wajar <= 300)
    combo_match = re.search(r"(\d+)_(\d+)", clean_item)
    if combo_match:
        v1, v2 = int(combo_match.group(1)), int(combo_match.group(2))
        if v2 <= 300:
            return f"denom_{v1 + v2}"

    nums = re.findall(r"\d+", clean_item)
    if nums:
        num_val = nums[0]
        if num_val == "169": num_val = "170"
        elif num_val == "282": num_val = "296"
        elif num_val == "56": num_val = "59"
        return f"denom_{num_val}"

    return clean_item if clean_item else item

# 3. Hitung skor kebersihan nama item secara universal
def get_item_cleanliness_score(row: Dict[str, Any], item_key: str, desc_key: Optional[str], code_key: Optional[str] = None) -> float:
    desc = str(row.get(desc_key, "") or "").strip().lower() if desc_key else ""
    item = str(row.get(item_key, "") or "").strip().lower()
    code = str(row.get(code_key, "") or "").strip() if code_key else ""

    score = 0.0

    # Prioritaskan jika Code terisi resmi
    if code and code not in ["-", "none", ""]:
        score += 40.0

    # Prioritaskan jika memiliki Deskripsi lengkap
    if desc and desc not in ["-", "none", ""]:
        score += 50.0

    # Prioritaskan jika mengandung kata satuan mata uang resmi (diamond, voucher, bonds, robux, dll)
    unit_words = ["diamond", "voucher", "bond", "coin", "crystal", "robux", "gold", "gem", "koin", "zeny", "flower", "pass", "card"]
    for unit in unit_words:
        if unit in item:
            score += 20.0
            if unit in desc:
                score += 10.0
            break

    # Beri penalti jika mengandung singkatan channel / promo / vendor
    vendor_tags = ["_wd", "_dmver", "_rv_", "_gpy", "_byu", "_isat", "_telco", "_ft_", "_upo", "_dg_", "_idv", "_bns", "_promo", "_disc", "_cstm", "_tact", "_restac", "upsell", "recharge", "funpack", "fita", "mkios_", "_data", "_digipos", "digipos"]
    for tag in vendor_tags:
        if tag in item:
            score -= 15.0

    # Nilai lebih untuk teks yang lebih ringkas
    score -= len(item) * 0.1
    return score

# 4. Pilih baris item terbaik dari kumpulan kandidat varian yang sama
def pick_best_row(row_list: List[Dict[str, Any]], item_key: str, desc_key: Optional[str], code_key: Optional[str] = None) -> Dict[str, Any]:
    best_row = row_list[0]
    best_score = get_item_cleanliness_score(best_row, item_key, desc_key, code_key)

    for current_row in row_list[1:]:
        current_score = get_item_cleanliness_score(current_row, item_key, desc_key, code_key)
        if current_score > best_score:
            best_row = current_row
            best_score = current_score

    return best_row

# 5. Deduplikasi varian item per supplier untuk seluruh baris data katalog
def deduplicate_catalog_records(
    records: List[Dict[str, Any]],
    group_keys: Optional[List[str]] = None,
    item_key_param: Optional[str] = None,
    desc_key_param: Optional[str] = None,
    code_key_param: Optional[str] = None
):
    from collections import defaultdict

    original_cols = [str(k) for k in records[0].keys() if not str(k).startswith("_")]

    # Standarisasi nama kolom header: ID, Content, Item, Code, Description, Supplier
    display_cols = []
    for col_name in original_cols:
        col_lower = col_name.lower().replace("_", " ").strip()
        if "id" in col_lower and ("tim" in col_lower or "team" in col_lower or col_lower == "id"):
            display_cols.append("ID")
        elif "content" in col_lower or "game" in col_lower:
            display_cols.append("Content")
        elif "item" in col_lower:
            display_cols.append("Item")
        elif "code" in col_lower or "sku" in col_lower:
            display_cols.append("Code")
        elif "desc" in col_lower:
            display_cols.append("Description")
        elif "supplier" in col_lower or "vendor" in col_lower:
            display_cols.append("Supplier")
        else:
            display_cols.append(col_name)

    # Tentukan kolom acuan (gunakan parameter LLM jika ada, atau deteksi otomatis)
    item_key = item_key_param or original_cols[0]
    supplier_key = None
    content_key = None
    desc_key = desc_key_param
    code_key = code_key_param

    for col in original_cols:
        col_clean = str(col).lower().strip()
        if not item_key_param and "item" in col_clean:
            item_key = col
        elif "supplier" in col_clean or "vendor" in col_clean:
            supplier_key = col
        elif "content" in col_clean or "game" in col_clean:
            content_key = col
        elif not desc_key_param and ("description" in col_clean or "desc" in col_clean):
            desc_key = col
        elif not code_key_param and ("code" in col_clean or "sku" in col_clean):
            code_key = col

    # Kelompokkan per (Entitas Grup, Kunci Produk)
    groups = defaultdict(list)
    for row in records:
        item_val = row.get(item_key)
        desc_val = row.get(desc_key) if desc_key else ""

        # Lewati item uji coba / sampah
        if is_test_item(item_val, desc_val):
            continue

        if group_keys:
            entity_tuple = tuple(str(row.get(gk, "") or "").strip().lower() for gk in group_keys)
        else:
            supp = str(row.get(supplier_key, "") or "").strip().lower() if supplier_key else ""
            cont = str(row.get(content_key, "") or "").strip().lower() if content_key else ""
            entity_tuple = (supp, cont)

        # Jika ada kolom Code dan terisi, bersihkan Code sebagai kunci produk resmi
        code_val = str(row.get(code_key, "") or "").strip() if code_key else ""
        if code_val and code_val.lower() not in ["-", "none", ""]:
            prod_key = get_product_key(code_val, desc_val)
        else:
            prod_key = get_product_key(item_val, desc_val)

        groups[(entity_tuple, prod_key)].append(row)

    # Ambil 1 baris terbersih untuk setiap produk unik
    cleaned_rows = []
    for (entity_tuple, prod_key), row_list in groups.items():
        best_row = pick_best_row(row_list, item_key, desc_key, code_key)
        cleaned_rows.append([best_row.get(col_name) for col_name in original_cols])

    return display_cols, cleaned_rows

# 6. Normalisasi data flat menjadi tabel terstruktur via LLM Planner
def normalize_dataset(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    client = get_llm_client()
    if not client:
        raise ValueError("Variabel LLM_API_KEY belum disetel di file .env")

    model_name = get_llm_model()
    sample_records = records[:5]

    # Minta LLM menganalisis jenis data dan menentukan strategi normalisasi secara otomatis
    prompt = f"""Kamu adalah AI arsitek data untuk normalisasi spreadsheet.
Analisis struktur kolom dan contoh data berikut:
{json.dumps(sample_records, indent=2)}

Tentukan strategi normalisasi terbaik:
1. Pilih 'mode': 'catalog' jika tabel adalah daftar produk/SKU/katalog dengan varian promo dan supplier/vendor.
   Berikan:
   - "key_columns": daftar nama kolom pengelompokan entitas (misal ["Supplier", "Content"])
   - "code_column": nama kolom kode master jika ada (misal "Code" atau null)
   - "item_column": nama kolom kode item (misal "Item")
   - "desc_column": nama kolom deskripsi (misal "Description")
   - "erd_summary": penjelasan singkat

2. Pilih 'mode': 'erd' jika tabel adalah data transaksi datar (seperti Game, Denomination, Channel, Price) yang harus dipecah menjadi beberapa tabel relasional (Game Title, Game Denom, Channel, Game Dist Channel).
   Berikan:
   - "erd_summary": penjelasan alur relasi
   - "tables": daftar tabel relasional yang dipecah (nama tabel, kolom, PK, FK, dan data baris berupa array nilai, contoh: "rows": [["Free Fire", "FF"]]).

Keluarkan HANYA format JSON valid."""

    response = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": "Kamu adalah analis data. Hanya balas dalam format JSON yang valid."},
            {"role": "user", "content": prompt}
        ],
        response_format={"type": "json_object"},
        temperature=0.0
    )

    plan = json.loads(response.choices[0].message.content)
    mode = str(plan.get("mode", "")).lower().strip()

    # KASUS 1: Mode Katalog -> Eksekusi deduplikasi masal dengan konfigurasi dari LLM
    if mode == "catalog" or "catalog" in mode:
        key_cols = plan.get("key_columns")
        code_col = plan.get("code_column")
        item_col = plan.get("item_column")
        desc_col = plan.get("desc_column")

        clean_cols, clean_rows = deduplicate_catalog_records(
            records,
            group_keys=key_cols,
            item_key_param=item_col,
            desc_key_param=desc_col,
            code_key_param=code_col
        )

        return {
            "erd_summary": plan.get("erd_summary", "Deduplikasi varian katalog per supplier"),
            "tables": [
                {
                    "table_name": "Tabel Bersih",
                    "columns": clean_cols,
                    "primary_key": next((k for k in clean_cols if str(k).lower().strip() == "id"), None),
                    "foreign_keys": [],
                    "rows": clean_rows
                }
            ]
        }

    # KASUS 2: Mode Relasional ERD (4 Tabel)
    tables = plan.get("tables", [])
    for tbl in tables:
        if "table_name" not in tbl and "name" in tbl:
            tbl["table_name"] = tbl["name"]
        if isinstance(tbl.get("primary_key"), list):
            tbl["primary_key"] = ", ".join(str(x) for x in tbl["primary_key"])

        cols = tbl.get("columns", [])
        raw_rows = tbl.get("rows", [])
        formatted_rows = []
        for r in raw_rows:
            if isinstance(r, dict):
                formatted_rows.append([r.get(c) for c in cols])
            elif isinstance(r, (list, tuple)):
                formatted_rows.append(list(r))
            else:
                formatted_rows.append([r])
        tbl["rows"] = formatted_rows

    return {
        "erd_summary": plan.get("erd_summary", "Relasi normalisasi berhasil dibentuk"),
        "tables": tables
    }

# 7. Buat file Excel multi-tab dari hasil normalisasi
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
