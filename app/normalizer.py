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

# Deteksi jenis satuan mata uang game agar mata uang berbeda (misal B-Chips vs Crystals, Key vs Diamond) tidak tertukar
def detect_currency_unit(text: str) -> str:
    clean_text = str(text or "").lower()
    unit_patterns = [
        (r"\b(b-chips|b_chips|bchips|bc)\b", "bchips"),
        (r"\b(crystals|crystal|cr)\b", "crystal"),
        (r"\b(diamonds|diamond|dm|dmn)\b", "diamond"),
        (r"\b(keys|key)\b", "key"),
        (r"\b(souls|soul)\b", "soul"),
        (r"\b(money)\b", "money"),
        (r"\b(gems|gem)\b", "gem"),
        (r"\b(ap)\b", "ap"),
        (r"\b(zeny)\b", "zeny"),
        (r"\b(big cat coins|bcc)\b", "bcc"),
        (r"\b(coins|coin|koins|koin)\b", "coin"),
        (r"\b(vouchers|voucher|vcr)\b", "voucher"),
        (r"\b(bonds|bond)\b", "bond"),
        (r"\b(tokens|token)\b", "token"),
        (r"\b(robux|rbx)\b", "robux")
    ]
    for pattern, unit_label in unit_patterns:
        if re.search(pattern, clean_text):
            return unit_label
    return ""

# 2. Ekstraksi identitas produk/denominasi kanonikal secara universal
def get_product_key(item_text: Any, desc_text: Any) -> str:
    item = str(item_text or "").lower().strip()
    desc = str(desc_text or "").lower().strip()

    # Ambil angka nominal dari Description jika ada (Description adalah nama produk resmi)
    if desc and desc not in ["-", "none", ""]:
        d_clean = re.sub(r"\(.*?\)", "", desc)
        # Cek jika ada format penjumlahan (misal: 1580 + 110) yang hasilnya cocok dengan nama Item
        plus_match = re.search(r"(\d+)\s*\+\s*(\d+)", d_clean)
        if plus_match:
            total_sum = int(plus_match.group(1)) + int(plus_match.group(2))
            if str(total_sum) in item:
                unit_type = detect_currency_unit(f"{d_clean} {item}")
                return f"denom_{total_sum}_{unit_type}" if unit_type else f"denom_{total_sum}"

        # Buang tanda + dan semua teks bonus promo di belakangnya agar melebur ke base item
        d_clean = re.sub(r"\+.*", "", d_clean)
        d_clean = re.sub(r"\b(funpack|langganan|fita|promo|blanja|vcr|voucher|disc)\b", "", d_clean)
        d_clean = re.sub(r"\s+", " ", d_clean).strip()

        # Jika deskripsi adalah paket/pass/durasi/multiplier (misal: x2, x3, Pass, Card, Hari),
        # jangan diubah menjadi denom_<angka> agar tidak bertabrakan dengan nominal murni (contoh: x3 != 3 Diamonds)
        if re.search(r"\b(x\d+|pass|card|paket|bundle|fund|hari|days|bulan|minggu|plan|member|service|privilege|elite|monthly|weekly)\b", d_clean):
            return d_clean

        # Ambil angka paling kanan (nominal pecahan, bukan angka judul game seperti 'Honkai Impact 3') serta dukung angka desimal (1.3B, 10.4B)
        nums = re.findall(r"\d+(?:\.\d+)?", d_clean)
        if nums:
            num_val = nums[-1]
            # Standarisasi alias nominal kembar (169 -> 170, 282 -> 296, 56 -> 59)
            if num_val == "169": num_val = "170"
            elif num_val == "282": num_val = "296"
            elif num_val == "56": num_val = "59"
            unit_type = detect_currency_unit(f"{d_clean} {item}")
            return f"denom_{num_val}_{unit_type}" if unit_type else f"denom_{num_val}"
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
            unit_type = detect_currency_unit(item)
            return f"denom_{v1 + v2}_{unit_type}" if unit_type else f"denom_{v1 + v2}"

    nums = re.findall(r"\d+", clean_item)
    if nums:
        num_val = nums[0]
        if num_val == "169": num_val = "170"
        elif num_val == "282": num_val = "296"
        elif num_val == "56": num_val = "59"
        unit_type = detect_currency_unit(item)
        return f"denom_{num_val}_{unit_type}" if unit_type else f"denom_{num_val}"

    return clean_item if clean_item else item

# 3. Pilih baris item terbaik dari sekumpulan varian produk yang sama
def pick_best_row(
    row_list: List[Dict[str, Any]],
    item_key: str,
    desc_key: Optional[str],
    code_key: Optional[str] = None,
    extra_unit_words: Optional[List[str]] = None,
    extra_promo_tags: Optional[List[str]] = None
) -> Dict[str, Any]:
    unit_words = ["diamond", "voucher", "bond", "coin", "crystal", "robux", "gold", "gem", "koin", "zeny", "flower", "pass", "card"]
    if extra_unit_words:
        for word in extra_unit_words:
            clean_word = str(word).lower().strip()
            if clean_word and clean_word not in unit_words:
                unit_words.append(clean_word)

    promo_tags = ["_wd", "_dmver", "_rv_", "_gpy", "_byu", "_isat", "_telco", "_ft_", "_upo", "_dg_", "_idv", "_bns", "_promo", "_disc", "_cstm", "_tact", "_restac", "upsell", "recharge", "funpack", "fita", "mkios_", "_data", "_digipos"]
    if extra_promo_tags:
        for tag in extra_promo_tags:
            clean_tag = str(tag).lower().strip()
            if clean_tag and clean_tag not in unit_words and clean_tag not in promo_tags:
                promo_tags.append(clean_tag)

    def cek_prioritas(row):
        code_ada = bool(code_key and str(row.get(code_key, "") or "").strip() not in ["-", "none", ""])
        desc_ada = bool(desc_key and str(row.get(desc_key, "") or "").strip() not in ["-", "none", ""])
        item_str = str(row.get(item_key, "") or "").strip().lower()
        nama_resmi = any(unit in item_str for unit in unit_words)
        bukan_promo = not any(tag in item_str for tag in promo_tags)

        # Urutan pemenang: Code resmi -> Ada deskripsi -> Nama satuan resmi -> Bukan promo -> Teks terpendek
        return (code_ada, desc_ada, nama_resmi, bukan_promo, -len(item_str))

    return max(row_list, key=cek_prioritas)

# 4. Deduplikasi varian item per supplier untuk seluruh baris data katalog
def deduplicate_catalog_records(
    records: List[Dict[str, Any]],
    group_keys: Optional[List[str]] = None,
    item_key_param: Optional[str] = None,
    desc_key_param: Optional[str] = None,
    code_key_param: Optional[str] = None,
    extra_unit_words: Optional[List[str]] = None,
    extra_promo_tags: Optional[List[str]] = None
):
    from collections import defaultdict

    original_cols = [str(col_key) for col_key in records[0].keys() if not str(col_key).startswith("_")]
    cols_lookup = {col_name.lower().strip(): col_name for col_name in original_cols}

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

    # Validasi nama kolom dari LLM agar cocok dengan header asli
    def resolve_col_name(candidate: Optional[str]) -> Optional[str]:
        if not candidate:
            return None
        clean_cand = str(candidate).strip()
        if clean_cand.lower() in ["null", "none", ""]:
            return None
        if clean_cand in original_cols:
            return clean_cand
        return cols_lookup.get(clean_cand.lower())

    item_key = resolve_col_name(item_key_param) or original_cols[0]
    desc_key = resolve_col_name(desc_key_param)
    code_key = resolve_col_name(code_key_param)
    supplier_key = None
    content_key = None

    for col in original_cols:
        col_clean = str(col).lower().strip()
        if not resolve_col_name(item_key_param) and "item" in col_clean:
            item_key = col
        elif "supplier" in col_clean or "vendor" in col_clean:
            supplier_key = col
        elif "content" in col_clean or "game" in col_clean:
            content_key = col
        elif not desc_key and ("description" in col_clean or "desc" in col_clean):
            desc_key = col
        elif not code_key and ("code" in col_clean or "sku" in col_clean):
            code_key = col

    resolved_group_keys = []
    if group_keys:
        for group_col in group_keys:
            matched_col = resolve_col_name(group_col)
            if matched_col:
                resolved_group_keys.append(matched_col)

    # Langkah 1: Gabungkan baris dengan Code resmi yang sama persis per (Entitas Grup, Cleaned Code)
    step1_groups = defaultdict(list)
    no_code_rows = []
    for row in records:
        item_val = row.get(item_key)
        desc_val = row.get(desc_key) if desc_key else ""

        # Lewati item uji coba / sampah
        if is_test_item(item_val, desc_val):
            continue

        if resolved_group_keys:
            entity_tuple = tuple(str(row.get(group_col, "") or "").strip().lower() for group_col in resolved_group_keys)
        else:
            supp = str(row.get(supplier_key, "") or "").strip().lower() if supplier_key else ""
            cont = str(row.get(content_key, "") or "").strip().lower() if content_key else ""
            entity_tuple = (supp, cont)

        code_val = str(row.get(code_key, "") or "").strip().lower() if code_key else ""
        if code_val and code_val not in ["-", "none", ""]:
            clean_code = re.sub(r"(_data|_promo.*|_disc.*|_new\d*)$", "", code_val)
            step1_groups[(entity_tuple, clean_code)].append(row)
        else:
            no_code_rows.append(row)

    merged_by_code = [
        pick_best_row(row_list, item_key, desc_key, code_key, extra_unit_words, extra_promo_tags)
        for row_list in step1_groups.values()
    ] + no_code_rows

    # Langkah 2: Gabungkan varian alias produk yang memiliki kunci produk & satuan yang sama
    step2_groups = defaultdict(list)
    for row in merged_by_code:
        item_val = row.get(item_key)
        desc_val = row.get(desc_key) if desc_key else ""

        if resolved_group_keys:
            entity_tuple = tuple(str(row.get(group_col, "") or "").strip().lower() for group_col in resolved_group_keys)
        else:
            supp = str(row.get(supplier_key, "") or "").strip().lower() if supplier_key else ""
            cont = str(row.get(content_key, "") or "").strip().lower() if content_key else ""
            entity_tuple = (supp, cont)

        prod_key = get_product_key(item_val, desc_val)
        step2_groups[(entity_tuple, prod_key)].append(row)

    # Ambil 1 baris terbersih untuk setiap produk unik
    cleaned_rows = []
    for (entity_tuple, prod_key), row_list in step2_groups.items():
        best_row = pick_best_row(row_list, item_key, desc_key, code_key, extra_unit_words, extra_promo_tags)
        cleaned_rows.append([best_row.get(col_name) for col_name in original_cols])

    return display_cols, cleaned_rows

# 5. Normalisasi data flat menjadi tabel terstruktur via LLM Planner
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
   - "promo_tags": daftar akhiran/kata promo yang terdeteksi di kode Item (misal ["_promo", "_disc"])
   - "unit_words": daftar satuan produk yang terdeteksi (misal ["diamond", "voucher"])
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
        extra_units = plan.get("unit_words")
        extra_promos = plan.get("promo_tags")

        clean_cols, clean_rows = deduplicate_catalog_records(
            records,
            group_keys=key_cols,
            item_key_param=item_col,
            desc_key_param=desc_col,
            code_key_param=code_col,
            extra_unit_words=extra_units,
            extra_promo_tags=extra_promos
        )

        primary_key_col = None
        for col_name in clean_cols:
            if str(col_name).lower().strip() == "id":
                primary_key_col = col_name
                break

        return {
            "erd_summary": plan.get("erd_summary", "Deduplikasi varian katalog per supplier"),
            "tables": [
                {
                    "table_name": "Tabel Bersih",
                    "columns": clean_cols,
                    "primary_key": primary_key_col,
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
            tbl["primary_key"] = ", ".join(str(key_col) for key_col in tbl["primary_key"])

        cols = tbl.get("columns", [])
        raw_rows = tbl.get("rows", [])
        formatted_rows = []
        for row_item in raw_rows:
            if isinstance(row_item, dict):
                formatted_rows.append([row_item.get(col_name) for col_name in cols])
            elif isinstance(row_item, (list, tuple)):
                formatted_rows.append(list(row_item))
            else:
                formatted_rows.append([row_item])
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
