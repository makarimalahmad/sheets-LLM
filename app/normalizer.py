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
    has_item = False
    has_supplier = False

    for col in sample_record.keys():
        col_name = str(col).lower()
        if "item" in col_name:
            has_item = True
        if "supplier" in col_name or "vendor" in col_name:
            has_supplier = True

    return has_item and has_supplier

# Cek apakah item adalah data uji coba / sampah (contoh: rom_test, laplace_test, cod_test)
def is_test_item(item_text: Any, desc_text: Any) -> bool:
    item = str(item_text or "").lower().strip()
    desc = str(desc_text or "").lower().strip()

    if "test" in item or "test" in desc:
        if desc == "" or desc == "-" or desc == "none" or "test" in desc:
            return True

    return False

# Ekstraksi identitas produk/denominasi kanonikal secara universal
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
        # Jika ada format bonus penjumlahan (misal: 11+1 Diamonds -> 12, 1580+110 Bonds -> 1690, 28+44+85 -> 157)
        plus_match = re.search(r"\d+(?:\s*\+\s*\d+)+", d_clean)
        if plus_match:
            plus_nums = [int(x) for x in re.findall(r"\d+", plus_match.group(0))]
            if plus_nums:
                return f"denom_{sum(plus_nums)}"

        d_clean = re.sub(r"\+.*", "", d_clean)
        d_clean = re.sub(r"\b(funpack|langganan|fita|promo|blanja|vcr|voucher|disc)\b", "", d_clean)
        nums = re.findall(r"\d+", d_clean)
        if nums:
            num_val = nums[0]
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

# Hitung skor kebersihan nama item secara universal
def get_item_cleanliness_score(row: Dict[str, Any], item_key: str, desc_key: Optional[str], code_key: Optional[str] = None) -> float:
    desc = str(row.get(desc_key, "") or "").strip().lower() if desc_key else ""
    item = str(row.get(item_key, "") or "").strip().lower()
    code = str(row.get(code_key, "") or "").strip() if code_key else ""

    score = 0.0

    # 1. Prioritaskan jika Code terisi resmi
    if code and code not in ["-", "none", ""]:
        score += 40.0

    # 2. Prioritaskan jika memiliki Deskripsi lengkap
    if desc and desc not in ["-", "none", ""]:
        score += 50.0

    # 3. Prioritaskan jika mengandung kata satuan mata uang resmi (diamond, voucher, bonds, robux, dll)
    unit_words = ["diamond", "voucher", "bond", "coin", "crystal", "robux", "gold", "gem", "koin", "zeny", "flower", "pass", "card"]
    for unit in unit_words:
        if unit in item:
            score += 20.0
            if unit in desc:
                score += 10.0
            break

    # 4. Beri penalti jika mengandung singkatan channel / promo / vendor
    vendor_tags = ["_wd", "_dmver", "_rv_", "_gpy", "_byu", "_isat", "_telco", "_ft_", "_upo", "_dg_", "_idv", "_bns", "_promo", "_disc", "_cstm", "_tact", "_restac", "upsell", "recharge", "funpack", "fita", "mkios_", "_data", "_digipos", "digipos"]
    for tag in vendor_tags:
        if tag in item:
            score -= 15.0

    # 5. Nilai lebih untuk teks yang lebih ringkas
    score -= len(item) * 0.1

    return score

# Pilih baris item terbaik dari kumpulan kandidat varian yang sama
def pick_best_row(row_list: List[Dict[str, Any]], item_key: str, desc_key: Optional[str], code_key: Optional[str] = None) -> Dict[str, Any]:
    best_row = row_list[0]
    best_score = get_item_cleanliness_score(best_row, item_key, desc_key, code_key)

    for current_row in row_list[1:]:
        current_score = get_item_cleanliness_score(current_row, item_key, desc_key, code_key)
        if current_score > best_score:
            best_row = current_row
            best_score = current_score

    return best_row

# Deduplikasi varian item per supplier untuk seluruh baris data
def deduplicate_catalog_records(records: List[Dict[str, Any]]):
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

    # Cari nama kolom yang cocok dengan loop sederhana
    item_key = original_cols[0]
    supplier_key = None
    content_key = None
    desc_key = None
    code_key = None

    for col in original_cols:
        col_clean = str(col).lower().strip()
        if "item" in col_clean:
            item_key = col
        elif "supplier" in col_clean or "vendor" in col_clean:
            supplier_key = col
        elif "content" in col_clean or "game" in col_clean:
            content_key = col
        elif "description" in col_clean or "desc" in col_clean:
            desc_key = col
        elif "code" in col_clean or "sku" in col_clean:
            code_key = col

    # 1. Kelompokkan per (Supplier, Content, Kunci Produk)
    groups = defaultdict(list)
    for row in records:
        item_val = row.get(item_key)
        desc_val = row.get(desc_key) if desc_key else ""

        # Lewati item uji coba / sampah
        if is_test_item(item_val, desc_val):
            continue

        supp = str(row.get(supplier_key, "") or "").strip().lower() if supplier_key else ""
        cont = str(row.get(content_key, "") or "").strip().lower() if content_key else ""

        # Jika ada kolom Code dan terisi, utamakan Code sebagai kunci produk resmi
        code_val = str(row.get(code_key, "") or "").strip().lower() if code_key else ""
        if code_val and code_val not in ["-", "none", ""]:
            prod_key = f"code_{code_val}"
        else:
            prod_key = get_product_key(item_val, desc_val)

        groups[(supp, cont, prod_key)].append(row)

    # 2. Ambil 1 baris terbersih untuk setiap produk unik
    cleaned_rows = []
    for (supp, cont, prod_key), row_list in groups.items():
        best_row = pick_best_row(row_list, item_key, desc_key, code_key)
        cleaned_rows.append([best_row.get(col_name) for col_name in original_cols])

    return display_cols, cleaned_rows

# Normalisasi data flat menjadi tabel terstruktur via LLM
def normalize_dataset(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    client = get_llm_client()
    if not client:
        raise ValueError("Variabel LLM_API_KEY belum disetel di file .env")

    model_name = get_llm_model()
    sample_records = records[:15]

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
