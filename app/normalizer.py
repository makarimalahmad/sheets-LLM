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

# Ekstraksi identitas produk/denominasi kanonikal secara universal
def get_universal_product_key(row_item: Dict[str, Any], item_key: str, desc_key: Optional[str]) -> str:
    desc_text = str(row_item.get(desc_key, "") or "").strip().lower() if desc_key else ""
    item_text = str(row_item.get(item_key, "") or "").strip().lower()

    # 1. Deteksi paket langganan, membership, pass, atau bundle khusus
    special_passes = [
        ("twilight", "twilight_pass"),
        ("weekly diamond pass", "weekly_diamond_pass"),
        ("weekly pass", "weekly_diamond_pass"),
        ("wdp", "weekly_diamond_pass"),
        ("starlight", "starlight_member"),
        ("star_telkomselidr", "starlight_member"),
        ("monthly epic", "monthly_epic_bundle"),
        ("weekly elite", "weekly_elite_bundle"),
        ("diamond pass", "diamond_pass"),
        ("aurum pass", "aurum_pass"),
        ("aurumpass", "aurum_pass"),
        ("companionship", "companionship_pack"),
        ("cpack", "companionship_pack"),
        ("express supply pass", "express_supply_pass"),
        ("hsr_esp", "express_supply_pass"),
        ("royal grow pass", "royal_grow_pass"),
        ("growtopia_rgp", "royal_grow_pass"),
        ("gem abundance", "gem_abundance"),
        ("growtopia_abundance", "gem_abundance"),
        ("gem bounty", "gem_bounty"),
        ("growtopia_bounty", "gem_bounty"),
        ("gem fountain", "gem_fountain"),
        ("growtopia_fountain", "gem_fountain"),
        ("rainin", "rainin_gems"),
        ("chest o", "chest_o_gems"),
        ("battle pass", "battle_pass"),
        ("anti_peluru", "anti_peluru"),
        ("anti peluru", "anti_peluru"),
        ("komposit", "kotak_komposit"),
        ("pemula", "pilihan_pemula")
    ]
    for pattern, key_name in special_passes:
        if pattern in desc_text or pattern in item_text:
            return key_name

    # 2. Ambil angka nominal dari Description (nama produk resmi bahasa manusia)
    if desc_text and desc_text != "-" and desc_text != "none":
        # Bersihkan bundling tambahan seperti fita, exercise plan, kuota mb/gb, funpack
        d_clean = re.sub(r"\+.*fita.*", "", desc_text)
        d_clean = re.sub(r"\+.*plan.*", "", d_clean)
        d_clean = re.sub(r"\+.*mb.*", "", d_clean)
        d_clean = re.sub(r"\+.*gb.*", "", d_clean)
        d_clean = re.sub(r"\(.*?\)", "", d_clean)
        d_clean = re.sub(r"\b(funpack|langganan|fita|promo|blanja|vcr|voucher|vouchers|disc|discount)\b", "", d_clean)

        # Jika ada format penjumlahan bonus (contoh: 11+1 Diamonds -> 12)
        plus_match = re.search(r"(\d+)\s*\+\s*(\d+)", d_clean)
        if plus_match:
            v1, v2 = int(plus_match.group(1)), int(plus_match.group(2))
            if v1 in [11, 25, 53, 77, 154, 256, 367, 503, 774, 1708, 217, 4003]:
                return f"denom_{v1 + v2}"

        d_clean = re.sub(r"\+.*", "", d_clean)
        d_nums = re.findall(r"\d+", d_clean)
        if d_nums:
            num_val = d_nums[0]
            # Standarisasi alias nominal kembar (169 -> 170, 282 -> 296, 56 -> 59)
            if num_val == "169": num_val = "170"
            elif num_val == "282": num_val = "296"
            elif num_val == "56": num_val = "59"
            return f"denom_{num_val}"

    # 3. Jika Description kosong atau tidak ada angka di Description, bersihkan nama Item
    clean_item = item_text
    clean_item = re.sub(r"^(dgd_|v\d+_|v\d+-|np_|mkios_|nontelco_|dtu_|mcash_p_|promobl_|blanja-v\d+-|diamond_telkomselidr_|star_telkomselidr_)", "", clean_item)
    clean_item = re.sub(r"^(mlbb_|freefire_|aov_|cod_|rom_|hsr_|lad_|supersus_|booya_|bleach_|king_|rm_|gc_|ab_|as_|aether_gazer_)", "", clean_item)
    clean_item = re.sub(r"(_diamonds|_diamond|_voucher|_vcr|_disc.*|_promo.*|_dg.*|_ndione|_new\d*|_tact|_restac|_fntch|_rv_s|_ioh|_idv|_mytsel|_bns.*|_promoupoint|_funpack.*|_fita.*)$", "", clean_item)

    # Deteksi kode item penjumlahan (contoh: DTU_MLBB_11_1 -> 12)
    combo_match = re.search(r"(\d+)_(\d+)", clean_item)
    if combo_match:
        v1, v2 = int(combo_match.group(1)), int(combo_match.group(2))
        if v1 in [11, 25, 53, 77, 154, 256, 367, 503, 774]:
            return f"denom_{v1 + v2}"

    item_nums = re.findall(r"\d+", clean_item)
    if item_nums:
        num_val = item_nums[0]
        if num_val == "169": num_val = "170"
        elif num_val == "282": num_val = "296"
        elif num_val == "56": num_val = "59"
        return f"denom_{num_val}"

    return clean_item if clean_item else item_text

# Cek apakah suatu baris adalah item uji coba/sampah (contoh: rom_test, laplace_test, cod_test, rox_test)
def is_junk_test_item(row_item: Dict[str, Any], item_key: str, desc_key: Optional[str]) -> bool:
    item_str = str(row_item.get(item_key, "") or "").lower().strip()
    desc_str = str(row_item.get(desc_key, "") or "").lower().strip() if desc_key else ""

    # Cek kata test di nama item atau deskripsi
    has_test_word = bool(re.search(r"\btest\b|_test\b|testitem|test_item|item_test|item test|^[a-z]+test$", item_str)) or ("test" in desc_str)
    if has_test_word:
        # Jika deskripsinya kosong, tanda strip '-', atau deskripsinya eksplisit bertuliskan kata test -> SAMPAH!
        if not desc_str or desc_str in ["-", "none", ""] or ("test" in desc_str):
            return True
    return False

# Pilih baris item terbaik dari kumpulan kandidat varian yang sama
def choose_best_catalog_item(candidate_rows: List[Dict[str, Any]], game_name: str, item_key: str, desc_key: str) -> Dict[str, Any]:
    clean_game = game_name.replace("_", "").lower()

    def score_row(row_candidate: Dict[str, Any]):
        desc = str(row_candidate.get(desc_key, "") or "").strip()
        item = str(row_candidate.get(item_key, "") or "").strip().lower()

        # Prioritas 1: Punya deskripsi lengkap (bukan kosong/None/strip)
        has_desc = 20 if (desc and desc not in ["-", "none", ""]) else 0

        # Prioritas 2: Hindari nama item yang memiliki akhiran promo/distributor/channel/bundling
        has_promo_suffix = 1 if any(tag in item for tag in [
            "idv", "ndione", "ndispecial", "ndityp", "new26", "newioh", "dg_", "dgspend", "dgtp",
            "disc", "promo", "spend", "tact", "restac", "ewl", "mager", "ucnormal", "fita", "funpack",
            "p_dg", "p_brandcom", "p_la", "p_blu", "bns", "free", "sktes", "cstm", "ctm", "dgd_", "blanja", "dtu_"
        ]) else 0

        # Prioritas 3: Nama item yang bersih/standar (misal freefire_12 lebih bagus dari freefire_12_funpack_op)
        has_clean_game_prefix = 5 if item.startswith((clean_game, game_name.lower(), "aov_", "cod_")) and not has_promo_suffix else 0
        is_clean_base = 3 if not has_promo_suffix else 0

        # Prioritas 4: Nama item mengandung kata game jika relevan
        has_game_name = 3 if clean_game in item.replace("_", "") else 0

        # Prioritas 5: Lebih pendek lebih bagus
        return (has_desc, -has_promo_suffix, has_clean_game_prefix, is_clean_base, has_game_name, -len(item))

    sorted_candidates = sorted(candidate_rows, key=score_row, reverse=True)
    return sorted_candidates[0]

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

    # Filter buang baris test sampah terlebih dahulu
    clean_input_records = [r for r in records if not is_junk_test_item(r, item_key, desc_key)]

    # Kelompokkan per (Supplier, Content) agar supplier dan game berbeda tidak pernah tercampur
    by_supp_content = defaultdict(list)
    for row_item in clean_input_records:
        supp_val = str(row_item.get(supplier_key, "") or "").strip().lower() if supplier_key else ""
        cont_val = str(row_item.get(content_key, "") or "").strip().lower() if content_key else ""
        by_supp_content[(supp_val, cont_val)].append(row_item)

    cleaned_rows = []

    for (supp_val, cont_val), group_rows in by_supp_content.items():
        # Kelompokkan item berdasarkan identitas produk kanonikal
        product_groups = defaultdict(list)

        for row in group_rows:
            product_key = get_universal_product_key(row, item_key, desc_key)
            product_groups[product_key].append(row)

        # Pilih 1 baris terbaik yang bersih untuk setiap produk unik
        for product_key, candidate_rows in product_groups.items():
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
