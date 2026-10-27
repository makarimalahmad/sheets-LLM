import os
import unittest
from starlette.testclient import TestClient
from app.main import app
from app.database import DEFAULT_DB_PATH

class TestMultiSheetHarmonization(unittest.TestCase):
    def setUp(self):
        if os.path.exists(DEFAULT_DB_PATH):
            os.remove(DEFAULT_DB_PATH)
        self.client = TestClient(app)

    def tearDown(self):
        if os.path.exists(DEFAULT_DB_PATH):
            try:
                os.remove(DEFAULT_DB_PATH)
            except PermissionError:
                pass

    def test_multi_sheet_consolidation_zero_shot(self):
        """
        Pengujian multi-sheet tanpa kamus kata kunci manual (Zero Hardcoded Dictionary).
        Mengirim 3 sheet dengan gaya bahasa dan singkatan yang sangat berbeda.
        """
        payload = {
            "sheets": [
                {
                    "sheet_name": "Cabang_A",
                    "headers": ["EUP in VAT", "cogs", "untung"],
                    "rows": [["150000", "100000", "50000"]]
                },
                {
                    "sheet_name": "Cabang_B",
                    "headers": ["end user price include vat", "hpp", "profit margin"],
                    "rows": [["300000", "200000", "100000"]]
                },
                {
                    "sheet_name": "Cabang_C",
                    "headers": ["harga jual user + ppn", "modal pokok", "margin bersih"],
                    "rows": [["450000", "300000", "150000"]]
                }
            ]
        }

        res = self.client.post("/api/consolidate", json=payload)
        self.assertEqual(res.status_code, 200)
        data = res.json()

        # Validasi output konsolidasi
        self.assertEqual(data["status"], "success")
        self.assertEqual(data["total_sheets"], 3)
        self.assertEqual(data["total_records"], 3)

        # Pastikan ada canonical fields yang ditemukan LLM
        self.assertGreaterEqual(len(data["canonical_fields"]), 2)

        # Pastikan setiap sheet memiliki mapping
        self.assertIn("Cabang_A", data["sheet_mappings"])
        self.assertIn("Cabang_B", data["sheet_mappings"])
        self.assertIn("Cabang_C", data["sheet_mappings"])

        # Periksa bahwa ketiga sheet memetakan kolom harga ke field kanonikal yang persis sama
        map_a = data["sheet_mappings"]["Cabang_A"]["EUP in VAT"]
        map_b = data["sheet_mappings"]["Cabang_B"]["end user price include vat"]
        map_c = data["sheet_mappings"]["Cabang_C"]["harga jual user + ppn"]

        self.assertEqual(map_a, map_b, "EUP in VAT dan end user price include vat harus dipetakan ke field kanonikal yang sama")
        self.assertEqual(map_b, map_c, "end user price include vat dan harga jual user + ppn harus dipetakan ke field kanonikal yang sama")

        # Cek struktur setiap record pada output gabungan
        for rec in data["records"]:
            self.assertIn("_source_sheet", rec)
            self.assertIn(map_a, rec)
            self.assertIsNotNone(rec[map_a])

    def test_preview_harmonize_schema(self):
        payload = {
            "sheets": [
                {
                    "sheet_name": "Sheet_1",
                    "headers": ["EUP ex VAT", "unit price"],
                    "rows": [["100000", "90000"]]
                },
                {
                    "sheet_name": "Sheet_2",
                    "headers": ["end user price excl vat", "harga per unit"],
                    "rows": [["200000", "180000"]]
                }
            ]
        }
        res = self.client.post("/api/harmonize-schema", json=payload)
        self.assertEqual(res.status_code, 200)
        data = res.json()
        self.assertIn("canonical_fields", data)
        self.assertIn("sheet_mappings", data)
        self.assertEqual(len(data["sheet_mappings"]), 2)

if __name__ == "__main__":
    unittest.main()
