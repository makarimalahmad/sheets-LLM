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

        self.assertEqual(data["status"], "success")
        self.assertEqual(data["total_sheets"], 3)
        self.assertEqual(data["total_records"], 3)
        self.assertGreaterEqual(len(data["canonical_fields"]), 2)

        self.assertIn("Cabang_A", data["sheet_mappings"])
        self.assertIn("Cabang_B", data["sheet_mappings"])
        self.assertIn("Cabang_C", data["sheet_mappings"])

        map_a = data["sheet_mappings"]["Cabang_A"]["EUP in VAT"]
        map_b = data["sheet_mappings"]["Cabang_B"]["end user price include vat"]
        map_c = data["sheet_mappings"]["Cabang_C"]["harga jual user + ppn"]

        self.assertEqual(map_a, map_b)
        self.assertEqual(map_b, map_c)

        for rec in data["records"]:
            self.assertIn("_source_sheet", rec)
            self.assertIn(map_a, rec)
            self.assertIsNotNone(rec[map_a])

        # Test ambil riwayat dari database
        res_history = self.client.get("/api/consolidated-records")
        self.assertEqual(res_history.status_code, 200)
        history_data = res_history.json()
        self.assertEqual(len(history_data["records"]), 3)

        # Test bersihkan data
        res_clear = self.client.delete("/api/clear-data")
        self.assertEqual(res_clear.status_code, 200)

        res_after = self.client.get("/api/consolidated-records")
        self.assertEqual(len(res_after.json()["records"]), 0)

if __name__ == "__main__":
    unittest.main()
