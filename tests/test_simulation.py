import unittest
from starlette.testclient import TestClient
from app.main import app

class TestMultiSheetHarmonization(unittest.TestCase):
    def setUp(self):
        self.client = TestClient(app)

    def test_multi_sheet_consolidation_stateless(self):
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
                }
            ]
        }

        res = self.client.post("/api/consolidate", json=payload)
        self.assertEqual(res.status_code, 200)
        data = res.json()

        self.assertEqual(data["status"], "success")
        self.assertEqual(data["total_sheets"], 2)
        self.assertEqual(data["total_records"], 2)
        self.assertGreaterEqual(len(data["canonical_fields"]), 2)

        self.assertIn("Cabang_A", data["sheet_mappings"])
        self.assertIn("Cabang_B", data["sheet_mappings"])
        self.assertEqual(len(data["records"]), 2)

    def test_normalization_and_excel_export(self):
        # Data sampel seperti di Normalisasi Tabel.xlsx
        flat_records = [
            {"Channel": "A", "Game Title": "Free Fire", "Denomination": "10 diamond", "Price": 12000},
            {"Channel": "A", "Game Title": "Free Fire", "Denomination": "20 diamond", "Price": 25000},
            {"Channel": "B", "Game Title": "Free Fire", "Denomination": "10 diamond", "Price": 10000},
            {"Channel": "B", "Game Title": "Free Fire", "Denomination": "20 diamond", "Price": 20000}
        ]

        # 1. Uji endpoint /api/normalize
        res = self.client.post("/api/normalize", json={"records": flat_records})
        self.assertEqual(res.status_code, 200)
        data = res.json()
        self.assertEqual(data["status"], "success")
        self.assertIn("tables", data)
        self.assertGreaterEqual(len(data["tables"]), 2)

        # Pastikan tiap tabel memiliki struktur rapi
        for table in data["tables"]:
            self.assertIn("table_name", table)
            self.assertIn("columns", table)
            self.assertIn("rows", table)

        # 2. Uji endpoint /api/export-normalized-excel
        export_res = self.client.get("/api/export-normalized-excel")
        self.assertEqual(export_res.status_code, 200)
        self.assertEqual(
            export_res.headers["content-type"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        self.assertGreater(len(export_res.content), 1000)

if __name__ == "__main__":
    unittest.main()
