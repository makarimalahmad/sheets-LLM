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

    def test_catalog_variant_deduplication(self):
        catalog_records = [
            {"ID": 33, "Content": "arena_of_valor", "Item": "aov_1430", "Description": "AOV 1430 Vouchers", "Supplier": "garena"},
            {"ID": 1444, "Content": "arena_of_valor", "Item": "aov_1430_dg_gpy", "Description": "AOV 1430 Vouchers", "Supplier": "garena"},
            {"ID": 2056, "Content": "arena_of_valor", "Item": "aov_1430_dg_spend", "Description": "AOV 1430 Vouchers", "Supplier": "garena"},
            {"ID": 5001, "Content": "arena_of_valor", "Item": "aov_1430", "Description": "AOV 1430 Vouchers", "Supplier": "coda"},
            {"ID": 4, "Content": "arena_of_valor", "Item": "aov_18", "Description": "AOV 18 Voucher", "Supplier": "garena"}
        ]

        res = self.client.post("/api/normalize", json={"records": catalog_records})
        self.assertEqual(res.status_code, 200)
        data = res.json()
        self.assertEqual(data["status"], "success")
        self.assertEqual(len(data["tables"]), 1)
        
        clean_table = data["tables"][0]
        self.assertEqual(clean_table["table_name"], "Tabel Bersih")
        # 3 varian aov_1430 garena harus melebur jadi 1, coda tetap 1, aov_18 tetap 1 (Total: 3 baris)
        self.assertEqual(len(clean_table["rows"]), 3)

        # Baris pertama harus base item 'aov_1430' dari 'garena'
        garena_1430 = [r for r in clean_table["rows"] if r[2] == "aov_1430" and r[4] == "garena"]
        self.assertEqual(len(garena_1430), 1)
        self.assertEqual(garena_1430[0][0], 33)

        # Baris coda tetap ada
        coda_1430 = [r for r in clean_table["rows"] if r[2] == "aov_1430" and r[4] == "coda"]
        self.assertEqual(len(coda_1430), 1)

if __name__ == "__main__":
    unittest.main()
