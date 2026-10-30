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

        # Data langsung kembali dari AI ke respons tanpa database
        self.assertEqual(len(data["records"]), 2)

if __name__ == "__main__":
    unittest.main()
