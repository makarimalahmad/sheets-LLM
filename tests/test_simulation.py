import os
import unittest
from starlette.testclient import TestClient
from app.main import app
from app.database import DEFAULT_DB_PATH

class TestSheetsSimulation(unittest.TestCase):
    def setUp(self):
        if os.path.exists(DEFAULT_DB_PATH):
            os.remove(DEFAULT_DB_PATH)
        self.client = TestClient(app)

    def tearDown(self):
        if os.path.exists(DEFAULT_DB_PATH):
            os.remove(DEFAULT_DB_PATH)

    def test_end_to_end_ingestion_and_highest_eup(self):
        sheet1_payload = {
            "sheet_name": "Sheet1_User1",
            "headers": [
                "end user price include vat",
                "end user price ex vat",
                "payment rate",
                "payment cost",
                "unit price",
                "cogs",
                "margin"
            ],
            "rows": [
                ["111000", "100000", "0.03", "3000", "100000", "75000", "25000"]
            ]
        }
        res1 = self.client.post("/api/ingest", json=sheet1_payload)
        self.assertEqual(res1.status_code, 200)
        data1 = res1.json()
        self.assertEqual(data1["status"], "success")
        self.assertEqual(data1["column_mapping"]["end user price include vat"], "eup_inc_vat")

        sheet2_payload = {
            "sheet_name": "Sheet2_User2",
            "headers": [
                "EUP in VAT",
                "EUP ex VAT",
                "payment rate",
                "payment cost",
                "unit price",
                "cogs",
                "margin"
            ],
            "rows": [
                ["250000", "225225", "0.03", "7500", "225225", "150000", "75225"]
            ]
        }
        res2 = self.client.post("/api/ingest", json=sheet2_payload)
        self.assertEqual(res2.status_code, 200)
        data2 = res2.json()
        self.assertEqual(data2["status"], "success")
        self.assertEqual(data2["column_mapping"]["EUP in VAT"], "eup_inc_vat")

        highest_res = self.client.get("/api/products/highest-eup")
        self.assertEqual(highest_res.status_code, 200)
        highest_data = highest_res.json()
        self.assertEqual(highest_data["eup_inc_vat"], 250000.0)
        self.assertEqual(highest_data["source"], "Sheet2_User2")

        analytics_res = self.client.get("/api/analytics")
        self.assertEqual(analytics_res.status_code, 200)
        analytics_data = analytics_res.json()
        self.assertEqual(analytics_data["total_products"], 2)
        self.assertEqual(analytics_data["highest_eup_inc_vat"], 250000.0)
        self.assertEqual(analytics_data["lowest_eup_inc_vat"], 111000.0)

    def test_formatted_numbers_and_currency(self):
        payload = {
            "sheet_name": "Sheet_Formatted",
            "headers": ["EUP in VAT", "cogs", "payment rate"],
            "rows": [
                ["Rp 1.500.000,00", "750000", "2.5%"]
            ]
        }
        res = self.client.post("/api/ingest", json=payload)
        self.assertEqual(res.status_code, 200)
        
        products = self.client.get("/api/products").json()
        self.assertEqual(len(products), 1)
        self.assertEqual(products[0]["eup_inc_vat"], 1500000.0)
        self.assertEqual(products[0]["payment_rate"], 2.5)

if __name__ == "__main__":
    unittest.main()
