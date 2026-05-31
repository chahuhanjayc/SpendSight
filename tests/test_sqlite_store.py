import os
import tempfile
import unittest

from spendsight_store import SQLiteStore, migrate_legacy_json


class SQLiteStoreTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.tmp.name, "spendsight.db")

    def tearDown(self):
        self.tmp.cleanup()

    def test_store_round_trips_users_and_user_data(self):
        store = SQLiteStore(self.db_path)
        store.save_users([
            {"id": "admin", "password": "scrypt:test", "role": "admin"},
        ])
        data = {
            "expenses": [
                {
                    "id": "e1",
                    "amount": 125.5,
                    "category": "Groceries",
                    "subcategory": "Milk",
                    "date": "2026-05-31",
                    "payment_method": "Card",
                    "review_status": "reviewed",
                }
            ],
            "templates": [],
            "extra_income": [],
            "fixed_expenses": [],
            "goals": [{"id": "g1", "name": "Emergency", "target_amount": 1000}],
            "transaction_rules": [],
            "accounts": [{"id": "a1", "name": "Bank", "type": "bank", "balance": 5000}],
            "net_worth_snapshots": [{"id": "n1", "date": "2026-05-31", "net_worth": 5000}],
            "budget_envelopes": [{"id": "b1", "month": "2026-05", "category": "Groceries", "assigned": 3000}],
            "receipts": [{"id": "r1", "original_filename": "receipt.png", "extracted": {"amount": "125.5"}}],
            "recurring_payments": [],
            "payment_methods": ["Cash", "Card"],
            "custom_categories": {"Groceries": ["Milk"]},
            "budget_limits": {"Groceries": 3000},
            "income": {"monthly_salary": 50000, "salary_history": [{"amount": 50000, "effective_from": "2026-01-01"}]},
            "billing_start_day": 1,
            "currency_code": "INR",
            "schema_version": 3,
        }

        store.save_user_data("admin", data)
        loaded = store.load_user_data("admin")

        self.assertEqual(store.load_users(), [{"id": "admin", "password": "scrypt:test", "role": "admin"}])
        self.assertEqual(loaded["expenses"][0]["subcategory"], "Milk")
        self.assertEqual(loaded["accounts"][0]["balance"], 5000)
        self.assertEqual(loaded["budget_envelopes"][0]["category"], "Groceries")
        self.assertEqual(loaded["receipts"][0]["extracted"]["amount"], "125.5")
        self.assertEqual(loaded["payment_methods"], ["Cash", "Card"])
        self.assertEqual(loaded["custom_categories"], {"Groceries": ["Milk"]})
        self.assertEqual(loaded["budget_limits"], {"Groceries": 3000.0})

    def test_data_save_does_not_delete_cloud_tokens(self):
        store = SQLiteStore(self.db_path)
        store.save_users([{"id": "admin", "password": "hash", "role": "admin"}])
        store.save_cloud_tokens("admin", {"gdrive": {"access_token": "token"}})

        store.save_user_data("admin", {"expenses": [], "income": {"monthly_salary": 0}})

        self.assertEqual(store.load_cloud_tokens("admin"), {"gdrive": {"access_token": "token"}})

    def test_migrate_legacy_json_imports_users_and_user_data(self):
        workspace = self.tmp.name
        os.makedirs(os.path.join(workspace, "user_data"), exist_ok=True)
        with open(os.path.join(workspace, "users.json"), "w", encoding="utf-8") as handle:
            handle.write('[{"id":"admin","password":"hash","role":"admin"}]')
        with open(os.path.join(workspace, "user_data", "data_admin.json"), "w", encoding="utf-8") as handle:
            handle.write('{"expenses":[{"id":"e1","amount":10,"date":"2026-05-31"}],"income":{"monthly_salary":0}}')

        store = migrate_legacy_json(workspace, self.db_path)

        self.assertEqual(store.load_users()[0]["id"], "admin")
        self.assertEqual(store.load_user_data("admin")["expenses"][0]["amount"], 10)


if __name__ == "__main__":
    unittest.main()
