import os
import tempfile
import unittest

import app as spend_app


class AppStorageModeTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.old_env = {
            "SPENDSIGHT_STORAGE": os.environ.get("SPENDSIGHT_STORAGE"),
            "SPENDSIGHT_ADMIN_PASSWORD": os.environ.get("SPENDSIGHT_ADMIN_PASSWORD"),
        }
        self.old_globals = {
            "SQLITE_DB_FILE": spend_app.SQLITE_DB_FILE,
            "USERS_FILE": spend_app.USERS_FILE,
            "DATA_DIR": spend_app.DATA_DIR,
            "_sqlite_store": spend_app._sqlite_store,
        }
        spend_app.SQLITE_DB_FILE = os.path.join(self.tmp.name, "spendsight.db")
        spend_app.USERS_FILE = os.path.join(self.tmp.name, "users.json")
        spend_app.DATA_DIR = os.path.join(self.tmp.name, "user_data")
        os.makedirs(spend_app.DATA_DIR, exist_ok=True)
        spend_app._sqlite_store = None
        os.environ["SPENDSIGHT_STORAGE"] = "sqlite"
        os.environ["SPENDSIGHT_ADMIN_PASSWORD"] = "Strong!123"

    def tearDown(self):
        for key, value in self.old_env.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value
        for key, value in self.old_globals.items():
            setattr(spend_app, key, value)
        self.tmp.cleanup()

    def test_sqlite_mode_loads_and_saves_users(self):
        users = spend_app.load_users()
        self.assertEqual(users[0]["id"], "admin")
        self.assertEqual(users[0]["role"], "admin")

        spend_app.save_users([
            {"id": "admin", "password": users[0]["password"], "role": "admin"},
            {"id": "partner", "password": "scrypt:test", "role": "user"},
        ])

        self.assertEqual([u["id"] for u in spend_app.load_users()], ["admin", "partner"])
        self.assertFalse(os.path.exists(spend_app.USERS_FILE))

    def test_sqlite_mode_round_trips_current_user_data(self):
        with spend_app.app.test_request_context("/"):
            data = spend_app.load_data()
            data["expenses"].append({
                "id": "e1",
                "amount": 25,
                "category": "Groceries",
                "subcategory": "Milk",
                "date": "2026-05-31",
                "payment_method": "Cash",
            })
            spend_app.save_data(data)
            loaded = spend_app.load_data()

        self.assertEqual(loaded["expenses"][0]["subcategory"], "Milk")
        self.assertFalse(os.path.exists(os.path.join(spend_app.DATA_DIR, "data_admin.json")))


if __name__ == "__main__":
    unittest.main()
