import json
import os
import io
import tempfile
import unittest
from datetime import date

import app as spend_app


class PersistenceRepairTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmp.cleanup()

    def test_atomic_write_json_creates_parent_directory_and_serializes_complete_json(self):
        target = os.path.join(self.tmp.name, "nested", "state.json")

        spend_app._atomic_write_json(
            target,
            {"expenses": [{"amount": 25, "date": date(2026, 5, 31)}]},
            default=str,
        )

        with open(target, "r", encoding="utf-8") as f:
            saved = json.load(f)

        self.assertEqual(saved, {"expenses": [{"amount": 25, "date": "2026-05-31"}]})
        leftovers = [name for name in os.listdir(os.path.dirname(target)) if name.endswith(".tmp")]
        self.assertEqual(leftovers, [])

    def test_atomic_write_json_preserves_existing_file_when_serialization_fails(self):
        target = os.path.join(self.tmp.name, "state.json")
        with open(target, "w", encoding="utf-8") as f:
            json.dump({"stable": True}, f)

        with self.assertRaises(TypeError):
            spend_app._atomic_write_json(target, {"bad": object()})

        with open(target, "r", encoding="utf-8") as f:
            self.assertEqual(json.load(f), {"stable": True})

        leftovers = [
            name
            for name in os.listdir(self.tmp.name)
            if name.startswith(".state.json.") and name.endswith(".tmp")
        ]
        self.assertEqual(leftovers, [])

    def test_load_json_or_default_quarantines_corrupt_json(self):
        target = os.path.join(self.tmp.name, "broken.json")
        with open(target, "w", encoding="utf-8") as f:
            f.write("{not valid json")

        loaded = spend_app._load_json_or_default(target, {"fresh": True})

        self.assertEqual(loaded, {"fresh": True})
        self.assertFalse(os.path.exists(target))
        quarantined = [name for name in os.listdir(self.tmp.name) if name.startswith("broken.json.") and name.endswith(".corrupt")]
        self.assertEqual(len(quarantined), 1)

    def test_read_limited_response_rejects_large_restore_payload(self):
        response = io.BytesIO(b"x" * 6)

        with self.assertRaises(ValueError):
            spend_app._read_limited_response(response, max_bytes=5)

    def test_validate_restore_data_accepts_backup_and_normalizes_defaults(self):
        backup = {
            "expenses": [{"amount": "12.50", "date": "2026-05-31", "category": "Groceries"}],
            "income": {"monthly_salary": "50000"},
            "payment_methods": [],
        }

        validated = spend_app._validate_restore_data(backup)

        self.assertEqual(validated["expenses"][0]["quantity"], None)
        self.assertEqual(validated["expenses"][0]["unit"], "")
        self.assertEqual(validated["payment_methods"], spend_app.DEFAULT_PAYMENT_METHODS)
        self.assertEqual(validated["templates"], [])
        self.assertEqual(validated["extra_income"], [])
        self.assertEqual(validated["fixed_expenses"], [])
        self.assertEqual(validated["budget_limits"], {})
        self.assertEqual(validated["income"]["salary_history"][0]["amount"], 50000.0)

    def test_validate_restore_data_rejects_malformed_backup_shapes_and_values(self):
        invalid_backups = [
            ("root is not object", []),
            ("expenses is not list", {"expenses": "not-a-list"}),
            ("expense item is not object", {"expenses": ["not-an-object"]}),
            ("income is not object", {"income": []}),
            ("amount is invalid", {"expenses": [{"amount": "not-money"}]}),
            ("date is invalid", {"expenses": [{"date": "31-05-2026"}]}),
        ]

        for label, backup in invalid_backups:
            with self.subTest(label=label):
                with self.assertRaises(ValueError):
                    spend_app._validate_restore_data(backup)

    def test_cloud_credentials_configured_requires_real_values_for_all_service_fields(self):
        helper = getattr(spend_app, "_cloud_credentials_configured", None)
        if helper is None:
            self.skipTest("_cloud_credentials_configured is not available yet")

        old_credentials = spend_app.CLOUD_CREDENTIALS
        spend_app.CLOUD_CREDENTIALS = {
            "gdrive": {
                "client_id": "real-google-client-id",
                "client_secret": "real-google-client-secret",
            },
            "onedrive": {"client_id": "YOUR_MICROSOFT_CLIENT_ID"},
            "dropbox": {"app_key": "real-dropbox-key", "app_secret": ""},
        }
        try:
            self.assertTrue(helper("gdrive"))
            self.assertFalse(helper("onedrive"))
            self.assertFalse(helper("dropbox"))
            self.assertFalse(helper("unknown"))
        finally:
            spend_app.CLOUD_CREDENTIALS = old_credentials


if __name__ == "__main__":
    unittest.main()
