import json
import os
import re
import tempfile
import unittest

from werkzeug.security import generate_password_hash

import app as spend_app


def extract_csrf(response):
    match = re.search(r'name="csrf_token" value="([^"]+)"', response.get_data(as_text=True))
    if not match:
        raise AssertionError("CSRF token not found in response")
    return match.group(1)


class SpendSightRepairTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.old_users_file = spend_app.USERS_FILE
        self.old_data_dir = spend_app.DATA_DIR
        self.old_cloud_tokens_file = spend_app.CLOUD_TOKENS_FILE

        spend_app.USERS_FILE = os.path.join(self.tmp.name, "users.json")
        spend_app.DATA_DIR = os.path.join(self.tmp.name, "user_data")
        spend_app.CLOUD_TOKENS_FILE = os.path.join(self.tmp.name, "cloud_tokens.json")
        os.makedirs(spend_app.DATA_DIR, exist_ok=True)

        with open(spend_app.USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(
                [{"id": "admin", "password": generate_password_hash("StrongPass1!"), "role": "admin"}],
                f,
            )

        spend_app.app.config.update(TESTING=True)
        self.client = spend_app.app.test_client()

    def tearDown(self):
        spend_app.USERS_FILE = self.old_users_file
        spend_app.DATA_DIR = self.old_data_dir
        spend_app.CLOUD_TOKENS_FILE = self.old_cloud_tokens_file
        self.tmp.cleanup()

    def login(self):
        response = self.client.get("/login")
        token = extract_csrf(response)
        login_response = self.client.post(
            "/login",
            data={"username": "admin", "password": "StrongPass1!", "csrf_token": token},
        )
        self.assertEqual(login_response.status_code, 302)
        return token

    def load_admin_data(self):
        path = os.path.join(spend_app.DATA_DIR, "data_admin.json")
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    def write_admin_data(self, data):
        path = os.path.join(spend_app.DATA_DIR, "data_admin.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)

    def test_json_post_requires_csrf_header(self):
        self.login()
        response = self.client.post("/api/budget-limits", json={"Groceries": 1000})
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.get_json()["error"], "Invalid or missing CSRF token.")

    def test_logout_is_post_only_and_requires_csrf(self):
        token = self.login()
        self.assertEqual(self.client.get("/logout").status_code, 405)
        response = self.client.post("/logout", data={"csrf_token": token})
        self.assertEqual(response.status_code, 302)
        self.assertEqual(self.client.get("/").status_code, 302)

    def test_cloud_connect_is_post_only(self):
        self.login()
        self.assertEqual(self.client.get("/cloud/connect/gdrive").status_code, 405)

    def test_add_expense_rejects_negative_amount(self):
        token = self.login()
        response = self.client.post(
            "/add",
            data={
                "csrf_token": token,
                "amount": "-25",
                "category": "Groceries",
                "subcategory": "Milk",
                "date": spend_app.today_str(),
                "payment_method": "Cash",
            },
        )
        self.assertEqual(response.status_code, 302)
        self.assertEqual(self.load_admin_data()["expenses"], [])

    def test_view_rejects_zero_page_size_without_crashing(self):
        self.login()
        response = self.client.get("/view?per_page=0")
        self.assertEqual(response.status_code, 200)

    def test_primary_pages_render_after_login(self):
        self.login()
        for path in ["/", "/view", "/analytics", "/settings", "/income", "/planning", "/rules", "/import/csv", "/templates"]:
            with self.subTest(path=path):
                response = self.client.get(path)
                self.assertEqual(response.status_code, 200)

    def test_template_quick_add_is_post_only(self):
        token = self.login()
        self.write_admin_data(
            {
                "expenses": [],
                "templates": [
                    {
                        "id": "template-1",
                        "name": "Milk",
                        "category": "Groceries",
                        "subcategory": "Milk",
                        "amount": 55,
                        "payment_method": "Cash",
                    }
                ],
            }
        )

        get_response = self.client.get("/add-from-template/template-1")
        self.assertEqual(get_response.status_code, 405)

        post_response = self.client.post(
            "/add-from-template/template-1",
            data={"csrf_token": token},
        )
        self.assertEqual(post_response.status_code, 302)
        saved = self.load_admin_data()
        self.assertEqual(len(saved["expenses"]), 1)
        self.assertEqual(saved["expenses"][0]["subcategory"], "Milk")


if __name__ == "__main__":
    unittest.main()
