import unittest
from datetime import date

import app as spend_app


class MarketFeatureHelperTests(unittest.TestCase):
    def test_accounts_summary_separates_assets_liabilities_and_delta(self):
        data = {
            "accounts": [
                {"id": "bank", "name": "Bank", "type": "bank", "balance": 100000, "include_in_net_worth": True},
                {"id": "card", "name": "Card", "type": "credit_card", "balance": 12000, "include_in_net_worth": True},
                {"id": "old", "name": "Old", "type": "wallet", "balance": 500, "include_in_net_worth": False},
            ],
            "net_worth_snapshots": [{"date": "2026-05-01", "net_worth": 80000}],
        }

        summary = spend_app.build_accounts_summary(data)

        self.assertEqual(summary["assets"], 100000.0)
        self.assertEqual(summary["liabilities"], 12000.0)
        self.assertEqual(summary["net_worth"], 88000.0)
        self.assertEqual(summary["net_worth_delta_since_snapshot"], 8000.0)

    def test_envelope_budget_tracks_spend_rollover_and_left_to_assign(self):
        data = {
            "income": {"monthly_salary": 50000, "salary_history": [{"amount": 50000, "effective_from": "2026-01-01"}]},
            "extra_income": [],
            "fixed_expenses": [],
            "goals": [{"monthly_contribution": 5000}],
            "budget_envelopes": [
                {"id": "prev", "month": "2026-04", "category": "Groceries", "assigned": 10000, "annual_amount": 0, "rollover_enabled": True},
                {"id": "curr", "month": "2026-05", "category": "Groceries", "assigned": 12000, "annual_amount": 12000, "rollover_enabled": True},
            ],
            "expenses": [
                {"date": "2026-04-10", "category": "Groceries", "amount": 7000},
                {"date": "2026-05-10", "category": "Groceries", "amount": 5000},
            ],
            "budget_limits": {},
            "custom_categories": {},
        }

        budget = spend_app.build_envelope_budget(data, "2026-05")
        row = next(r for r in budget["rows"] if r["category"] == "Groceries")

        self.assertEqual(row["rollover"], 3000.0)
        self.assertEqual(row["annual_monthly"], 1000.0)
        self.assertEqual(row["available"], 16000.0)
        self.assertEqual(row["remaining"], 11000.0)
        self.assertEqual(budget["left_to_assign"], 32000.0)

    def test_review_inbox_marks_duplicate_candidates(self):
        data = {
            "expenses": [
                {"id": "a", "date": "2026-05-01", "amount": 499, "subcategory": "Netflix", "notes": "", "review_status": "reviewed"},
                {"id": "b", "date": "2026-05-01", "amount": 499, "subcategory": "Netflix", "notes": "", "review_status": "needs_review", "source": "csv_import"},
            ]
        }

        inbox = spend_app.build_review_inbox(data)

        self.assertEqual(inbox["pending_count"], 1)
        self.assertEqual(inbox["duplicate_count"], 1)
        self.assertTrue(inbox["pending"][0]["is_duplicate_candidate"])

    def test_receipt_summary_links_receipts_and_duplicate_candidates(self):
        data = {
            "receipts": [
                {
                    "id": "r1",
                    "original_filename": "receipt.png",
                    "stored_filename": "r1.png",
                    "uploaded_at": "2026-05-31T10:00:00",
                    "status": "expense_created",
                    "expense_id": "e1",
                    "extracted": {"date": "2026-05-31", "amount": "100", "merchant": "Cafe"},
                }
            ],
            "expenses": [
                {"id": "e1", "receipt_id": "r1", "date": "2026-05-31", "amount": 100, "subcategory": "Cafe", "notes": "Cafe"}
            ],
        }

        summary = spend_app.build_receipt_summary(data)

        self.assertEqual(summary["receipt_count"], 1)
        self.assertEqual(summary["linked_count"], 1)
        self.assertEqual(summary["duplicate_count"], 1)

    def test_account_selection_sets_account_id_and_preserves_payment_display(self):
        data = {
            "accounts": [
                {"id": "acc-card", "name": "Visa Card", "type": "credit_card", "balance": 1000},
            ],
            "payment_methods": ["Cash"],
        }
        expense = {"amount": 100, "payment_method": "Cash"}

        spend_app.apply_account_selection(expense, data, "acc-card", "Cash")

        self.assertEqual(expense["account_id"], "acc-card")
        self.assertEqual(expense["payment_method"], "Visa Card")
        self.assertIn("Visa Card", data["payment_methods"])

    def test_receipt_ocr_parser_is_wired_for_indian_receipts(self):
        text = """
        DMART READY
        Date: 31/05/2026
        Grand Total Rs. 1,234.50
        Paid by UPI
        """

        extracted = spend_app.parse_receipt_ocr_text(text, {"expenses": []}, ref_date=date(2026, 5, 31))

        self.assertEqual(extracted["merchant"], "DMART READY")
        self.assertEqual(extracted["date"], "2026-05-31")
        self.assertEqual(extracted["amount"], 1234.50)
        self.assertEqual(extracted["payment_method"], "UPI")

    def test_recurring_calendar_combines_fixed_subscriptions_and_goals(self):
        data = {
            "income": {"monthly_salary": 50000, "salary_history": [{"amount": 50000, "effective_from": "2026-01-01"}]},
            "extra_income": [],
            "fixed_expenses": [
                {
                    "id": "rent",
                    "name": "Rent",
                    "amount": 15000,
                    "type": "fixed",
                    "frequency": "monthly",
                    "day_of_month": 5,
                    "start_year": 2026,
                    "start_month": 1,
                    "category": "EMI / Finance",
                    "payment_method": "Bank",
                }
            ],
            "goals": [{"id": "goal", "name": "Emergency", "monthly_contribution": 5000}],
            "expenses": [
                {"id": "n1", "amount": 499, "category": "Entertainment", "subcategory": "Netflix", "payment_method": "Card", "date": "2026-03-15"},
                {"id": "n2", "amount": 499, "category": "Entertainment", "subcategory": "Netflix", "payment_method": "Card", "date": "2026-04-15"},
                {"id": "n3", "amount": 499, "category": "Entertainment", "subcategory": "Netflix", "payment_method": "Card", "date": "2026-05-15"},
            ],
            "recurring_payments": [],
        }

        calendar = spend_app.build_recurring_calendar(data, "2026-06", ref_today=date(2026, 5, 31))
        names = {item["name"] for item in calendar["items"]}

        self.assertIn("Rent", names)
        self.assertIn("Netflix", names)
        self.assertIn("Emergency contribution", names)
        self.assertEqual(calendar["total_due"], 20499.0)
        self.assertEqual(calendar["cash_after_unpaid"], 29501.0)


if __name__ == "__main__":
    unittest.main()
