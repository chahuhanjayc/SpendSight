import unittest
from datetime import date

from feature_drafts.envelope_budget import (
    build_envelope_budget,
    build_monthly_assignment_update,
    fixed_expense_occurrences,
)


def row_for(result, category):
    for row in result["rows"]:
        if row["category"] == category:
            return row
    raise AssertionError(f"Missing row for {category!r}: {result['rows']!r}")


class EnvelopeBudgetHelperTests(unittest.TestCase):
    def test_budget_limits_drive_monthly_envelopes_and_left_to_assign(self):
        data = {
            "billing_start_day": 1,
            "income": {"monthly_salary": 10000},
            "budget_limits": {"Groceries": 3000, "Fuel": 1000},
            "expenses": [
                {"date": "2026-05-04", "category": "Groceries", "amount": 1200},
            ],
        }

        result = build_envelope_budget(data, ref_today=date(2026, 5, 15))
        groceries = row_for(result, "Groceries")

        self.assertEqual(result["summary"]["assigned"], 4000)
        self.assertEqual(result["summary"]["left_to_assign"], 6000)
        self.assertEqual(groceries["assigned"], 3000)
        self.assertEqual(groceries["spent"], 1200)
        self.assertEqual(groceries["remaining"], 1800)
        self.assertEqual(groceries["pct_spent"], 40.0)

    def test_rollover_carries_prior_month_surplus_and_shortfall(self):
        data = {
            "billing_start_day": 1,
            "income": {"monthly_salary": 5000},
            "expenses": [
                {"date": "2026-01-10", "category": "Groceries", "amount": 200},
                {"date": "2026-02-10", "category": "Groceries", "amount": 1300},
                {"date": "2026-03-10", "category": "Groceries", "amount": 400},
            ],
            "envelope_budget": {
                "start_month": "2026-01",
                "category_settings": {
                    "Groceries": {"monthly_budget": 1000, "rollover": True}
                },
            },
        }

        result = build_envelope_budget(data, ref_today=date(2026, 3, 15))
        groceries = row_for(result, "Groceries")

        self.assertEqual(groceries["balance_start"], 500)
        self.assertEqual(groceries["assigned"], 1000)
        self.assertEqual(groceries["spent"], 400)
        self.assertEqual(groceries["remaining"], 1100)

    def test_annual_setting_creates_rollover_sinking_fund(self):
        data = {
            "billing_start_day": 1,
            "income": {"monthly_salary": 50000},
            "expenses": [],
            "envelope_budget": {
                "start_month": "2026-01",
                "category_settings": {
                    "Insurance": {
                        "annual_amount": 12000,
                        "due_month": 12,
                    }
                },
            },
        }

        result = build_envelope_budget(data, ref_today=date(2026, 5, 15))
        insurance = row_for(result, "Insurance")

        self.assertTrue(insurance["rollover"])
        self.assertEqual(insurance["sinking_fund_amount"], 1000)
        self.assertEqual(insurance["balance_start"], 4000)
        self.assertEqual(insurance["assigned"], 1000)
        self.assertEqual(insurance["remaining"], 5000)
        self.assertEqual(result["sinking_funds"][0]["category"], "Insurance")

    def test_fixed_annual_expense_can_auto_create_sinking_fund(self):
        data = {
            "billing_start_day": 1,
            "income": {"monthly_salary": 50000},
            "fixed_expenses": [
                {
                    "id": "school-fee",
                    "name": "School Fee",
                    "type": "fixed",
                    "frequency": "yearly",
                    "amount": 24000,
                    "start_year": 2026,
                    "start_month": 7,
                    "day_of_month": 5,
                    "category": "Education",
                }
            ],
        }

        result = build_envelope_budget(data, ref_today=date(2026, 5, 15))
        education = row_for(result, "Education")

        self.assertTrue(education["rollover"])
        self.assertEqual(education["assigned"], 2000)
        self.assertEqual(education["sinking_fund"]["kind"], "scheduled-fixed")
        self.assertEqual(education["spent"], 0)

    def test_fixed_expense_occurrences_respect_frequency(self):
        fixed = [
            {
                "id": "insurance",
                "name": "Insurance",
                "type": "fixed",
                "frequency": "yearly",
                "amount": 12000,
                "start_year": 2026,
                "start_month": 5,
                "day_of_month": 10,
                "category": "Insurance",
            },
            {
                "id": "rent",
                "name": "Rent",
                "type": "fixed",
                "frequency": "monthly",
                "amount": 20000,
                "start_year": 2026,
                "start_month": 1,
                "day_of_month": 1,
                "category": "Rent",
            },
        ]

        may = fixed_expense_occurrences(fixed, date(2026, 5, 1), date(2026, 5, 31))
        june = fixed_expense_occurrences(fixed, date(2026, 6, 1), date(2026, 6, 30))

        self.assertEqual({item["category"] for item in may}, {"Insurance", "Rent"})
        self.assertEqual({item["category"] for item in june}, {"Rent"})

    def test_monthly_assignment_update_returns_merged_config_without_mutating_input(self):
        data = {
            "envelope_budget": {
                "start_month": "2026-05",
                "monthly_assignments": {"2026-05": {"Groceries": 1000}},
            }
        }

        updated = build_monthly_assignment_update(
            data,
            {"Groceries": "1,500", "Fuel": "800", "Empty": ""},
            period_key="2026-06",
        )

        self.assertEqual(updated["monthly_assignments"]["2026-06"], {"Groceries": 1500, "Fuel": 800})
        self.assertNotIn("2026-06", data["envelope_budget"]["monthly_assignments"])


if __name__ == "__main__":
    unittest.main()
