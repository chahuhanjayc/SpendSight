import unittest
from datetime import date

import app as spend_app


def require_helper(name):
    helper = getattr(spend_app, name, None)
    if helper is None:
        raise AssertionError(f"Expected app.{name}(data, ref_today=None) to be available.")
    return helper


def subscription_fixture():
    return {
        "expenses": [
            {
                "id": "netflix-jan",
                "amount": 499.0,
                "category": "Entertainment",
                "subcategory": "Netflix",
                "payment_method": "Credit Card",
                "date": "2026-01-05",
                "notes": "Netflix Premium",
            },
            {
                "id": "netflix-feb",
                "amount": 499.0,
                "category": "Entertainment",
                "subcategory": "Netflix",
                "payment_method": "Credit Card",
                "date": "2026-02-05",
                "notes": "Netflix Premium",
            },
            {
                "id": "netflix-mar",
                "amount": 499.0,
                "category": "Entertainment",
                "subcategory": "Netflix",
                "payment_method": "Credit Card",
                "date": "2026-03-05",
                "notes": "Netflix Premium",
            },
            {
                "id": "netflix-apr",
                "amount": 549.0,
                "category": "Entertainment",
                "subcategory": "Netflix",
                "payment_method": "Credit Card",
                "date": "2026-04-05",
                "notes": "Netflix Premium",
            },
            {
                "id": "netflix-may",
                "amount": 549.0,
                "category": "Entertainment",
                "subcategory": "Netflix",
                "payment_method": "Credit Card",
                "date": "2026-05-05",
                "notes": "Netflix Premium",
            },
            {
                "id": "one-off",
                "amount": 2100.0,
                "category": "Shopping",
                "subcategory": "Headphones",
                "payment_method": "UPI",
                "date": "2026-05-18",
                "notes": "One-time purchase",
            },
        ]
    }


def get_subscriptions(result):
    if isinstance(result, dict):
        return result.get("subscriptions", result.get("items", []))
    return result


def find_subscription(result, name):
    subscriptions = get_subscriptions(result)
    if not isinstance(subscriptions, list):
        raise AssertionError("Subscription insights should expose a list of subscriptions.")

    for item in subscriptions:
        label = str(item.get("name") or item.get("merchant") or item.get("label") or "")
        if label.lower() == name.lower():
            return item
    raise AssertionError(f"Subscription named {name!r} was not found in {subscriptions!r}.")


class ProductFeatureHelperTests(unittest.TestCase):
    def test_detects_monthly_subscription_from_repeated_expenses(self):
        build_subscription_insights = require_helper("build_subscription_insights")

        result = build_subscription_insights(subscription_fixture(), ref_today=date(2026, 5, 31))
        netflix = find_subscription(result, "Netflix")

        self.assertEqual(netflix["name"], "Netflix")
        self.assertEqual(netflix["category"], "Entertainment")
        self.assertEqual(netflix["cadence"], "monthly")
        self.assertEqual(netflix["occurrences"], 5)
        self.assertAlmostEqual(netflix["latest_amount"], 549.0)

    def test_subscription_insight_reports_price_change_and_next_due(self):
        build_subscription_insights = require_helper("build_subscription_insights")

        result = build_subscription_insights(subscription_fixture(), ref_today=date(2026, 5, 31))
        netflix = find_subscription(result, "Netflix")

        self.assertEqual(netflix["next_due_on"], "2026-06-05")
        self.assertIn("price_change", netflix)
        self.assertEqual(
            netflix["price_change"],
            {
                "from_amount": 499.0,
                "to_amount": 549.0,
                "delta": 50.0,
                "direction": "increase",
            },
        )

    def test_goals_summary_tracks_progress_and_contribution(self):
        get_goals_summary = require_helper("get_goals_summary")
        data = {
            "goals": [
                {
                    "id": "emergency",
                    "name": "Emergency Fund",
                    "target_amount": 100000.0,
                    "current_amount": 25000.0,
                    "monthly_contribution": 7500.0,
                },
                {
                    "id": "travel",
                    "name": "Travel Fund",
                    "target_amount": 50000.0,
                    "current_amount": 10000.0,
                    "monthly_contribution": 5000.0,
                },
            ]
        }

        summary = get_goals_summary(data, ref_today=date(2026, 5, 31))
        emergency = next(g for g in summary["goals"] if g["name"] == "Emergency Fund")

        self.assertEqual(summary["total_target"], 150000.0)
        self.assertEqual(summary["total_current"], 35000.0)
        self.assertEqual(summary["total_monthly_contribution"], 12500.0)
        self.assertAlmostEqual(summary["overall_progress_pct"], 23.3, places=1)

        self.assertEqual(emergency["target_amount"], 100000.0)
        self.assertEqual(emergency["current_amount"], 25000.0)
        self.assertEqual(emergency["monthly_contribution"], 7500.0)
        self.assertEqual(emergency["remaining_amount"], 75000.0)
        self.assertEqual(emergency["progress_pct"], 25.0)
        self.assertEqual(emergency["months_to_target"], 10)

    def test_transaction_rules_apply_category_subcategory_and_payment(self):
        helper = require_helper("apply_transaction_rules")
        data = {
            "custom_categories": {},
            "payment_methods": ["Cash"],
            "transaction_rules": [
                {
                    "name": "Netflix",
                    "match_text": "netflix",
                    "min_amount": 400,
                    "max_amount": 700,
                    "set_category": "Entertainment",
                    "set_subcategory": "Netflix",
                    "set_payment_method": "Credit Card",
                    "enabled": True,
                }
            ],
        }
        expense = {"amount": 549, "category": "Other", "subcategory": "", "payment_method": "Cash", "notes": "NETFLIX PREMIUM"}

        updated = helper(expense, data)

        self.assertEqual(updated["category"], "Entertainment")
        self.assertEqual(updated["subcategory"], "Netflix")
        self.assertEqual(updated["payment_method"], "Credit Card")
        self.assertEqual(updated["rules_applied"], ["Netflix"])
        self.assertIn("Netflix", data["custom_categories"]["Entertainment"])

    def test_csv_import_deduplicates_and_applies_rules(self):
        parser = require_helper("parse_expense_csv")
        data = {
            "expenses": [
                {
                    "date": "2026-05-05",
                    "amount": 549.0,
                    "subcategory": "Netflix",
                    "notes": "Netflix Premium",
                }
            ],
            "custom_categories": {},
            "transaction_rules": [
                {
                    "name": "Netflix",
                    "match_text": "netflix",
                    "set_category": "Entertainment",
                    "set_subcategory": "Netflix",
                    "set_payment_method": "Credit Card",
                    "enabled": True,
                }
            ],
        }
        csv_text = "\n".join([
            "date,amount,description,account",
            "2026-05-05,549,Netflix Premium,Card",
            "2026-06-05,549,Netflix Premium,Card",
        ])

        result = parser(csv_text, data)

        self.assertEqual(result["imported_count"], 1)
        self.assertEqual(result["skipped_count"], 1)
        imported = result["expenses"][0]
        self.assertEqual(imported["date"], "2026-06-05")
        self.assertEqual(imported["category"], "Entertainment")
        self.assertEqual(imported["subcategory"], "Netflix")
        self.assertEqual(imported["payment_method"], "Credit Card")


if __name__ == "__main__":
    unittest.main()
