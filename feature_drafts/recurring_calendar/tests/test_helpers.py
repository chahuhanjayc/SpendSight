import unittest
from datetime import date

from feature_drafts.recurring_calendar import (
    build_recurring_calendar,
    detect_monthly_subscriptions,
    fixed_expense_occurrences,
)


def item_named(result, name, source_type=None, due_on=None):
    for item in result["items"]:
        if item["name"] != name:
            continue
        if source_type and item["source_type"] != source_type:
            continue
        if due_on and item["due_on"] != due_on:
            continue
        return item
    raise AssertionError(f"Missing item {name!r}: {result['items']!r}")


def month_row(result, month):
    for row in result["calendar_rows"]:
        if row["month"] == month:
            return row
    raise AssertionError(f"Missing month {month!r}: {result['calendar_rows']!r}")


def netflix_expenses():
    return [
        {
            "id": f"netflix-{month}",
            "amount": amount,
            "category": "Entertainment",
            "subcategory": "Netflix",
            "payment_method": "Credit Card",
            "date": f"2026-{month:02d}-05",
            "notes": "Netflix Premium",
        }
        for month, amount in ((1, 499), (2, 499), (3, 499), (4, 549), (5, 549))
    ]


class RecurringCalendarHelperTests(unittest.TestCase):
    def test_fixed_expense_rows_include_paid_status_and_cash_after_bills(self):
        data = {
            "income": {"monthly_salary": 50000},
            "fixed_expenses": [
                {
                    "id": "rent",
                    "name": "Apartment Rent",
                    "type": "fixed",
                    "frequency": "monthly",
                    "amount": 20000,
                    "start_year": 2026,
                    "start_month": 1,
                    "day_of_month": 5,
                    "category": "Rent",
                    "payment_method": "Bank",
                }
            ],
            "expenses": [
                {
                    "id": "rent-june",
                    "amount": 20000,
                    "category": "Rent",
                    "subcategory": "Apartment Rent",
                    "payment_method": "Bank",
                    "date": "2026-06-05",
                    "notes": "June apartment rent",
                }
            ],
        }

        result = build_recurring_calendar(data, ref_today=date(2026, 6, 1), months_ahead=1)
        rent = item_named(result, "Apartment Rent", "fixed_expense", "2026-06-05")
        june = month_row(result, "2026-06")

        self.assertEqual(rent["status"], "paid")
        self.assertEqual(rent["status_assumption"], "matched_expense")
        self.assertEqual(result["summary"]["paid_total"], 20000)
        self.assertEqual(result["summary"]["unpaid_total"], 0)
        self.assertEqual(june["scheduled_total"], 20000)
        self.assertEqual(june["cash_after_bills"], 30000)
        self.assertEqual(june["cash_after_unpaid_bills"], 50000)

    def test_fixed_expense_occurrences_respect_emi_and_frequency_bounds(self):
        fixed = [
            {
                "id": "loan",
                "name": "Loan EMI",
                "type": "emi",
                "amount": 8000,
                "start_year": 2026,
                "start_month": 5,
                "total_months": 2,
            },
            {
                "id": "insurance",
                "name": "Insurance",
                "type": "fixed",
                "frequency": "yearly",
                "amount": 12000,
                "start_year": 2026,
                "start_month": 6,
                "day_of_month": 15,
            },
        ]

        occurrences = fixed_expense_occurrences(fixed, date(2026, 6, 1), date(2026, 7, 31))

        self.assertEqual([item["name"] for item in occurrences], ["Loan EMI", "Insurance"])
        self.assertEqual([item["due_on"] for item in occurrences], ["2026-06-01", "2026-06-15"])

    def test_detected_subscriptions_generate_upcoming_unpaid_items(self):
        data = {"expenses": netflix_expenses(), "income": {"monthly_salary": 10000}}

        insights = detect_monthly_subscriptions(data, ref_today=date(2026, 5, 31))
        result = build_recurring_calendar(data, ref_today=date(2026, 5, 31), months_ahead=2)
        june = item_named(result, "Netflix", "subscription", "2026-06-05")
        july = item_named(result, "Netflix", "subscription", "2026-07-05")

        self.assertEqual(insights["count"], 1)
        self.assertEqual(june["amount"], 549)
        self.assertEqual(june["status"], "unpaid")
        self.assertEqual(june["status_assumption"], "future_due_unpaid")
        self.assertEqual(july["amount"], 549)
        self.assertEqual(result["summary"]["source_counts"]["subscription"], 2)
        self.assertEqual(month_row(result, "2026-06")["cash_after_bills"], 9451)

    def test_goals_monthly_contributions_are_scheduled_and_reduce_cash(self):
        data = {
            "income": {"monthly_salary": 12000},
            "recurring_calendar": {"goal_contribution_day": 10},
            "goals": [
                {
                    "id": "emergency",
                    "name": "Emergency Fund",
                    "target_amount": 10000,
                    "current_amount": 7000,
                    "monthly_contribution": 2000,
                }
            ],
        }

        result = build_recurring_calendar(data, ref_today=date(2026, 6, 1), months_ahead=2)
        june = item_named(result, "Emergency Fund", "goal_contribution", "2026-06-10")
        july = item_named(result, "Emergency Fund", "goal_contribution", "2026-07-10")

        self.assertEqual(june["amount"], 2000)
        self.assertEqual(june["status"], "unpaid")
        self.assertEqual(june["status_assumption"], "goal_contribution_assumed_unpaid")
        self.assertEqual(july["amount"], 1000)
        self.assertEqual(month_row(result, "2026-06")["cash_after_bills"], 10000)
        self.assertEqual(month_row(result, "2026-07")["cash_after_bills"], 11000)

    def test_trial_and_renewal_metadata_are_included_when_present(self):
        data = {
            "income": {"monthly_salary": 5000},
            "recurring_calendar": {
                "trial_renewals": [
                    {
                        "id": "canva",
                        "name": "Canva Pro",
                        "category": "Software",
                        "trial_ends_on": "2026-06-14",
                        "renewal_amount": 499,
                        "cancel_by": "2026-06-13",
                    }
                ]
            },
            "renewals": [
                {
                    "id": "domain",
                    "name": "Domain Name",
                    "category": "Business",
                    "date": "2026-06-20",
                    "amount": 1200,
                }
            ],
        }

        result = build_recurring_calendar(data, ref_today=date(2026, 6, 1), months_ahead=1)
        trial = item_named(result, "Canva Pro trial ends", "trial", "2026-06-14")
        renewal = item_named(result, "Domain Name renewal", "renewal", "2026-06-20")

        self.assertEqual(trial["amount"], 499)
        self.assertEqual(trial["metadata"]["cancel_by"], "2026-06-13")
        self.assertEqual(renewal["amount"], 1200)
        self.assertEqual(month_row(result, "2026-06")["unpaid_total"], 1699)

    def test_month_horizon_is_bounded(self):
        data = {
            "fixed_expenses": [
                {
                    "id": "rent",
                    "name": "Rent",
                    "type": "fixed",
                    "frequency": "monthly",
                    "amount": 1,
                    "start_year": 2026,
                    "start_month": 1,
                }
            ]
        }

        result = build_recurring_calendar(data, ref_today=date(2026, 1, 1), months_ahead=999)

        self.assertEqual(result["range"]["months_ahead"], 24)
        self.assertEqual(result["range"]["end"], "2027-12-31")
        self.assertEqual(result["summary"]["source_counts"]["fixed_expense"], 24)


if __name__ == "__main__":
    unittest.main()
