import unittest

from feature_drafts.accounts_networth import (
    append_balance_snapshot,
    build_account_balances,
    build_net_worth_series,
    build_net_worth_snapshot,
    detect_transfer_candidates,
    migrate_payment_methods_to_accounts,
    new_account,
    normalize_accounts_data,
    validate_accounts_data,
)


class AccountsNetWorthTests(unittest.TestCase):
    def account_fixture(self):
        return {
            "currency_code": "INR",
            "payment_methods": ["Cash", "Card 1"],
            "expenses": [],
            "accounts": [
                new_account(
                    "Savings",
                    "bank",
                    account_id="savings",
                    opening_balance=10000,
                    opening_date="2026-05-01",
                    payment_methods=["Bank"],
                ),
                new_account(
                    "Card 1",
                    "credit_card",
                    account_id="card-1",
                    opening_balance=5000,
                    opening_date="2026-05-01",
                ),
                new_account(
                    "Brokerage",
                    "investment",
                    account_id="brokerage",
                    opening_balance=25000,
                    opening_date="2026-05-01",
                ),
            ],
            "account_balance_snapshots": [
                {
                    "id": "savings-2026-05-15",
                    "account_id": "savings",
                    "date": "2026-05-15",
                    "balance": 12000,
                    "source": "manual",
                },
                {
                    "id": "card-2026-05-15",
                    "account_id": "card-1",
                    "date": "2026-05-15",
                    "balance": 4500,
                    "source": "manual",
                },
            ],
        }

    def test_migrates_flat_payment_methods_to_account_aliases(self):
        data = {
            "currency_code": "INR",
            "payment_methods": ["Cash", "Card 1", "HDFC Bank"],
            "expenses": [{"amount": 100, "date": "2026-05-31", "payment_method": "Card 1"}],
        }

        migrated = migrate_payment_methods_to_accounts(data, opening_date="2026-05-31")
        accounts = {account["name"]: account for account in migrated["accounts"]}

        self.assertEqual(accounts["Cash"]["type"], "cash")
        self.assertEqual(accounts["Card 1"]["type"], "credit_card")
        self.assertEqual(accounts["Card 1"]["is_liability"], True)
        self.assertIn("HDFC Bank", accounts)

    def test_net_worth_snapshot_subtracts_liabilities_and_uses_latest_balance(self):
        snapshot = build_net_worth_snapshot(self.account_fixture(), as_of="2026-05-31")

        self.assertEqual(snapshot["assets_total"], 37000.0)
        self.assertEqual(snapshot["liabilities_total"], 4500.0)
        self.assertEqual(snapshot["net_worth"], 32500.0)

        account_rows = {row["account_id"]: row for row in snapshot["accounts"]}
        self.assertEqual(account_rows["savings"]["balance"], 12000.0)
        self.assertEqual(account_rows["card-1"]["signed_balance"], -4500.0)

    def test_append_balance_snapshot_returns_copy(self):
        data = self.account_fixture()

        updated = append_balance_snapshot(
            data,
            "savings",
            15000,
            snapshot_date="2026-05-31",
            source="manual",
        )

        self.assertEqual(len(data["account_balance_snapshots"]), 2)
        self.assertEqual(len(updated["account_balance_snapshots"]), 3)
        balances = build_account_balances(updated, as_of="2026-05-31")
        savings = next(row for row in balances if row["account_id"] == "savings")
        self.assertEqual(savings["balance"], 15000.0)

    def test_net_worth_series_uses_snapshot_dates(self):
        series = build_net_worth_series(self.account_fixture(), end_date="2026-05-31")
        dates = [point["date"] for point in series]

        self.assertIn("2026-05-01", dates)
        self.assertIn("2026-05-15", dates)
        self.assertIn("2026-05-31", dates)
        self.assertEqual(series[-1]["net_worth"], 32500.0)

    def test_transfer_detection_pairs_opposite_signed_transactions(self):
        data = normalize_accounts_data(self.account_fixture())
        transactions = [
            {
                "id": "bank-out",
                "date": "2026-05-20",
                "account_id": "savings",
                "amount": 10000,
                "direction": "outflow",
                "notes": "Credit card payment",
            },
            {
                "id": "card-in",
                "date": "2026-05-21",
                "account_id": "card-1",
                "amount": 10000,
                "direction": "inflow",
                "notes": "Card payment received",
            },
        ]

        result = detect_transfer_candidates(transactions, data)

        self.assertEqual(len(result["candidates"]), 1)
        candidate = result["candidates"][0]
        self.assertEqual(candidate["kind"], "paired_transfer")
        self.assertEqual(candidate["confidence"], "high")
        self.assertEqual(candidate["from_account_id"], "savings")
        self.assertEqual(candidate["to_account_id"], "card-1")

    def test_transfer_detection_flags_single_sided_keyword_for_review(self):
        result = detect_transfer_candidates(
            [
                {
                    "id": "expense-transfer",
                    "date": "2026-05-22",
                    "payment_method": "Cash",
                    "amount": 2000,
                    "notes": "Wallet topup",
                }
            ],
            {"accounts": [new_account("Cash", "cash", account_id="cash", opening_date="2026-05-01")]},
        )

        self.assertEqual(result["candidates"][0]["kind"], "single_sided_transfer_review")
        self.assertEqual(result["candidates"][0]["confidence"], "low")

    def test_validation_reports_bad_snapshot_reference(self):
        data = {
            "accounts": [new_account("Cash", "cash", account_id="cash", opening_date="2026-05-01")],
            "account_balance_snapshots": [
                {"id": "missing-1", "account_id": "missing", "date": "2026-05-31", "balance": 50}
            ],
        }

        result = validate_accounts_data(data)

        self.assertFalse(result["valid"])
        self.assertIn("unknown account", result["errors"][0])


if __name__ == "__main__":
    unittest.main()
