import unittest

from feature_drafts.account_transactions import (
    apply_account_link_to_transaction,
    build_account_link_context,
    build_account_movements,
    build_account_transaction_view_model,
    detect_account_transfers,
    payment_method_display,
    recompute_account_balances,
    resolve_account_id_for_payment_method,
    resolve_transaction_account_id,
    transaction_account_delta,
    transaction_signed_cashflow,
)


def account_fixture():
    return {
        "payment_methods": ["Cash", "HDFC UPI", "Card 1"],
        "accounts": [
            {
                "id": "cash",
                "name": "Cash Wallet",
                "type": "cash",
                "opening_balance": 1000,
                "opening_date": "2026-05-01",
                "payment_methods": ["Cash"],
            },
            {
                "id": "hdfc-bank",
                "name": "HDFC Bank",
                "type": "bank",
                "opening_balance": 5000,
                "opening_date": "2026-05-01",
                "payment_methods": ["HDFC UPI", "HDFC Debit"],
            },
            {
                "id": "card-1",
                "name": "Card 1",
                "type": "credit_card",
                "opening_balance": 0,
                "opening_date": "2026-05-01",
                "payment_methods": ["Card 1"],
            },
        ],
    }


class AccountTransactionHelperTests(unittest.TestCase):
    def test_resolves_account_id_from_existing_payment_method_aliases(self):
        data = account_fixture()

        self.assertEqual(resolve_account_id_for_payment_method("HDFC UPI", data), "hdfc-bank")
        self.assertEqual(
            resolve_transaction_account_id({"payment_method": "Card 1"}, data),
            "card-1",
        )

    def test_payment_method_drafts_cover_unmigrated_payment_methods(self):
        data = {"payment_methods": ["Cash", "Wallet Pay"], "accounts": []}
        context = build_account_link_context(data)

        self.assertEqual(resolve_account_id_for_payment_method("Wallet Pay", data), "wallet-pay")
        self.assertEqual([account["id"] for account in context["accounts"]], ["cash", "wallet-pay"])

    def test_account_link_preserves_existing_payment_method_display(self):
        data = account_fixture()
        transaction = {"amount": 250, "payment_method": "HDFC UPI", "date": "2026-05-20"}

        linked = apply_account_link_to_transaction(transaction, data)

        self.assertEqual(linked["account_id"], "hdfc-bank")
        self.assertEqual(linked["payment_method"], "HDFC UPI")
        self.assertNotIn("account_id", transaction)

    def test_payment_method_display_falls_back_to_account_label(self):
        data = account_fixture()
        transaction = {"amount": 250, "account_id": "hdfc-bank", "date": "2026-05-20"}

        self.assertEqual(payment_method_display(transaction, data), "HDFC UPI")

    def test_view_model_marks_selected_account(self):
        model = build_account_transaction_view_model(
            account_fixture(),
            {"payment_method": "Card 1", "amount": 1200},
        )

        selected = [option for option in model["account_options"] if option["selected"]]
        self.assertEqual(selected[0]["id"], "card-1")
        self.assertEqual(model["payment_method_display"], "Card 1")

    def test_transaction_cashflow_and_liability_delta(self):
        card = account_fixture()["accounts"][2]
        expense = {"amount": 500, "payment_method": "Card 1"}
        payment = {"amount": 200, "payment_method": "Card 1", "direction": "inflow"}

        self.assertEqual(transaction_signed_cashflow(expense), -500)
        self.assertEqual(transaction_account_delta(expense, card), 500)
        self.assertEqual(transaction_account_delta(payment, card), -200)

    def test_detects_paired_transfer_from_payment_method_accounts(self):
        data = account_fixture()
        transactions = [
            {
                "id": "bank-out",
                "date": "2026-05-20",
                "payment_method": "HDFC UPI",
                "amount": 1000,
                "direction": "outflow",
                "notes": "Credit card payment",
            },
            {
                "id": "card-in",
                "date": "2026-05-21",
                "payment_method": "Card 1",
                "amount": 1000,
                "direction": "inflow",
                "notes": "Card payment received",
            },
        ]

        result = detect_account_transfers(transactions, data)

        self.assertEqual(len(result["candidates"]), 1)
        candidate = result["candidates"][0]
        self.assertEqual(candidate["kind"], "paired_transfer")
        self.assertEqual(candidate["confidence"], "high")
        self.assertEqual(candidate["from_account_id"], "hdfc-bank")
        self.assertEqual(candidate["to_account_id"], "card-1")

    def test_detects_explicit_single_row_transfer(self):
        data = account_fixture()
        transaction = {
            "id": "cash-to-bank",
            "date": "2026-05-10",
            "account_id": "cash",
            "to_account_id": "hdfc-bank",
            "amount": 300,
            "notes": "Transfer to bank",
        }

        transfers = detect_account_transfers([transaction], data)
        movements = build_account_movements([transaction], data)["movements"]

        self.assertEqual(transfers["candidates"][0]["kind"], "explicit_transfer")
        self.assertEqual(sum(m["delta"] for m in movements if m["account_id"] == "cash"), -300)
        self.assertEqual(sum(m["delta"] for m in movements if m["account_id"] == "hdfc-bank"), 300)

    def test_detects_single_sided_transfer_keyword_for_review(self):
        result = detect_account_transfers(
            [
                {
                    "id": "wallet-topup",
                    "date": "2026-05-22",
                    "payment_method": "HDFC UPI",
                    "amount": 2000,
                    "notes": "Wallet topup",
                }
            ],
            account_fixture(),
        )

        self.assertEqual(result["candidates"][0]["kind"], "single_sided_transfer_review")

    def test_recomputes_balances_from_opening_balances_and_transactions(self):
        data = account_fixture()
        data["expenses"] = [
            {"id": "grocery", "date": "2026-05-03", "payment_method": "HDFC UPI", "amount": 500},
            {"id": "salary", "date": "2026-05-04", "account_id": "hdfc-bank", "amount": 1500, "direction": "inflow"},
            {"id": "card-spend", "date": "2026-05-05", "payment_method": "Card 1", "amount": 700},
            {"id": "card-pay", "date": "2026-05-06", "payment_method": "Card 1", "amount": 200, "direction": "inflow"},
        ]

        summary = recompute_account_balances(data, as_of="2026-05-31")
        rows = {row["account_id"]: row for row in summary["accounts"]}

        self.assertEqual(rows["hdfc-bank"]["balance"], 6000)
        self.assertEqual(rows["card-1"]["balance"], 500)
        self.assertEqual(rows["card-1"]["signed_balance"], -500)
        self.assertEqual(summary["totals"]["net_worth"], 6500)

    def test_manual_balance_anchor_applies_only_later_transactions(self):
        data = account_fixture()
        data["account_balance_snapshots"] = [
            {
                "id": "hdfc-manual",
                "account_id": "hdfc-bank",
                "date": "2026-05-10",
                "balance": 4500,
                "source": "manual",
            }
        ]
        data["expenses"] = [
            {"id": "before", "date": "2026-05-05", "payment_method": "HDFC UPI", "amount": 500},
            {"id": "after", "date": "2026-05-11", "payment_method": "HDFC UPI", "amount": 250},
        ]

        summary = recompute_account_balances(data, as_of="2026-05-31")
        hdfc = next(row for row in summary["accounts"] if row["account_id"] == "hdfc-bank")

        self.assertEqual(hdfc["anchor_balance"], 4500)
        self.assertEqual(hdfc["anchor_source"], "manual")
        self.assertEqual(hdfc["transaction_delta"], -250)
        self.assertEqual(hdfc["balance"], 4250)

    def test_undated_current_balance_does_not_double_count_history(self):
        data = {
            "payment_methods": ["Cash"],
            "accounts": [{"id": "cash", "name": "Cash", "type": "cash", "balance": 900, "payment_methods": ["Cash"]}],
            "expenses": [{"id": "old", "date": "2026-05-01", "payment_method": "Cash", "amount": 100}],
        }

        summary = recompute_account_balances(data, as_of="2026-05-31")
        cash = summary["accounts"][0]

        self.assertEqual(cash["anchor_source"], "account.balance")
        self.assertEqual(cash["balance"], 900)
        self.assertEqual(cash["transaction_count"], 0)


if __name__ == "__main__":
    unittest.main()
