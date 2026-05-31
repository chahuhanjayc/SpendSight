# Account Transactions Draft

This draft stays outside the live Flask app. It adds pure helpers for gradually
moving transaction entry from flat `payment_method` values toward account-linked
transactions while preserving existing reports and templates.

## Scope

- Resolve `account_id` from current `payment_method` values and account aliases.
- Keep `payment_method` populated as the backward-compatible display label.
- Detect explicit, paired, and keyword-only transfer candidates for review.
- Recompute account balances from opening/current/manual balance anchors plus
  account-linked transactions.

## Parent Wiring

Import only after review:

```python
from feature_drafts.account_transactions import (
    apply_account_link_to_transaction,
    build_account_transaction_view_model,
    detect_account_transfers,
    recompute_account_balances,
)
```

Expense add/edit GET:

```python
account_tx = build_account_transaction_view_model(data, expense if editing else None)
return render_spendsight_template(
    "edit_expense.html",
    expense=expense,
    payment_methods=data.get("payment_methods", DEFAULT_PAYMENT_METHODS),
    account_tx=account_tx,
)
```

Expense add/edit POST:

```python
expense = {
    "amount": amount,
    "category": category,
    "subcategory": subcategory,
    "date": tx_date,
    "payment_method": request.form.get("payment_method", "Cash"),
    "notes": notes,
}
expense = apply_account_link_to_transaction(
    expense,
    data,
    account_id=request.form.get("account_id", ""),
)
```

The helper preserves a non-empty `payment_method`; existing dashboards continue
to render while new templates can use `expense.account_id`.

Accounts summary route:

```python
summary = recompute_account_balances(data)
transfers = detect_account_transfers(data)
return render_spendsight_template(
    "accounts.html",
    accounts=summary["accounts"],
    summary=summary,
    transfer_candidates=transfers["candidates"],
)
```

Optional template change when adopting the UX: add an `account_id` select beside
the existing payment method field, using `account_tx.account_options`. Keep the
old `payment_method` input or hidden field until all reports read account labels
through the helper.
