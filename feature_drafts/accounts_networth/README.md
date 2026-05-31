# Accounts & Net Worth Draft

This draft is isolated from the Flask app. It adds pure helpers and a Jinja template candidate without changing `app.py` or existing templates.

## Proposed Additive JSON Keys

- `accounts`: account records for cash, bank, wallet, credit card, loan, investment, and other asset/liability accounts.
- `account_balance_snapshots`: dated balances per account.
- `net_worth_snapshots`: optional stored computed snapshots for trend/history views.

Liability account balances are stored as positive amounts owed. Net worth subtracts liability balances.

## Parent Wiring Sketch

1. Import helpers in `app.py` after review:

   ```python
   from feature_drafts.accounts_networth import (
       build_account_balances,
       build_net_worth_snapshot,
       detect_transfer_candidates,
       migrate_payment_methods_to_accounts,
       normalize_accounts_data,
       append_balance_snapshot,
   )
   ```

2. Extend `_default_data()` and `_normalize_data()` with additive defaults:

   ```python
   "accounts": [],
   "account_balance_snapshots": [],
   "net_worth_snapshots": [],
   ```

   Then call `normalize_accounts_data(d)` near the end of `_normalize_data()` once the parent is ready to enforce account validation.

3. Add a route without changing existing behavior:

   ```python
   @app.route("/accounts")
   @login_required
   def accounts_networth():
       data = migrate_payment_methods_to_accounts(load_data(), opening_date=today_str())
       net_worth = build_net_worth_snapshot(data)
       accounts = build_account_balances(data)
       transfers = detect_transfer_candidates(data.get("expenses", []), data)
       return render_spendsight_template(
           "accounts_networth.html",
           net_worth=net_worth,
           accounts=accounts,
           asset_count=sum(1 for a in accounts if not a["is_liability"]),
           liability_count=sum(1 for a in accounts if a["is_liability"]),
           transfer_candidates=transfers["candidates"],
       )
   ```

4. Copy or move `feature_drafts/accounts_networth/templates/accounts_networth.html` into `templates/accounts_networth.html` when adopting the feature.

5. Add a sidebar item in `templates/base.html` under Insights:

   ```html
   <a href="{{ url_for('accounts_networth') }}" class="{{ 'active' if request.endpoint == 'accounts_networth' }}">
     <i class="bi bi-wallet2"></i> Accounts
   </a>
   ```

## Notes

- Existing `payment_methods` remain supported. `migrate_payment_methods_to_accounts()` creates draft accounts with aliases so current expenses can map forward gradually.
- Transfer detection is intentionally conservative. Paired opposite-signed transactions can be high confidence; keyword-only matches are review items.
- The helpers do not read files, write files, call Flask, or mutate the input data dict.
