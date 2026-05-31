# SpendSight Product Backlog

This backlog is based on the current codebase and market patterns from YNAB, Monarch Money, PocketGuard, Rocket Money, Goodbudget, and similar personal finance apps.

## Implemented in this local pass

- Security and data safety: CSRF, POST-only mutations, safer cloud restore, per-user cloud tokens, atomic JSON writes, corrupt JSON quarantine, removal of leaked archive artifacts.
- Planning: savings goals with target, current balance, monthly contribution, progress, and projected completion.
- Recurring intelligence: monthly subscription detection, next due date, monthly total, and price-change alerting.
- Automation: transaction rules for category, subcategory, and payment method cleanup.
- Import: CSV statement import with duplicate detection and rule application.
- Accounts and net worth: account ledger, liabilities, net-worth summary, and manual snapshots.
- Envelope budgeting: month-specific category envelopes, annual set-asides, rollover, left-to-assign, and safe-to-spend.
- Review workflow: imported/OCR/receipt-created transactions now land in a review inbox before approval.
- Receipt vault: receipt uploads, extracted field metadata, duplicate signals, and create-expense flow.
- PWA shell: manifest and service worker for installable/offline-friendly web app behavior.
- SQLite migration scaffold: relational schema, adapter, JSON migration command, and round-trip tests.
- SQLite runtime mode: `SPENDSIGHT_STORAGE=json|sqlite|dual` can now run the app from SQLite, with cloud backup/restore exporting through the normalized data API.
- Account-linked transactions: add/edit/review/receipt/OCR flows can store `account_id` while preserving `payment_method` for legacy reports.
- OCR parsing: pasted receipt text now extracts merchant, date, amount, payment method, category suggestion, confidence, and duplicate matches.
- Recurring calendar: `/recurring-calendar` combines fixed bills, EMIs, detected subscriptions, goal contributions, paid/unpaid status, and cash-after-bills.

## P0: Foundation

- Add formal schema migration versioning on top of the SQLite adapter.
- Move default deployment from JSON to SQLite after one more end-to-end data parity pass.

## P1: Market-grade planning

- Goals linked to accounts and transactions.
- Rules preview and optional retroactive application to historical transactions.

## P2: Accounts and net worth

- Add transfer matching between accounts.
- Add account balance history charts and cashflow reports.

## P3: Reporting and UX

- Saved report filters and richer CSV/PDF exports.
- Split transaction support.
- Tags and hidden-from-budget transactions.
- Mobile camera receipt capture and server-side OCR fallback.
- Move large inline CSS/JS from `base.html` into static assets.

## Source references reviewed

- YNAB features: bank import, targets, loan planner, spending/net-worth reports.
- Monarch Money: category/flex budgeting, transaction rules, recurring detection, goals.
- PocketGuard: recurring bills and leftover planning.
- Rocket Money: subscription tracking, budget monitoring, price/bill-change alerts.
