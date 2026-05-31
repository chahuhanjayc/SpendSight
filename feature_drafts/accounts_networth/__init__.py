"""Draft account and net-worth helpers for SpendSight.

This package is intentionally isolated under feature_drafts so the parent app
can review and wire it into Flask without touching the current route/template
surface.
"""

from .accounts_networth import (
    ACCOUNT_TYPES,
    TRANSFER_DETECTION_ASSUMPTIONS,
    append_balance_snapshot,
    append_net_worth_snapshot,
    build_account_balances,
    build_net_worth_series,
    build_net_worth_snapshot,
    detect_transfer_candidates,
    ensure_accounts_schema,
    migrate_payment_methods_to_accounts,
    new_account,
    new_balance_snapshot,
    normalize_accounts_data,
    transfer_detection_assumptions,
    validate_accounts_data,
)

__all__ = [
    "ACCOUNT_TYPES",
    "TRANSFER_DETECTION_ASSUMPTIONS",
    "append_balance_snapshot",
    "append_net_worth_snapshot",
    "build_account_balances",
    "build_net_worth_series",
    "build_net_worth_snapshot",
    "detect_transfer_candidates",
    "ensure_accounts_schema",
    "migrate_payment_methods_to_accounts",
    "new_account",
    "new_balance_snapshot",
    "normalize_accounts_data",
    "transfer_detection_assumptions",
    "validate_accounts_data",
]
