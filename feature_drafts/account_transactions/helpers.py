"""Pure helpers for an account-linked transaction UX draft.

The current SpendSight transaction shape stores ``payment_method`` directly on
expenses. These helpers add optional account linking without requiring parent
schema edits:

* existing payment methods can resolve to account ids through aliases,
* ``payment_method`` remains the backward-compatible display field,
* likely account transfers can be reviewed before hiding them from spending,
* balances can be recomputed from account anchors plus linked transactions.

All functions are Flask-free, file-free, and return copies or derived models.
"""

from __future__ import annotations

from copy import deepcopy
from datetime import date, datetime
import re
from typing import Any


ACCOUNT_REFERENCE_FIELDS = ("account_id", "account", "source_account_id", "from_account_id")
SOURCE_ACCOUNT_FIELDS = ("from_account_id", "source_account_id")
DESTINATION_ACCOUNT_FIELDS = (
    "to_account_id",
    "destination_account_id",
    "target_account_id",
    "transfer_account_id",
)

TRANSFER_KEYWORDS = (
    "transfer",
    "self transfer",
    "between accounts",
    "credit card payment",
    "card payment",
    "loan payment",
    "wallet topup",
    "wallet top-up",
    "wallet top up",
    "topup",
    "top-up",
    "upi self",
    "neft",
    "imps",
    "rtgs",
)

INFLOW_DIRECTIONS = {"in", "inflow", "credit", "deposit", "income", "received", "refund"}
OUTFLOW_DIRECTIONS = {"out", "outflow", "debit", "withdrawal", "expense", "paid", "payment"}
LIABILITY_TYPES = {"credit_card", "loan", "liability", "other_liability", "debt"}
PAYMENT_METHOD_DRAFT_SOURCE = "payment_method_draft"

_WHITESPACE_RE = re.compile(r"\s+")
_NON_ID_RE = re.compile(r"[^A-Za-z0-9_-]+")


def build_account_link_context(
    data_or_accounts: dict[str, Any] | list[dict[str, Any]] | None,
    *,
    include_payment_method_drafts: bool = True,
) -> dict[str, Any]:
    """Return account records plus lookup maps for account-linked transactions.

    ``include_payment_method_drafts`` creates derived account options for
    payment methods that are not already account aliases. This keeps the draft
    usable against today's flat ``payment_methods`` list without mutating data.
    """

    raw_accounts, payment_methods = _raw_accounts_and_payment_methods(data_or_accounts)
    accounts: list[dict[str, Any]] = []
    seen_ids: set[str] = set()
    warnings: list[str] = []

    for raw in raw_accounts:
        if not isinstance(raw, dict):
            continue
        account = _normalize_account_record(raw, seen_ids)
        accounts.append(account)
        seen_ids.add(account["id"].casefold())

    if include_payment_method_drafts:
        existing_aliases = {
            _match_key(alias)
            for account in accounts
            for alias in account.get("payment_methods", [])
            if _match_key(alias)
        }
        existing_aliases.update(_match_key(account["id"]) for account in accounts)
        existing_aliases.update(_match_key(account["name"]) for account in accounts)

        for payment_method in payment_methods:
            label = _clean_string(payment_method)
            if not label:
                continue
            if _match_key(label) in existing_aliases:
                continue
            account_id = _unique_id(_slug(label), seen_ids)
            account_type = _infer_account_type_from_payment_method(label)
            account = {
                "id": account_id,
                "name": label,
                "type": account_type,
                "institution": "",
                "currency_code": "",
                "is_liability": account_type in LIABILITY_TYPES,
                "include_in_net_worth": True,
                "is_archived": False,
                "payment_methods": [label],
                "display_payment_method": label,
                "source": PAYMENT_METHOD_DRAFT_SOURCE,
                "raw": {},
            }
            accounts.append(account)
            seen_ids.add(account_id.casefold())
            existing_aliases.add(_match_key(label))

    by_id = {account["id"]: account for account in accounts}
    by_id_casefold = {account["id"].casefold(): account["id"] for account in accounts}
    alias_to_account_id: dict[str, str] = {}
    display_by_account_id: dict[str, str] = {}

    for account in accounts:
        display = _account_display_name(account)
        display_by_account_id[account["id"]] = display
        aliases = _account_aliases(account)
        for alias in aliases:
            key = _match_key(alias)
            if not key:
                continue
            owner = alias_to_account_id.get(key)
            if owner and owner != account["id"]:
                warnings.append(
                    f"Alias {alias!r} maps to both {owner!r} and {account['id']!r}; keeping {owner!r}."
                )
                continue
            alias_to_account_id[key] = account["id"]

    payment_method_options = [_clean_string(item) for item in payment_methods if _clean_string(item)]
    return {
        "accounts": accounts,
        "by_id": by_id,
        "by_id_casefold": by_id_casefold,
        "alias_to_account_id": alias_to_account_id,
        "display_by_account_id": display_by_account_id,
        "payment_methods": payment_method_options,
        "warnings": warnings,
    }


def resolve_account_id_for_payment_method(
    payment_method: Any,
    data_or_accounts: dict[str, Any] | list[dict[str, Any]] | None,
    *,
    default: str = "",
    include_payment_method_drafts: bool = True,
) -> str:
    """Resolve today's ``payment_method`` value to an account id."""

    context = build_account_link_context(
        data_or_accounts,
        include_payment_method_drafts=include_payment_method_drafts,
    )
    return _resolve_account_reference(payment_method, context, default=default)


def resolve_transaction_account_id(
    transaction: dict[str, Any],
    data_or_accounts: dict[str, Any] | list[dict[str, Any]] | None,
    *,
    default: str = "",
    include_payment_method_drafts: bool = True,
) -> str:
    """Resolve an account id for a transaction-like dict.

    Explicit account fields win. If none are present, the current
    ``payment_method`` value is resolved through account aliases or draft
    payment-method accounts.
    """

    if not isinstance(transaction, dict):
        return default
    context = build_account_link_context(
        data_or_accounts,
        include_payment_method_drafts=include_payment_method_drafts,
    )
    return _resolve_transaction_account_id_from_context(transaction, context, default=default)


def payment_method_display(
    transaction: dict[str, Any],
    data_or_accounts: dict[str, Any] | list[dict[str, Any]] | None = None,
    *,
    default: str = "Cash",
    include_payment_method_drafts: bool = True,
) -> str:
    """Return the old-template-safe payment method label for a transaction."""

    if not isinstance(transaction, dict):
        return default
    existing = _clean_string(transaction.get("payment_method"))
    if existing:
        return existing
    if data_or_accounts is None:
        return default

    context = build_account_link_context(
        data_or_accounts,
        include_payment_method_drafts=include_payment_method_drafts,
    )
    account_id = _resolve_transaction_account_id_from_context(transaction, context)
    if account_id:
        return context["display_by_account_id"].get(account_id, default)
    return default


def apply_account_link_to_transaction(
    transaction: dict[str, Any],
    data_or_accounts: dict[str, Any] | list[dict[str, Any]] | None,
    *,
    account_id: Any = None,
    preserve_payment_method: bool = True,
    default_payment_method: str = "Cash",
    include_payment_method_drafts: bool = True,
) -> dict[str, Any]:
    """Return a transaction copy with ``account_id`` and display-safe payment method.

    This is the parent route bridge: add an account id selected by the new UI,
    but keep ``payment_method`` populated for all existing templates and reports.
    """

    if not isinstance(transaction, dict):
        raise TypeError("transaction must be a dict.")

    context = build_account_link_context(
        data_or_accounts,
        include_payment_method_drafts=include_payment_method_drafts,
    )
    linked = deepcopy(transaction)
    resolved = _resolve_account_reference(account_id, context) if account_id is not None else ""
    if not resolved:
        resolved = _resolve_transaction_account_id_from_context(linked, context)
    if resolved:
        linked["account_id"] = resolved

    existing_display = _clean_string(linked.get("payment_method"))
    if not preserve_payment_method or not existing_display:
        linked["payment_method"] = context["display_by_account_id"].get(
            resolved,
            existing_display or default_payment_method,
        )
    return linked


def backfill_transaction_account_links(
    transactions: list[dict[str, Any]],
    data_or_accounts: dict[str, Any] | list[dict[str, Any]] | None,
    *,
    include_payment_method_drafts: bool = True,
) -> dict[str, Any]:
    """Return transaction copies with account ids resolved where possible."""

    context = build_account_link_context(
        data_or_accounts,
        include_payment_method_drafts=include_payment_method_drafts,
    )
    linked: list[dict[str, Any]] = []
    unresolved: list[dict[str, Any]] = []
    for index, transaction in enumerate(transactions or [], start=1):
        if not isinstance(transaction, dict):
            unresolved.append({"index": index, "reason": "transaction is not an object"})
            continue
        item = apply_account_link_to_transaction(
            transaction,
            data_or_accounts,
            include_payment_method_drafts=include_payment_method_drafts,
        )
        linked.append(item)
        if not item.get("account_id"):
            unresolved.append(
                {
                    "index": index,
                    "id": _clean_string(item.get("id") or item.get("transaction_id")),
                    "payment_method": _clean_string(item.get("payment_method")),
                    "reason": "no account alias matched the payment method",
                }
            )
    return {"transactions": linked, "unresolved": unresolved, "warnings": context["warnings"]}


def build_account_transaction_view_model(
    data: dict[str, Any],
    transaction: dict[str, Any] | None = None,
    *,
    include_payment_method_drafts: bool = True,
) -> dict[str, Any]:
    """Build a small route/template model for account-linked transaction forms."""

    context = build_account_link_context(
        data,
        include_payment_method_drafts=include_payment_method_drafts,
    )
    selected_account_id = ""
    display = "Cash"
    if transaction:
        selected_account_id = _resolve_transaction_account_id_from_context(transaction, context)
        display = payment_method_display(transaction, context["accounts"])

    return {
        "account_options": [
            {
                "id": account["id"],
                "label": account["name"],
                "payment_method": context["display_by_account_id"].get(account["id"], account["name"]),
                "type": account["type"],
                "is_liability": account["is_liability"],
                "is_archived": account["is_archived"],
                "is_draft": account.get("source") == PAYMENT_METHOD_DRAFT_SOURCE,
                "selected": account["id"] == selected_account_id,
            }
            for account in context["accounts"]
            if not account.get("is_archived")
        ],
        "payment_methods": context["payment_methods"],
        "selected_account_id": selected_account_id,
        "payment_method_display": display,
        "warnings": context["warnings"],
    }


def transaction_signed_cashflow(transaction: dict[str, Any]) -> float:
    """Return transaction cashflow sign before liability balance conversion.

    Positive means money moved into the account. Negative means money moved out.
    Existing SpendSight expenses have positive ``amount`` values with no
    direction, so they default to outflows.
    """

    if not isinstance(transaction, dict):
        raise TypeError("transaction must be a dict.")
    if "signed_amount" in transaction and _clean_string(transaction.get("signed_amount")):
        return _money(transaction.get("signed_amount"), "signed_amount")
    if "cashflow" in transaction and _clean_string(transaction.get("cashflow")):
        return _money(transaction.get("cashflow"), "cashflow")

    amount = _money(transaction.get("amount", 0), "amount")
    direction_text = _clean_string(
        transaction.get("direction")
        or transaction.get("flow")
        or transaction.get("transaction_direction")
        or transaction.get("kind")
    ).casefold()
    type_text = _clean_string(transaction.get("type") or transaction.get("transaction_type")).casefold()
    source_text = _clean_string(transaction.get("_collection_key") or transaction.get("source")).casefold()

    if direction_text in INFLOW_DIRECTIONS or type_text in INFLOW_DIRECTIONS:
        return abs(amount)
    if direction_text in OUTFLOW_DIRECTIONS or type_text in OUTFLOW_DIRECTIONS:
        return -abs(amount)
    if source_text in {"extra_income", "income", "income_transactions"}:
        return abs(amount)
    if amount < 0:
        return amount
    return -abs(amount)


def transaction_account_delta(transaction: dict[str, Any], account: dict[str, Any]) -> float:
    """Return the balance delta for ``account`` from ``transaction``."""

    cashflow = transaction_signed_cashflow(transaction)
    if _is_liability_account(account):
        return round(-cashflow, 2)
    return round(cashflow, 2)


def build_account_movements(
    transactions: list[dict[str, Any]],
    data_or_accounts: dict[str, Any] | list[dict[str, Any]] | None,
    *,
    include_payment_method_drafts: bool = True,
) -> dict[str, Any]:
    """Expand transaction-like rows into account balance movements."""

    context = build_account_link_context(
        data_or_accounts,
        include_payment_method_drafts=include_payment_method_drafts,
    )
    movements: list[dict[str, Any]] = []
    unlinked: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []

    for index, transaction in enumerate(transactions or [], start=1):
        if not isinstance(transaction, dict):
            skipped.append({"index": index, "reason": "transaction is not an object"})
            continue
        try:
            tx_date = _transaction_date(transaction)
            amount = abs(_money(transaction.get("amount", 0), "amount"))
            from_account_id, to_account_id = _explicit_transfer_accounts(transaction, context)
            if from_account_id and to_account_id:
                if from_account_id == to_account_id:
                    skipped.append({"index": index, "reason": "transfer accounts are the same"})
                    continue
                for account_id, cashflow, role in (
                    (from_account_id, -amount, "transfer_source"),
                    (to_account_id, amount, "transfer_destination"),
                ):
                    account = context["by_id"].get(account_id)
                    if not account:
                        unlinked.append(_unlinked_preview(index, transaction, f"unknown account_id {account_id!r}"))
                        continue
                    movements.append(
                        _movement_record(
                            index,
                            transaction,
                            account,
                            tx_date,
                            cashflow,
                            role=role,
                            paired_account_id=to_account_id if role == "transfer_source" else from_account_id,
                        )
                    )
                continue

            account_id = _resolve_transaction_account_id_from_context(transaction, context)
            account = context["by_id"].get(account_id)
            if not account:
                unlinked.append(_unlinked_preview(index, transaction, "no linked account"))
                continue
            movements.append(
                _movement_record(
                    index,
                    transaction,
                    account,
                    tx_date,
                    transaction_signed_cashflow(transaction),
                    role="transaction",
                    paired_account_id="",
                )
            )
        except ValueError as exc:
            skipped.append({"index": index, "reason": str(exc)})

    movements.sort(key=lambda item: (item["date"], item["account_id"], item["id"]))
    return {
        "movements": movements,
        "unlinked": unlinked,
        "skipped": skipped,
        "warnings": context["warnings"],
    }


def detect_account_transfers(
    transactions_or_data: dict[str, Any] | list[dict[str, Any]],
    data_or_accounts: dict[str, Any] | list[dict[str, Any]] | None = None,
    *,
    window_days: int = 2,
    amount_tolerance: float = 1.0,
    include_single_sided: bool = True,
    include_payment_method_drafts: bool = True,
) -> dict[str, Any]:
    """Detect likely transfers between accounts without mutating transactions."""

    transactions, account_source = _transactions_and_account_source(transactions_or_data, data_or_accounts)
    context = build_account_link_context(
        account_source,
        include_payment_method_drafts=include_payment_method_drafts,
    )
    compacted: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []

    for index, transaction in enumerate(transactions, start=1):
        if not isinstance(transaction, dict):
            skipped.append({"index": index, "reason": "transaction is not an object"})
            continue
        try:
            compacted.append(_compact_transaction(index, transaction, context))
        except ValueError as exc:
            skipped.append({"index": index, "reason": str(exc)})

    candidates: list[dict[str, Any]] = []
    used_keys: set[str] = set()

    for item in compacted:
        explicit_from, explicit_to = item["explicit_from_account_id"], item["explicit_to_account_id"]
        if not explicit_from or not explicit_to or explicit_from == explicit_to:
            continue
        candidates.append(
            {
                "kind": "explicit_transfer",
                "confidence": "high",
                "amount": item["amount"],
                "from_account_id": explicit_from,
                "to_account_id": explicit_to,
                "transaction": _public_transaction(item),
                "keyword_hits": item["keyword_hits"],
                "assumption": "Transaction already contains source and destination account fields.",
            }
        )
        used_keys.add(item["key"])

    possible_pairs: list[tuple[int, float, int, dict[str, Any], dict[str, Any], list[str]]] = []
    outflows = [item for item in compacted if item["signed_cashflow"] < 0 and item["key"] not in used_keys]
    inflows = [item for item in compacted if item["signed_cashflow"] > 0 and item["key"] not in used_keys]
    tolerance = max(0.0, float(amount_tolerance))
    days = max(0, int(window_days))

    for outflow in outflows:
        for inflow in inflows:
            if not outflow["account_id"] or not inflow["account_id"]:
                continue
            if outflow["account_id"] and inflow["account_id"] and outflow["account_id"] == inflow["account_id"]:
                continue
            amount_delta = round(abs(outflow["amount"] - inflow["amount"]), 2)
            if amount_delta > tolerance:
                continue
            day_delta = abs((outflow["date_obj"] - inflow["date_obj"]).days)
            if day_delta > days:
                continue
            keyword_hits = sorted(set(outflow["keyword_hits"] + inflow["keyword_hits"]))
            rank = _confidence_rank(amount_delta, day_delta, bool(keyword_hits))
            possible_pairs.append((rank, amount_delta, day_delta, outflow, inflow, keyword_hits))

    possible_pairs.sort(key=lambda item: (item[0], item[1], item[2], -item[3]["amount"]))
    for _rank, amount_delta, day_delta, outflow, inflow, keyword_hits in possible_pairs:
        if outflow["key"] in used_keys or inflow["key"] in used_keys:
            continue
        confidence = _confidence(amount_delta, day_delta, bool(keyword_hits))
        candidates.append(
            {
                "kind": "paired_transfer",
                "confidence": confidence,
                "amount": max(outflow["amount"], inflow["amount"]),
                "amount_delta": amount_delta,
                "day_delta": day_delta,
                "from_account_id": outflow["account_id"],
                "to_account_id": inflow["account_id"],
                "outflow": _public_transaction(outflow),
                "inflow": _public_transaction(inflow),
                "keyword_hits": keyword_hits,
                "assumption": "Matched by opposite signed amounts across different accounts within the configured date window.",
            }
        )
        used_keys.add(outflow["key"])
        used_keys.add(inflow["key"])

    if include_single_sided:
        for item in compacted:
            if item["key"] in used_keys or not item["keyword_hits"]:
                continue
            candidates.append(
                {
                    "kind": "single_sided_transfer_review",
                    "confidence": "low",
                    "amount": item["amount"],
                    "account_id": item["account_id"],
                    "transaction": _public_transaction(item),
                    "keyword_hits": item["keyword_hits"],
                    "assumption": "Keyword suggests a transfer, but only one side is visible in the supplied transactions.",
                }
            )

    candidates.sort(key=lambda item: (_confidence_rank_name(item["confidence"]), item.get("day_delta", 99), -item["amount"]))
    return {
        "assumptions": transfer_detection_assumptions(),
        "window_days": days,
        "amount_tolerance": round(tolerance, 2),
        "candidates": candidates,
        "skipped": skipped,
        "warnings": context["warnings"],
    }


def transfer_detection_assumptions() -> list[str]:
    return [
        "Existing SpendSight expense rows are treated as account outflows unless direction or signed_amount says otherwise.",
        "Paired transfers require opposite directions, similar amounts, different accounts, and nearby dates.",
        "Keyword-only matches are review candidates, not automatic spending exclusions.",
        "Single explicit transfer rows can carry from_account_id/source_account_id plus to_account_id/transfer_account_id.",
    ]


def recompute_account_balances(
    data: dict[str, Any],
    *,
    transactions: list[dict[str, Any]] | None = None,
    as_of: Any = None,
    include_archived: bool = False,
    include_payment_method_drafts: bool = True,
) -> dict[str, Any]:
    """Recompute account balances from balance anchors plus transactions.

    Anchor precedence is latest dated anchor on or before ``as_of``. Undated
    ``balance``, ``current_balance``, or ``manual_balance`` values are treated
    as current-at-``as_of`` manual anchors, so old transactions are not double
    counted. Opening balances include same-day transactions; manual/current
    anchors include only later transactions.
    """

    if not isinstance(data, dict):
        raise TypeError("data must be a dict.")
    as_of_date = _parse_date(as_of) or date.today()
    context = build_account_link_context(
        data,
        include_payment_method_drafts=include_payment_method_drafts,
    )
    transaction_rows = _transactions_from_data(data) if transactions is None else list(transactions or [])
    movement_result = build_account_movements(
        transaction_rows,
        context["accounts"],
        include_payment_method_drafts=False,
    )
    movements_by_account: dict[str, list[dict[str, Any]]] = {}
    for movement in movement_result["movements"]:
        if movement["date_obj"] <= as_of_date:
            movements_by_account.setdefault(movement["account_id"], []).append(movement)

    snapshots = _balance_snapshots(data, context)
    rows: list[dict[str, Any]] = []
    totals = {"assets": 0.0, "liabilities": 0.0, "net_worth": 0.0, "excluded": 0.0}

    for account in context["accounts"]:
        if account.get("is_archived") and not include_archived:
            continue
        anchor = _select_balance_anchor(account, snapshots.get(account["id"], []), as_of_date)
        account_movements = [
            movement
            for movement in movements_by_account.get(account["id"], [])
            if _movement_is_after_anchor(movement, anchor)
        ]
        transaction_delta = round(sum(movement["delta"] for movement in account_movements), 2)
        balance = round(anchor["balance"] + transaction_delta, 2)
        signed_balance = -abs(balance) if account["is_liability"] else balance
        included = bool(account.get("include_in_net_worth", True)) and not account.get("is_archived")

        if included:
            if account["is_liability"] or balance < 0:
                totals["liabilities"] = round(totals["liabilities"] + abs(balance), 2)
            else:
                totals["assets"] = round(totals["assets"] + balance, 2)
        else:
            totals["excluded"] = round(totals["excluded"] + balance, 2)

        rows.append(
            {
                "account_id": account["id"],
                "name": account["name"],
                "payment_method": context["display_by_account_id"].get(account["id"], account["name"]),
                "type": account["type"],
                "is_liability": account["is_liability"],
                "include_in_net_worth": account.get("include_in_net_worth", True),
                "is_archived": account.get("is_archived", False),
                "balance": balance,
                "signed_balance": round(signed_balance, 2),
                "anchor_balance": anchor["balance"],
                "anchor_source": anchor["source"],
                "anchor_date": anchor["date"],
                "transaction_delta": transaction_delta,
                "transaction_count": len(account_movements),
                "movement_ids": [movement["id"] for movement in account_movements],
            }
        )

    totals["net_worth"] = round(totals["assets"] - totals["liabilities"], 2)
    rows.sort(key=lambda item: (item["is_archived"], item["type"], item["name"].casefold()))
    return {
        "as_of": as_of_date.isoformat(),
        "accounts": rows,
        "totals": totals,
        "unlinked_transactions": movement_result["unlinked"],
        "skipped_transactions": movement_result["skipped"],
        "warnings": context["warnings"] + movement_result["warnings"],
    }


def _raw_accounts_and_payment_methods(data_or_accounts: Any) -> tuple[list[dict[str, Any]], list[Any]]:
    if isinstance(data_or_accounts, dict):
        accounts = data_or_accounts.get("accounts", [])
        payment_methods = data_or_accounts.get("payment_methods", [])
    elif isinstance(data_or_accounts, list):
        accounts = data_or_accounts
        payment_methods = []
    else:
        accounts = []
        payment_methods = []
    if not isinstance(accounts, list):
        accounts = []
    if isinstance(payment_methods, str):
        payment_methods = [payment_methods]
    if not isinstance(payment_methods, list):
        payment_methods = []
    return accounts, payment_methods


def _normalize_account_record(raw: dict[str, Any], seen_ids: set[str]) -> dict[str, Any]:
    name = _clean_string(raw.get("name") or raw.get("label") or raw.get("payment_method") or "Account")
    account_id = _clean_string(raw.get("id") or raw.get("account_id") or _slug(name))
    account_id = _unique_id(account_id, seen_ids)
    account_type = _clean_string(raw.get("type") or raw.get("account_type") or _infer_account_type_from_payment_method(name))
    account_type = account_type.casefold().replace(" ", "_").replace("-", "_") or "bank"
    aliases = _list_strings(raw.get("payment_methods"))
    aliases.extend(_list_strings(raw.get("aliases")))
    for alias in (raw.get("payment_method"), raw.get("display_payment_method"), name, account_id):
        cleaned = _clean_string(alias)
        if cleaned and cleaned not in aliases:
            aliases.append(cleaned)

    return {
        "id": account_id,
        "name": name,
        "type": account_type,
        "institution": _clean_string(raw.get("institution")),
        "currency_code": _clean_string(raw.get("currency_code")),
        "is_liability": _boolish(raw.get("is_liability"), account_type in LIABILITY_TYPES),
        "include_in_net_worth": _boolish(raw.get("include_in_net_worth"), True),
        "is_archived": _boolish(raw.get("is_archived", raw.get("archived")), False),
        "payment_methods": aliases,
        "display_payment_method": _clean_string(raw.get("display_payment_method")) or aliases[0],
        "source": _clean_string(raw.get("source")) or "account",
        "raw": deepcopy(raw),
    }


def _account_aliases(account: dict[str, Any]) -> list[str]:
    aliases = _list_strings(account.get("payment_methods"))
    for alias in (account.get("display_payment_method"), account.get("name"), account.get("id")):
        cleaned = _clean_string(alias)
        if cleaned and cleaned not in aliases:
            aliases.append(cleaned)
    return aliases


def _account_display_name(account: dict[str, Any]) -> str:
    for value in (
        account.get("display_payment_method"),
        (account.get("payment_methods") or [""])[0],
        account.get("name"),
        account.get("id"),
    ):
        cleaned = _clean_string(value)
        if cleaned:
            return cleaned
    return "Account"


def _resolve_transaction_account_id_from_context(
    transaction: dict[str, Any],
    context: dict[str, Any],
    *,
    default: str = "",
) -> str:
    for field in ACCOUNT_REFERENCE_FIELDS:
        resolved = _resolve_account_reference(transaction.get(field), context)
        if resolved:
            return resolved
    return _resolve_account_reference(transaction.get("payment_method"), context, default=default)


def _resolve_fields(transaction: dict[str, Any], fields: tuple[str, ...], context: dict[str, Any]) -> str:
    for field in fields:
        resolved = _resolve_account_reference(transaction.get(field), context)
        if resolved:
            return resolved
    return ""


def _resolve_account_reference(value: Any, context: dict[str, Any], *, default: str = "") -> str:
    cleaned = _clean_string(value)
    if not cleaned:
        return default
    by_id_casefold = context.get("by_id_casefold", {})
    if cleaned.casefold() in by_id_casefold:
        return by_id_casefold[cleaned.casefold()]
    return context.get("alias_to_account_id", {}).get(_match_key(cleaned), default)


def _explicit_transfer_accounts(transaction: dict[str, Any], context: dict[str, Any]) -> tuple[str, str]:
    primary = _resolve_transaction_account_id_from_context(transaction, context)
    explicit_from = _resolve_fields(transaction, SOURCE_ACCOUNT_FIELDS, context)
    explicit_to = _resolve_fields(transaction, DESTINATION_ACCOUNT_FIELDS, context)

    if explicit_from and explicit_to:
        return explicit_from, explicit_to
    if explicit_from and primary and primary != explicit_from:
        return explicit_from, primary
    if explicit_to and primary and primary != explicit_to:
        return primary, explicit_to
    return "", ""


def _movement_record(
    index: int,
    transaction: dict[str, Any],
    account: dict[str, Any],
    tx_date: date,
    signed_cashflow: float,
    *,
    role: str,
    paired_account_id: str,
) -> dict[str, Any]:
    delta = round(-signed_cashflow if account["is_liability"] else signed_cashflow, 2)
    return {
        "index": index,
        "id": _clean_string(transaction.get("id") or transaction.get("transaction_id") or f"row-{index}"),
        "date": tx_date.isoformat(),
        "date_obj": tx_date,
        "account_id": account["id"],
        "paired_account_id": paired_account_id,
        "amount": round(abs(signed_cashflow), 2),
        "signed_cashflow": round(signed_cashflow, 2),
        "delta": delta,
        "role": role,
        "payment_method": _clean_string(transaction.get("payment_method")) or _account_display_name(account),
        "description": _transaction_description(transaction),
    }


def _unlinked_preview(index: int, transaction: dict[str, Any], reason: str) -> dict[str, Any]:
    return {
        "index": index,
        "id": _clean_string(transaction.get("id") or transaction.get("transaction_id")),
        "date": _clean_string(transaction.get("date") or transaction.get("posted_date")),
        "amount": _coerce_money(transaction.get("amount")),
        "payment_method": _clean_string(transaction.get("payment_method")),
        "reason": reason,
    }


def _transactions_and_account_source(
    transactions_or_data: dict[str, Any] | list[dict[str, Any]],
    data_or_accounts: dict[str, Any] | list[dict[str, Any]] | None,
) -> tuple[list[dict[str, Any]], dict[str, Any] | list[dict[str, Any]] | None]:
    if isinstance(transactions_or_data, dict):
        return _transactions_from_data(transactions_or_data), data_or_accounts or transactions_or_data
    return list(transactions_or_data or []), data_or_accounts


def _transactions_from_data(data: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for key, default_type in (
        ("transactions", ""),
        ("account_transactions", ""),
        ("expenses", "expense"),
        ("extra_income", "income"),
    ):
        collection = data.get(key, [])
        if not isinstance(collection, list):
            continue
        for item in collection:
            if not isinstance(item, dict):
                continue
            row = dict(item)
            row.setdefault("_collection_key", key)
            if default_type:
                row.setdefault("transaction_type", default_type)
            rows.append(row)
    return rows


def _compact_transaction(index: int, transaction: dict[str, Any], context: dict[str, Any]) -> dict[str, Any]:
    tx_date = _transaction_date(transaction)
    signed_cashflow = transaction_signed_cashflow(transaction)
    account_id = _resolve_transaction_account_id_from_context(transaction, context)
    explicit_from, explicit_to = _explicit_transfer_accounts(transaction, context)
    return {
        "index": index,
        "key": _clean_string(transaction.get("id") or transaction.get("transaction_id") or f"row-{index}"),
        "id": _clean_string(transaction.get("id") or transaction.get("transaction_id")),
        "date": tx_date.isoformat(),
        "date_obj": tx_date,
        "account_id": account_id,
        "explicit_from_account_id": explicit_from,
        "explicit_to_account_id": explicit_to,
        "amount": round(abs(signed_cashflow), 2),
        "signed_cashflow": round(signed_cashflow, 2),
        "keyword_hits": _keyword_hits(transaction),
        "description": _transaction_description(transaction),
    }


def _public_transaction(item: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": item["id"],
        "date": item["date"],
        "account_id": item["account_id"],
        "amount": item["amount"],
        "signed_cashflow": item["signed_cashflow"],
        "description": item["description"],
    }


def _confidence(amount_delta: float, day_delta: int, has_keyword: bool) -> str:
    if amount_delta == 0 and day_delta <= 1:
        return "high" if has_keyword else "medium"
    if amount_delta <= 1 and day_delta <= 2:
        return "medium"
    return "low"


def _confidence_rank(amount_delta: float, day_delta: int, has_keyword: bool) -> int:
    return _confidence_rank_name(_confidence(amount_delta, day_delta, has_keyword))


def _confidence_rank_name(confidence: str) -> int:
    return {"high": 0, "medium": 1, "low": 2}.get(confidence, 3)


def _balance_snapshots(data: dict[str, Any], context: dict[str, Any]) -> dict[str, list[dict[str, Any]]]:
    snapshots_by_account: dict[str, list[dict[str, Any]]] = {}
    for key in ("account_balance_snapshots", "manual_account_balances", "account_manual_balances", "account_balances"):
        collection = data.get(key, [])
        if not isinstance(collection, list):
            continue
        for index, snapshot in enumerate(collection, start=1):
            if not isinstance(snapshot, dict):
                continue
            account_id = _resolve_account_reference(snapshot.get("account_id") or snapshot.get("account"), context)
            if not account_id:
                continue
            snapshot_date = _parse_date(snapshot.get("date") or snapshot.get("balance_date") or snapshot.get("as_of"))
            if not snapshot_date:
                continue
            try:
                balance = _money(snapshot.get("balance"), "balance")
            except ValueError:
                continue
            snapshots_by_account.setdefault(account_id, []).append(
                {
                    "balance": balance,
                    "date_obj": snapshot_date,
                    "date": snapshot_date.isoformat(),
                    "source": _clean_string(snapshot.get("source")) or f"{key}[{index}]",
                    "priority": 50,
                    "include_anchor_date_transactions": False,
                }
            )
    return snapshots_by_account


def _select_balance_anchor(
    account: dict[str, Any],
    snapshots: list[dict[str, Any]],
    as_of_date: date,
) -> dict[str, Any]:
    raw = account.get("raw", {}) if isinstance(account.get("raw"), dict) else {}
    anchors: list[dict[str, Any]] = []

    opening_date = _parse_date(raw.get("opening_date") or raw.get("created_on")) or date.min
    if "opening_balance" in raw:
        anchors.append(
            _anchor(
                _money(raw.get("opening_balance"), "opening_balance"),
                opening_date,
                "account.opening_balance",
                priority=10,
                include_anchor_date_transactions=True,
            )
        )

    for field, source, date_fields, priority in (
        ("current_balance", "account.current_balance", ("current_balance_date", "balance_date", "updated_at"), 40),
        ("manual_balance", "account.manual_balance", ("manual_balance_date", "balance_date", "updated_at"), 45),
        ("balance", "account.balance", ("balance_date", "balance_as_of", "updated_at"), 35),
    ):
        if field not in raw:
            continue
        anchor_date = _first_date(raw, date_fields) or as_of_date
        anchors.append(
            _anchor(
                _money(raw.get(field), field),
                anchor_date,
                source,
                priority=priority,
                include_anchor_date_transactions=False,
            )
        )

    anchors.extend(snapshots)
    usable = [anchor for anchor in anchors if anchor["date_obj"] <= as_of_date]
    if not usable:
        return _anchor(0, date.min, "default.zero", priority=0, include_anchor_date_transactions=True)
    selected = sorted(usable, key=lambda item: (item["date_obj"], item["priority"], item["source"]))[-1]
    return selected


def _anchor(
    balance: float,
    anchor_date: date,
    source: str,
    *,
    priority: int,
    include_anchor_date_transactions: bool,
) -> dict[str, Any]:
    return {
        "balance": round(balance, 2),
        "date_obj": anchor_date,
        "date": "" if anchor_date == date.min else anchor_date.isoformat(),
        "source": source,
        "priority": priority,
        "include_anchor_date_transactions": include_anchor_date_transactions,
    }


def _movement_is_after_anchor(movement: dict[str, Any], anchor: dict[str, Any]) -> bool:
    if anchor["date_obj"] == date.min:
        return True
    if anchor["include_anchor_date_transactions"]:
        return movement["date_obj"] >= anchor["date_obj"]
    return movement["date_obj"] > anchor["date_obj"]


def _first_date(raw: dict[str, Any], fields: tuple[str, ...]) -> date | None:
    for field in fields:
        parsed = _parse_date(raw.get(field))
        if parsed:
            return parsed
    return None


def _transaction_date(transaction: dict[str, Any]) -> date:
    parsed = _parse_date(transaction.get("date") or transaction.get("posted_date") or transaction.get("transaction_date"))
    if not parsed:
        raise ValueError("transaction date must be an ISO-like date.")
    return parsed


def _transaction_description(transaction: dict[str, Any]) -> str:
    fields = ("notes", "memo", "description", "merchant", "payee", "category", "subcategory", "type", "transaction_type")
    return _clean_string(" ".join(_clean_string(transaction.get(field)) for field in fields if _clean_string(transaction.get(field))))


def _keyword_hits(transaction: dict[str, Any]) -> list[str]:
    text = _match_key(_transaction_description(transaction))
    return [keyword for keyword in TRANSFER_KEYWORDS if _match_key(keyword) in text]


def _infer_account_type_from_payment_method(payment_method: Any) -> str:
    text = _match_key(payment_method)
    if text == "cash":
        return "cash"
    if "credit" in text or "card" in text:
        return "credit_card"
    if "wallet" in text or "paytm" in text or "phonepe" in text or "gpay" in text:
        return "wallet"
    return "bank"


def _is_liability_account(account: dict[str, Any]) -> bool:
    return _boolish(account.get("is_liability"), _clean_string(account.get("type")).casefold() in LIABILITY_TYPES)


def _list_strings(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        value = [value]
    if not isinstance(value, list):
        return []
    result: list[str] = []
    for item in value:
        cleaned = _clean_string(item)
        if cleaned and cleaned not in result:
            result.append(cleaned)
    return result


def _money(value: Any, field_name: str) -> float:
    try:
        return round(float(str(value).replace(",", "").strip()), 2)
    except (TypeError, ValueError):
        raise ValueError(f"{field_name} must be numeric.")


def _coerce_money(value: Any) -> float | None:
    try:
        return _money(value, "amount")
    except ValueError:
        return None


def _parse_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raw = _clean_string(value)
    if not raw:
        return None
    if "T" in raw:
        try:
            return datetime.fromisoformat(raw.replace("Z", "+00:00")).date()
        except ValueError:
            pass
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%d %b %Y", "%d %B %Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    try:
        return date.fromisoformat(raw)
    except ValueError:
        return None


def _boolish(value: Any, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    text = _clean_string(value).casefold()
    if not text:
        return default
    return text in {"1", "true", "yes", "y", "on"}


def _clean_string(value: Any) -> str:
    return _WHITESPACE_RE.sub(" ", str(value or "").strip())


def _match_key(value: Any) -> str:
    text = _clean_string(value).casefold()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return _WHITESPACE_RE.sub(" ", text).strip()


def _slug(value: Any) -> str:
    slug = _NON_ID_RE.sub("-", _clean_string(value).casefold()).strip("-_")
    return slug or "account"


def _unique_id(raw_id: str, seen_ids: set[str]) -> str:
    base = _slug(raw_id)
    candidate = base
    suffix = 2
    while candidate.casefold() in seen_ids:
        candidate = f"{base}-{suffix}"
        suffix += 1
    return candidate
