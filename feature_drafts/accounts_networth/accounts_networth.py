"""Pure account, balance, net-worth, and transfer helpers for SpendSight.

The helpers use the same JSON-dict style as the existing SpendSight app and add
only optional top-level keys:

    data["accounts"] = [
        {
            "id": "card-1",
            "name": "Card 1",
            "type": "credit_card",
            "currency_code": "INR",
            "opening_balance": 0.0,
            "opening_date": "2026-05-31",
            "is_liability": True,
            "include_in_net_worth": True,
            "payment_methods": ["Card 1"],
            "archived": False,
        }
    ]

    data["account_balance_snapshots"] = [
        {
            "id": "snapshot-id",
            "account_id": "card-1",
            "date": "2026-05-31",
            "balance": 12500.0,
            "source": "manual",
        }
    ]

Liability balances are stored as positive amounts owed. Net worth subtracts
liabilities and also treats negative asset balances as liabilities.
"""

from __future__ import annotations

from copy import deepcopy
from datetime import date, datetime
import re
import uuid


ACCOUNT_TYPES = {
    "cash": {"label": "Cash", "is_liability": False},
    "bank": {"label": "Bank", "is_liability": False},
    "wallet": {"label": "Wallet", "is_liability": False},
    "investment": {"label": "Investment", "is_liability": False},
    "other_asset": {"label": "Other Asset", "is_liability": False},
    "credit_card": {"label": "Credit Card", "is_liability": True},
    "loan": {"label": "Loan", "is_liability": True},
    "other_liability": {"label": "Other Liability", "is_liability": True},
}

ACCOUNT_TYPE_ALIASES = {
    "account": "bank",
    "asset": "other_asset",
    "checking": "bank",
    "savings": "bank",
    "card": "credit_card",
    "cc": "credit_card",
    "credit": "credit_card",
    "liability": "other_liability",
}

SNAPSHOT_SOURCES = {"manual", "import", "migration", "computed"}

TRANSFER_KEYWORDS = (
    "transfer",
    "self transfer",
    "between accounts",
    "upi self",
    "neft",
    "imps",
    "rtgs",
    "card payment",
    "credit card payment",
    "loan payment",
    "wallet topup",
    "wallet top-up",
)

TRANSFER_DETECTION_ASSUMPTIONS = [
    "Existing SpendSight expense rows are treated as outflows unless a signed amount or direction field says otherwise.",
    "High-confidence transfer pairs require opposite directions, similar amounts, different accounts, and nearby dates.",
    "Keyword-only matches are review candidates, not automatic exclusions from spending.",
    "Credit card and loan payments are modeled as transfers when one side reduces a bank/cash account and the other side reduces a liability balance.",
    "Amount tolerance accounts for fees, rounding, and statement currency formatting; it should stay small for auto-matching.",
]

_SAFE_ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{0,63}$")
_WHITESPACE_RE = re.compile(r"\s+")


def _today_iso(ref_today=None):
    if ref_today is None:
        return date.today().isoformat()
    if isinstance(ref_today, datetime):
        return ref_today.date().isoformat()
    if isinstance(ref_today, date):
        return ref_today.isoformat()
    return _parse_iso_date(ref_today, "date").isoformat()


def _parse_iso_date(value, field_name):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return date.fromisoformat(str(value))
    except (TypeError, ValueError):
        raise ValueError(f"{field_name} must be an ISO date in YYYY-MM-DD format.")


def _money(value, field_name="amount", *, allow_negative=True):
    try:
        amount = float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        raise ValueError(f"{field_name} must be numeric.")
    if not allow_negative and amount < 0:
        raise ValueError(f"{field_name} cannot be negative.")
    return round(amount, 2)


def _clean_text(value, default=""):
    cleaned = _WHITESPACE_RE.sub(" ", str(value or "").strip())
    return cleaned or default


def _boolish(value, default=False):
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    return str(value).strip().casefold() in {"1", "true", "yes", "y", "on"}


def _slug(value):
    slug = re.sub(r"[^A-Za-z0-9_-]+", "-", str(value or "").strip().lower())
    slug = slug.strip("-_")
    return slug or f"account-{uuid.uuid4().hex[:8]}"


def _normalize_account_type(raw_type):
    raw = _clean_text(raw_type or "bank").casefold().replace(" ", "_").replace("-", "_")
    raw = ACCOUNT_TYPE_ALIASES.get(raw, raw)
    if raw not in ACCOUNT_TYPES:
        raise ValueError(f"Unsupported account type: {raw_type!r}.")
    return raw


def _default_liability_for_type(account_type):
    return ACCOUNT_TYPES[account_type]["is_liability"]


def _currency_code(data_or_code):
    if isinstance(data_or_code, dict):
        raw = data_or_code.get("currency_code", "INR")
    else:
        raw = data_or_code or "INR"
    return _clean_text(raw, "INR").upper()[:8]


def _copy_data(data):
    if not isinstance(data, dict):
        raise ValueError("SpendSight data must be a JSON object.")
    return deepcopy(data)


def new_account(
    name,
    account_type="bank",
    *,
    account_id=None,
    opening_balance=0,
    opening_date=None,
    currency_code="INR",
    institution="",
    is_liability=None,
    include_in_net_worth=True,
    payment_methods=None,
    notes="",
    archived=False,
    created_at=None,
):
    """Create a JSON-serializable account dict.

    This function does not mutate the app data dict. Use
    ``normalize_accounts_data`` after adding the returned account to data.
    """

    account_type = _normalize_account_type(account_type)
    account_id = _clean_text(account_id or _slug(name))
    if not _SAFE_ID_RE.fullmatch(account_id):
        raise ValueError("account_id must use letters, numbers, underscores, or hyphens.")
    if not _clean_text(name):
        raise ValueError("Account name is required.")
    if is_liability is None:
        is_liability = _default_liability_for_type(account_type)
    opening_date = _today_iso(opening_date)
    aliases = [_clean_text(alias) for alias in (payment_methods or [])]
    aliases = [alias for alias in aliases if alias]
    if not aliases:
        aliases = [_clean_text(name)]

    return {
        "id": account_id,
        "name": _clean_text(name),
        "type": account_type,
        "institution": _clean_text(institution),
        "currency_code": _currency_code(currency_code),
        "opening_balance": _money(opening_balance, "opening_balance"),
        "opening_date": opening_date,
        "is_liability": bool(is_liability),
        "include_in_net_worth": _boolish(include_in_net_worth, True),
        "payment_methods": aliases,
        "notes": _clean_text(notes),
        "archived": _boolish(archived),
        "created_at": created_at or datetime.now().isoformat(timespec="seconds"),
    }


def new_balance_snapshot(
    account_id,
    balance,
    *,
    snapshot_date=None,
    source="manual",
    snapshot_id=None,
    notes="",
):
    """Create a JSON-serializable account balance snapshot."""

    account_id = _clean_text(account_id)
    if not account_id:
        raise ValueError("account_id is required.")
    source = _clean_text(source, "manual").casefold()
    if source not in SNAPSHOT_SOURCES:
        raise ValueError(f"snapshot source must be one of: {', '.join(sorted(SNAPSHOT_SOURCES))}.")
    snapshot_date = _today_iso(snapshot_date)
    snapshot_id = _clean_text(snapshot_id or f"{account_id}-{snapshot_date}-{uuid.uuid4().hex[:8]}")
    if not _SAFE_ID_RE.fullmatch(snapshot_id):
        raise ValueError("snapshot_id must use letters, numbers, underscores, or hyphens.")
    return {
        "id": snapshot_id,
        "account_id": account_id,
        "date": snapshot_date,
        "balance": _money(balance, "balance"),
        "source": source,
        "notes": _clean_text(notes),
    }


def _normalize_account(account, default_currency):
    if not isinstance(account, dict):
        raise ValueError("Account must be an object.")

    account_type = _normalize_account_type(account.get("type", "bank"))
    account_id = _clean_text(account.get("id") or _slug(account.get("name") or account_type))
    if not _SAFE_ID_RE.fullmatch(account_id):
        raise ValueError(f"Invalid account id: {account_id!r}.")

    name = _clean_text(account.get("name"))
    if not name:
        raise ValueError(f"Account {account_id!r} needs a name.")

    opening_date = account.get("opening_date") or account.get("created_on") or _today_iso()
    opening_date = _parse_iso_date(opening_date, f"accounts.{account_id}.opening_date").isoformat()
    raw_aliases = account.get("payment_methods")
    if raw_aliases is None:
        raw_aliases = account.get("aliases", [])
    if isinstance(raw_aliases, str):
        raw_aliases = [raw_aliases]
    if not isinstance(raw_aliases, list):
        raw_aliases = []
    aliases = []
    for alias in raw_aliases:
        cleaned = _clean_text(alias)
        if cleaned and cleaned not in aliases:
            aliases.append(cleaned)
    if name not in aliases:
        aliases.insert(0, name)

    normalized = {
        "id": account_id,
        "name": name,
        "type": account_type,
        "institution": _clean_text(account.get("institution")),
        "currency_code": _currency_code(account.get("currency_code") or default_currency),
        "opening_balance": _money(account.get("opening_balance", 0), f"accounts.{account_id}.opening_balance"),
        "opening_date": opening_date,
        "is_liability": _boolish(account.get("is_liability"), _default_liability_for_type(account_type)),
        "include_in_net_worth": _boolish(account.get("include_in_net_worth"), True),
        "payment_methods": aliases,
        "notes": _clean_text(account.get("notes")),
        "archived": _boolish(account.get("archived")),
    }
    if account.get("created_at"):
        normalized["created_at"] = _clean_text(account.get("created_at"))
    if "current_balance" in account:
        normalized["current_balance"] = _money(account.get("current_balance"), f"accounts.{account_id}.current_balance")
    return normalized


def _normalize_balance_snapshot(snapshot):
    if not isinstance(snapshot, dict):
        raise ValueError("Balance snapshot must be an object.")
    account_id = _clean_text(snapshot.get("account_id"))
    if not account_id:
        raise ValueError("Balance snapshot needs account_id.")
    snapshot_date = _parse_iso_date(snapshot.get("date"), f"snapshots.{account_id}.date").isoformat()
    source = _clean_text(snapshot.get("source", "manual"), "manual").casefold()
    if source not in SNAPSHOT_SOURCES:
        source = "manual"
    snapshot_id = _clean_text(snapshot.get("id") or f"{account_id}-{snapshot_date}")
    if not _SAFE_ID_RE.fullmatch(snapshot_id):
        raise ValueError(f"Invalid balance snapshot id: {snapshot_id!r}.")
    return {
        "id": snapshot_id,
        "account_id": account_id,
        "date": snapshot_date,
        "balance": _money(snapshot.get("balance", 0), f"snapshots.{account_id}.balance"),
        "source": source,
        "notes": _clean_text(snapshot.get("notes")),
    }


def ensure_accounts_schema(data):
    """Return a copy of data with account-related top-level keys present."""

    normalized = _copy_data(data)
    normalized.setdefault("accounts", [])
    normalized.setdefault("account_balance_snapshots", [])
    normalized.setdefault("net_worth_snapshots", [])
    if not isinstance(normalized["accounts"], list):
        normalized["accounts"] = []
    if not isinstance(normalized["account_balance_snapshots"], list):
        normalized["account_balance_snapshots"] = []
    if not isinstance(normalized["net_worth_snapshots"], list):
        normalized["net_worth_snapshots"] = []
    return normalized


def normalize_accounts_data(data):
    """Return a normalized copy of SpendSight data including account fields.

    Invalid shapes raise ValueError so route handlers can flash the message and
    skip saving.
    """

    normalized = ensure_accounts_schema(data)
    currency_code = _currency_code(normalized)

    accounts = []
    seen_account_ids = set()
    for account in normalized.get("accounts", []):
        item = _normalize_account(account, currency_code)
        if item["id"] in seen_account_ids:
            raise ValueError(f"Duplicate account id: {item['id']}.")
        seen_account_ids.add(item["id"])
        accounts.append(item)

    snapshots = []
    seen_snapshot_ids = set()
    for snapshot in normalized.get("account_balance_snapshots", []):
        item = _normalize_balance_snapshot(snapshot)
        if item["id"] in seen_snapshot_ids:
            raise ValueError(f"Duplicate balance snapshot id: {item['id']}.")
        if item["account_id"] not in seen_account_ids:
            raise ValueError(f"Balance snapshot references unknown account: {item['account_id']}.")
        seen_snapshot_ids.add(item["id"])
        snapshots.append(item)

    normalized["accounts"] = sorted(accounts, key=lambda item: (item["archived"], item["name"].casefold()))
    normalized["account_balance_snapshots"] = sorted(snapshots, key=lambda item: (item["date"], item["account_id"]))
    normalized["net_worth_snapshots"] = [
        item for item in normalized.get("net_worth_snapshots", []) if isinstance(item, dict)
    ]
    return normalized


def validate_accounts_data(data):
    """Validate account-related data and return JSON-friendly diagnostics."""

    errors = []
    warnings = []
    normalized = None
    try:
        normalized = normalize_accounts_data(data)
    except ValueError as exc:
        errors.append(str(exc))
        normalized = ensure_accounts_schema(data) if isinstance(data, dict) else None
        return {"valid": False, "errors": errors, "warnings": warnings, "data": normalized}

    account_ids = {account["id"] for account in normalized["accounts"]}
    alias_owner = {}
    for account in normalized["accounts"]:
        for alias in account.get("payment_methods", []):
            key = alias.casefold()
            if key in alias_owner and alias_owner[key] != account["id"]:
                warnings.append(
                    f"Payment method alias {alias!r} is used by both {alias_owner[key]!r} and {account['id']!r}."
                )
            alias_owner[key] = account["id"]

    for idx, expense in enumerate(normalized.get("expenses", []), start=1):
        if not isinstance(expense, dict):
            continue
        account_id = _clean_text(expense.get("account_id"))
        payment_method = _clean_text(expense.get("payment_method"))
        if account_id and account_id not in account_ids:
            warnings.append(f"Expense #{idx} references unknown account_id {account_id!r}.")
        elif payment_method and payment_method.casefold() not in alias_owner:
            warnings.append(f"Expense #{idx} payment_method {payment_method!r} is not linked to an account.")

    return {"valid": not errors, "errors": errors, "warnings": warnings, "data": normalized}


def _infer_account_type_from_payment_method(payment_method):
    text = _clean_text(payment_method).casefold()
    if text == "cash":
        return "cash"
    if "credit" in text or "card" in text:
        return "credit_card"
    if "wallet" in text or "paytm" in text or "phonepe" in text or "gpay" in text:
        return "wallet"
    return "bank"


def migrate_payment_methods_to_accounts(data, *, opening_date=None, opening_balance=0):
    """Return a copy of data with one account draft per payment method.

    This is a bridge helper for the current SpendSight model. It preserves the
    existing ``payment_methods`` list and adds missing account records with that
    payment method as an alias.
    """

    migrated = ensure_accounts_schema(data)
    existing_aliases = set()
    existing_ids = {str(account.get("id", "")).casefold() for account in migrated.get("accounts", []) if isinstance(account, dict)}
    for account in migrated.get("accounts", []):
        if not isinstance(account, dict):
            continue
        for alias in account.get("payment_methods") or []:
            existing_aliases.add(str(alias).casefold())

    for payment_method in migrated.get("payment_methods", []):
        payment_method = _clean_text(payment_method)
        if not payment_method or payment_method.casefold() in existing_aliases:
            continue
        account_id = _slug(payment_method)
        suffix = 2
        while account_id.casefold() in existing_ids:
            account_id = f"{_slug(payment_method)}-{suffix}"
            suffix += 1
        account = new_account(
            payment_method,
            _infer_account_type_from_payment_method(payment_method),
            account_id=account_id,
            opening_balance=opening_balance,
            opening_date=opening_date,
            currency_code=_currency_code(migrated),
            payment_methods=[payment_method],
            notes="Draft account generated from existing payment_methods.",
        )
        migrated["accounts"].append(account)
        existing_aliases.add(payment_method.casefold())
        existing_ids.add(account_id.casefold())

    return normalize_accounts_data(migrated)


def _snapshot_date(snapshot):
    return _parse_iso_date(snapshot["date"], "snapshot.date")


def _latest_snapshot_for_account(snapshots, account_id, as_of_date):
    candidates = [
        snapshot
        for snapshot in snapshots
        if snapshot.get("account_id") == account_id and _snapshot_date(snapshot) <= as_of_date
    ]
    if not candidates:
        return None
    return sorted(candidates, key=lambda item: (item["date"], item.get("id", "")))[-1]


def build_account_balances(data, *, as_of=None, include_archived=False):
    """Build account balance rows as of a date."""

    normalized = normalize_accounts_data(data)
    as_of_date = _parse_iso_date(as_of or _today_iso(), "as_of")
    rows = []
    for account in normalized["accounts"]:
        if account.get("archived") and not include_archived:
            continue
        latest = _latest_snapshot_for_account(
            normalized.get("account_balance_snapshots", []),
            account["id"],
            as_of_date,
        )
        if latest:
            balance = latest["balance"]
            source = latest["source"]
            source_date = latest["date"]
        elif "current_balance" in account:
            balance = account["current_balance"]
            source = "account.current_balance"
            source_date = account["opening_date"]
        else:
            balance = account["opening_balance"]
            source = "account.opening_balance"
            source_date = account["opening_date"]

        is_liability = bool(account["is_liability"])
        include = bool(account["include_in_net_worth"])
        if is_liability:
            signed_balance = -abs(balance)
        else:
            signed_balance = balance

        rows.append({
            "account_id": account["id"],
            "name": account["name"],
            "type": account["type"],
            "type_label": ACCOUNT_TYPES[account["type"]]["label"],
            "institution": account.get("institution", ""),
            "currency_code": account["currency_code"],
            "balance": round(balance, 2),
            "signed_balance": round(signed_balance, 2),
            "is_liability": is_liability,
            "include_in_net_worth": include,
            "archived": bool(account.get("archived")),
            "source": source,
            "source_date": source_date,
        })
    return rows


def build_net_worth_snapshot(data, *, as_of=None, snapshot_id=None, source="computed"):
    """Return a computed net-worth snapshot for the given data dict."""

    as_of = _today_iso(as_of)
    balances = build_account_balances(data, as_of=as_of)
    included = [row for row in balances if row["include_in_net_worth"]]

    assets_total = 0.0
    liabilities_total = 0.0
    by_type = {}
    for row in included:
        balance = row["balance"]
        if row["is_liability"]:
            liabilities_total += abs(balance)
        elif balance < 0:
            liabilities_total += abs(balance)
        else:
            assets_total += balance

        bucket = by_type.setdefault(row["type"], {"type_label": row["type_label"], "assets": 0.0, "liabilities": 0.0})
        if row["is_liability"] or balance < 0:
            bucket["liabilities"] = round(bucket["liabilities"] + abs(balance), 2)
        else:
            bucket["assets"] = round(bucket["assets"] + balance, 2)

    assets_total = round(assets_total, 2)
    liabilities_total = round(liabilities_total, 2)
    net_worth = round(assets_total - liabilities_total, 2)

    return {
        "id": snapshot_id or f"net-worth-{as_of}",
        "date": as_of,
        "assets_total": assets_total,
        "liabilities_total": liabilities_total,
        "net_worth": net_worth,
        "by_type": by_type,
        "accounts": balances,
        "source": source,
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }


def build_net_worth_series(data, *, start_date=None, end_date=None):
    """Build a net-worth time series from account snapshot dates."""

    normalized = normalize_accounts_data(data)
    end = _parse_iso_date(end_date or _today_iso(), "end_date")
    start = _parse_iso_date(start_date, "start_date") if start_date else None
    dates = set()
    for account in normalized["accounts"]:
        try:
            dates.add(_parse_iso_date(account.get("opening_date"), "account.opening_date"))
        except ValueError:
            pass
    for snapshot in normalized["account_balance_snapshots"]:
        snapshot_dt = _parse_iso_date(snapshot["date"], "snapshot.date")
        if snapshot_dt <= end:
            dates.add(snapshot_dt)
    dates.add(end)

    series = []
    for item_date in sorted(dates):
        if start and item_date < start:
            continue
        if item_date > end:
            continue
        snapshot = build_net_worth_snapshot(normalized, as_of=item_date.isoformat())
        series.append({
            "date": snapshot["date"],
            "assets_total": snapshot["assets_total"],
            "liabilities_total": snapshot["liabilities_total"],
            "net_worth": snapshot["net_worth"],
        })
    return series


def append_balance_snapshot(data, account_id, balance, *, snapshot_date=None, source="manual", notes=""):
    """Return a data copy with a new account balance snapshot appended."""

    updated = normalize_accounts_data(data)
    account_ids = {account["id"] for account in updated["accounts"]}
    if account_id not in account_ids:
        raise ValueError(f"Unknown account_id: {account_id}.")
    updated["account_balance_snapshots"].append(
        new_balance_snapshot(account_id, balance, snapshot_date=snapshot_date, source=source, notes=notes)
    )
    return normalize_accounts_data(updated)


def append_net_worth_snapshot(data, *, as_of=None):
    """Return a data copy with a computed net-worth snapshot stored."""

    updated = normalize_accounts_data(data)
    snapshot = build_net_worth_snapshot(updated, as_of=as_of)
    updated["net_worth_snapshots"] = [
        item for item in updated.get("net_worth_snapshots", []) if item.get("date") != snapshot["date"]
    ]
    updated["net_worth_snapshots"].append(snapshot)
    updated["net_worth_snapshots"].sort(key=lambda item: item.get("date", ""))
    return updated


def transfer_detection_assumptions():
    """Return the assumptions used by transfer detection."""

    return list(TRANSFER_DETECTION_ASSUMPTIONS)


def _account_alias_map(data_or_accounts):
    if not data_or_accounts:
        return {}
    if isinstance(data_or_accounts, dict):
        accounts = data_or_accounts.get("accounts", [])
    else:
        accounts = data_or_accounts
    aliases = {}
    for account in accounts:
        if not isinstance(account, dict):
            continue
        account_id = _clean_text(account.get("id"))
        if not account_id:
            continue
        names = [account.get("name"), account_id]
        names.extend(account.get("payment_methods") or [])
        names.extend(account.get("aliases") or [])
        for name in names:
            cleaned = _clean_text(name)
            if cleaned:
                aliases[cleaned.casefold()] = account_id
    return aliases


def _transaction_account_id(transaction, alias_map):
    for field in ("account_id", "account", "source_account_id"):
        value = _clean_text(transaction.get(field))
        if value:
            return alias_map.get(value.casefold(), value)
    payment_method = _clean_text(transaction.get("payment_method"))
    if payment_method:
        return alias_map.get(payment_method.casefold(), payment_method)
    return ""


def _transaction_description(transaction):
    fields = (
        "notes",
        "memo",
        "description",
        "merchant",
        "payee",
        "category",
        "subcategory",
        "type",
        "transaction_type",
    )
    return " ".join(_clean_text(transaction.get(field)) for field in fields if _clean_text(transaction.get(field))).casefold()


def _keyword_hits(transaction):
    text = _transaction_description(transaction)
    return [keyword for keyword in TRANSFER_KEYWORDS if keyword in text]


def _transaction_date(transaction):
    return _parse_iso_date(transaction.get("date") or transaction.get("posted_date"), "transaction.date")


def _signed_transaction_amount(transaction):
    if "signed_amount" in transaction:
        return _money(transaction.get("signed_amount"), "signed_amount")
    amount = _money(transaction.get("amount", 0), "amount")
    direction = _clean_text(
        transaction.get("direction") or transaction.get("flow") or transaction.get("transaction_direction")
    ).casefold()
    transaction_type = _clean_text(transaction.get("type") or transaction.get("transaction_type")).casefold()
    if direction in {"in", "inflow", "credit", "deposit"} or transaction_type in {"income", "credit", "deposit"}:
        return abs(amount)
    if direction in {"out", "outflow", "debit", "withdrawal", "expense"} or transaction_type in {"expense", "debit", "withdrawal"}:
        return -abs(amount)
    if amount < 0:
        return amount
    return -abs(amount)


def _compact_transaction(transaction, alias_map):
    signed_amount = _signed_transaction_amount(transaction)
    tx_date = _transaction_date(transaction)
    return {
        "id": _clean_text(transaction.get("id") or transaction.get("transaction_id") or transaction.get("uuid")),
        "date": tx_date.isoformat(),
        "date_obj": tx_date,
        "account_id": _transaction_account_id(transaction, alias_map),
        "amount": round(abs(signed_amount), 2),
        "signed_amount": round(signed_amount, 2),
        "keyword_hits": _keyword_hits(transaction),
        "description": _clean_text(
            transaction.get("notes")
            or transaction.get("memo")
            or transaction.get("description")
            or transaction.get("merchant")
            or transaction.get("subcategory")
        ),
        "raw": transaction,
    }


def _confidence(amount_delta, day_delta, has_keyword):
    if amount_delta == 0 and day_delta <= 1:
        return "high" if has_keyword else "medium"
    if amount_delta <= 1 and day_delta <= 2:
        return "medium"
    return "low"


def detect_transfer_candidates(
    transactions,
    data_or_accounts=None,
    *,
    window_days=2,
    amount_tolerance=1.0,
    include_single_sided=True,
):
    """Detect likely transfer candidates from transaction-like dicts.

    Returns a JSON-friendly dict with paired candidates and optional single
    sided review candidates. It does not mutate transactions.
    """

    if not isinstance(transactions, list):
        raise ValueError("transactions must be a list.")
    alias_map = _account_alias_map(data_or_accounts)
    compacted = []
    skipped = []
    for idx, transaction in enumerate(transactions, start=1):
        if not isinstance(transaction, dict):
            skipped.append({"index": idx, "reason": "transaction is not an object"})
            continue
        try:
            compacted.append(_compact_transaction(transaction, alias_map))
        except ValueError as exc:
            skipped.append({"index": idx, "reason": str(exc)})

    outflows = [item for item in compacted if item["signed_amount"] < 0]
    inflows = [item for item in compacted if item["signed_amount"] > 0]
    candidates = []
    used_ids = set()

    for outflow in outflows:
        for inflow in inflows:
            if outflow["id"] and inflow["id"] and outflow["id"] == inflow["id"]:
                continue
            if outflow["account_id"] and inflow["account_id"] and outflow["account_id"] == inflow["account_id"]:
                continue
            amount_delta = round(abs(outflow["amount"] - inflow["amount"]), 2)
            if amount_delta > amount_tolerance:
                continue
            day_delta = abs((outflow["date_obj"] - inflow["date_obj"]).days)
            if day_delta > window_days:
                continue
            keyword_hits = sorted(set(outflow["keyword_hits"] + inflow["keyword_hits"]))
            pair_key = tuple(sorted([outflow["id"] or f"out-{id(outflow)}", inflow["id"] or f"in-{id(inflow)}"]))
            candidates.append({
                "kind": "paired_transfer",
                "confidence": _confidence(amount_delta, day_delta, bool(keyword_hits)),
                "amount": max(outflow["amount"], inflow["amount"]),
                "amount_delta": amount_delta,
                "day_delta": day_delta,
                "from_account_id": outflow["account_id"],
                "to_account_id": inflow["account_id"],
                "outflow": _public_transaction(outflow),
                "inflow": _public_transaction(inflow),
                "keyword_hits": keyword_hits,
                "assumption": "Matched by opposite signed amounts across different accounts within the configured date window.",
            })
            used_ids.add(outflow["id"] or f"out-{id(outflow)}")
            used_ids.add(inflow["id"] or f"in-{id(inflow)}")

    if include_single_sided:
        for item in compacted:
            item_key = item["id"] or f"single-{id(item)}"
            if item_key in used_ids or not item["keyword_hits"]:
                continue
            candidates.append({
                "kind": "single_sided_transfer_review",
                "confidence": "low",
                "amount": item["amount"],
                "account_id": item["account_id"],
                "transaction": _public_transaction(item),
                "keyword_hits": item["keyword_hits"],
                "assumption": "Keyword suggests a transfer, but only one side is visible in the supplied transactions.",
            })

    candidates.sort(key=lambda item: (item.get("confidence") != "high", item.get("day_delta", 99), -item["amount"]))
    return {
        "assumptions": transfer_detection_assumptions(),
        "window_days": int(window_days),
        "amount_tolerance": round(float(amount_tolerance), 2),
        "candidates": candidates,
        "skipped": skipped,
    }


def _public_transaction(item):
    return {
        "id": item["id"],
        "date": item["date"],
        "account_id": item["account_id"],
        "amount": item["amount"],
        "signed_amount": item["signed_amount"],
        "description": item["description"],
    }
