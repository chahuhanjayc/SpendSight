"""Reusable envelope/rollover budget helpers for SpendSight.

These helpers are intentionally pure: they accept SpendSight's existing JSON
data dict from ``load_data()`` and return derived view models or config payloads.
They do not import Flask, read files, write files, or mutate the input dict.

Existing keys consumed:
- expenses: dated spending rows with category and amount
- budget_limits: existing monthly category budgets
- income: monthly_salary and optional salary_history
- extra_income: one-time or recurring extra income rows
- fixed_expenses: EMI/fixed expense rows from the Income page
- billing_start_day: billing cycle start day

Optional draft key:
data["envelope_budget"] = {
    "start_month": "2026-01",
    "default_rollover": false,
    "rollover_unassigned": true,
    "unassigned_balance": 0,
    "carryover": {"Groceries": 500},
    "monthly_assignments": {"2026-05": {"Groceries": 12000}},
    "category_settings": {
        "Groceries": {"monthly_budget": 12000, "rollover": true},
        "Insurance": {"annual_amount": 24000, "due_month": 12}
    }
}
"""

from __future__ import annotations

from calendar import monthrange
from collections import defaultdict
from copy import deepcopy
from datetime import date, datetime, timedelta
from typing import Any, Dict, Iterable, List, Mapping, MutableMapping, Optional, Tuple


ENVELOPE_KEY = "envelope_budget"
DEFAULT_FIXED_CATEGORY = "EMI / Finance"

_FREQUENCY_MONTHS = {
    "monthly": 1,
    "month": 1,
    "quarterly": 3,
    "quarter": 3,
    "half_yearly": 6,
    "half-yearly": 6,
    "semiannual": 6,
    "semi_annual": 6,
    "yearly": 12,
    "annual": 12,
    "annually": 12,
}


def build_envelope_budget(data: Mapping[str, Any], ref_today: Optional[date] = None) -> Dict[str, Any]:
    """Build the envelope budget view model for the current billing period.

    The result includes category rows, rollover balances, sinking-fund set-asides,
    spending progress, income, and the current left-to-assign number.
    """

    ref = _as_date(ref_today) or date.today()
    billing_start_day = _billing_start_day(data.get("billing_start_day", 1))
    period_start, period_end = get_billing_period(billing_start_day, ref)
    period_key = _month_key(period_start)

    config = normalize_envelope_config(data, ref_today=ref)
    config_start = _period_start_from_month_key(config["start_month"], billing_start_day)
    if config_start > period_start:
        config_start = period_start

    category_settings = config["category_settings"]
    budget_limits = _clean_money_map(data.get("budget_limits", {}))
    fixed_sinking = _fixed_sinking_suggestions(data.get("fixed_expenses", []))
    current_spend = _spend_by_category(data, period_start, period_end, include_fixed=True)
    current_fixed = _fixed_totals_by_category(current_spend["fixed_occurrences"])

    categories = _budget_categories(
        budget_limits=budget_limits,
        settings=category_settings,
        spend_totals=current_spend["totals"],
        fixed_sinking=fixed_sinking,
    )
    categories = [c for c in categories if c not in config["excluded_categories"]]

    balance_start_by_category: Dict[str, float] = {}
    for category in categories:
        setting = category_settings.get(category, {})
        rollover = _rollover_enabled(setting, config, has_sinking_fund=category in fixed_sinking)
        balance = _opening_balance(category, setting, config)
        if rollover:
            for start, end in _iter_billing_periods(config_start, _add_months(period_start, -1)):
                period_spend = _spend_by_category(data, start, end, include_fixed=True)
                fixed_due = _fixed_totals_by_category(period_spend["fixed_occurrences"])
                assignment = _assignment_for_category(
                    category=category,
                    period_start=start,
                    period_end=end,
                    config=config,
                    setting=setting,
                    budget_limits=budget_limits,
                    fixed_sinking=fixed_sinking,
                    fixed_due_by_category=fixed_due,
                )["assigned"]
                balance += assignment - period_spend["totals"].get(category, 0.0)
        else:
            balance = 0.0
        balance_start_by_category[category] = round(balance, 2)

    rows = []
    for category in categories:
        setting = category_settings.get(category, {})
        rollover = _rollover_enabled(setting, config, has_sinking_fund=category in fixed_sinking)
        assignment = _assignment_for_category(
            category=category,
            period_start=period_start,
            period_end=period_end,
            config=config,
            setting=setting,
            budget_limits=budget_limits,
            fixed_sinking=fixed_sinking,
            fixed_due_by_category=current_fixed,
        )
        assigned = assignment["assigned"]
        balance_start = balance_start_by_category.get(category, 0.0) if rollover else 0.0
        spent = current_spend["totals"].get(category, 0.0)
        available_before_spend = balance_start + assigned
        remaining = available_before_spend - spent
        pct_spent = (spent / available_before_spend * 100) if available_before_spend > 0 else 0.0
        status = _progress_status(assigned, spent, remaining, pct_spent)

        rows.append(
            {
                "category": category,
                "group": str(setting.get("group") or assignment.get("group") or "Monthly"),
                "rollover": rollover,
                "assigned": round(assigned, 2),
                "base_monthly_budget": round(assignment["base_monthly_budget"], 2),
                "sinking_fund_amount": round(assignment["sinking_fund_amount"], 2),
                "scheduled_assignment": round(assignment["scheduled_assignment"], 2),
                "assignment_source": assignment["source"],
                "balance_start": round(balance_start, 2),
                "available_before_spend": round(available_before_spend, 2),
                "spent": round(spent, 2),
                "remaining": round(remaining, 2),
                "pct_spent": round(pct_spent, 1),
                "status": status,
                "status_label": _status_label(status),
                "transaction_count": current_spend["counts"].get(category, 0),
                "fixed_expense_count": current_spend["fixed_counts"].get(category, 0),
                "sinking_fund": assignment["sinking_fund"],
                "notes": str(setting.get("notes") or ""),
            }
        )

    rows.sort(key=_row_sort_key)

    income = _income_for_period(data, period_start, period_end, ref)
    unassigned_start = _unassigned_balance_before(
        data=data,
        config=config,
        config_start=config_start,
        period_start=period_start,
        budget_limits=budget_limits,
        fixed_sinking=fixed_sinking,
        categories=categories,
    )
    assigned_total = round(sum(row["assigned"] for row in rows), 2)
    spent_total = round(sum(row["spent"] for row in rows), 2)
    remaining_total = round(sum(row["remaining"] for row in rows), 2)
    available_to_assign = round(income["total"] + unassigned_start, 2)
    left_to_assign = round(available_to_assign - assigned_total, 2)

    return {
        "period": {
            "key": period_key,
            "start": period_start.isoformat(),
            "end": period_end.isoformat(),
            "label": f"{period_start.strftime('%b %d')} - {period_end.strftime('%b %d')}",
            "days_total": (period_end - period_start).days + 1,
            "days_elapsed": min(max((ref - period_start).days + 1, 0), (period_end - period_start).days + 1),
        },
        "income": {
            "salary": round(income["salary"], 2),
            "extra_income": round(income["extra_income"], 2),
            "total": round(income["total"], 2),
            "unassigned_start": round(unassigned_start, 2),
            "available_to_assign": available_to_assign,
        },
        "summary": {
            "assigned": assigned_total,
            "spent": spent_total,
            "remaining": remaining_total,
            "left_to_assign": left_to_assign,
            "over_count": sum(1 for row in rows if row["status"] == "over"),
            "warning_count": sum(1 for row in rows if row["status"] == "warning"),
            "no_budget_count": sum(1 for row in rows if row["status"] == "no-budget"),
            "rollover_balance_start": round(sum(row["balance_start"] for row in rows), 2),
            "category_count": len(rows),
        },
        "rows": rows,
        "sinking_funds": [row for row in rows if row["sinking_fund"]],
        "config": {
            "start_month": _month_key(config_start),
            "period_key": period_key,
            "default_rollover": config["default_rollover"],
            "rollover_unassigned": config["rollover_unassigned"],
        },
    }


def normalize_envelope_config(data: Mapping[str, Any], ref_today: Optional[date] = None) -> Dict[str, Any]:
    """Return a sanitized envelope config from ``data[ENVELOPE_KEY]``."""

    ref = _as_date(ref_today) or date.today()
    raw = data.get(ENVELOPE_KEY, {})
    if not isinstance(raw, Mapping):
        raw = {}

    category_settings = _normalize_category_settings(
        raw.get("category_settings", raw.get("categories", raw.get("envelopes", {})))
    )
    monthly_assignments = _normalize_monthly_assignments(
        raw.get("monthly_assignments", raw.get("assignments", {}))
    )
    carryover = _clean_money_map(
        raw.get("carryover", raw.get("starting_balances", raw.get("opening_balances", {})))
    )

    start_month = _coerce_month_key(raw.get("start_month") or raw.get("rollover_start_month"))
    if not start_month:
        billing_start_day = _billing_start_day(data.get("billing_start_day", 1))
        start_month = _month_key(get_billing_period(billing_start_day, ref)[0])

    return {
        "start_month": start_month,
        "default_rollover": _to_bool(raw.get("default_rollover", False)),
        "rollover_unassigned": _to_bool(raw.get("rollover_unassigned", bool(raw.get("unassigned_balance")))),
        "unassigned_balance": _money(raw.get("unassigned_balance", 0)),
        "auto_assign_fixed_expenses": _to_bool(raw.get("auto_assign_fixed_expenses", True)),
        "auto_sinking_funds": _to_bool(raw.get("auto_sinking_funds", True)),
        "carryover": carryover,
        "monthly_assignments": monthly_assignments,
        "category_settings": category_settings,
        "excluded_categories": {
            _clean_category(c)
            for c in raw.get("excluded_categories", [])
            if _clean_category(c)
        },
    }


def build_monthly_assignment_update(
    data: Mapping[str, Any],
    assignments: Mapping[str, Any],
    period_key: Optional[str] = None,
    ref_today: Optional[date] = None,
) -> Dict[str, Any]:
    """Return a new ``data[ENVELOPE_KEY]`` payload with one period assignment merged.

    Use this in a POST route after parsing form/JSON values, then assign the
    returned dict to ``data[ENVELOPE_KEY]`` and call SpendSight's ``save_data``.
    """

    ref = _as_date(ref_today) or date.today()
    billing_start_day = _billing_start_day(data.get("billing_start_day", 1))
    target_period_key = _coerce_month_key(period_key) or _month_key(get_billing_period(billing_start_day, ref)[0])

    existing = data.get(ENVELOPE_KEY, {})
    payload = deepcopy(existing) if isinstance(existing, Mapping) else {}
    period_assignments = {}
    for category, raw_amount in assignments.items():
        clean_category = _clean_category(category)
        if not clean_category:
            continue
        amount = _money(raw_amount)
        if amount > 0:
            period_assignments[clean_category] = round(amount, 2)

    monthly = payload.setdefault("monthly_assignments", {})
    if not isinstance(monthly, MutableMapping):
        monthly = {}
        payload["monthly_assignments"] = monthly
    monthly[target_period_key] = period_assignments
    payload.setdefault("start_month", target_period_key)
    return dict(payload)


def get_billing_period(billing_start_day: int = 1, ref_today: Optional[date] = None) -> Tuple[date, date]:
    """Return the SpendSight billing period containing ``ref_today``."""

    ref = _as_date(ref_today) or date.today()
    day = _billing_start_day(billing_start_day)
    anchor = date(ref.year, ref.month, min(day, monthrange(ref.year, ref.month)[1]))
    if ref.day >= day:
        start = anchor
        end = _add_months(start, 1) - timedelta(days=1)
    else:
        end = anchor - timedelta(days=1)
        start = _add_months(anchor, -1)
    return start, end


def fixed_expense_occurrences(
    fixed_expenses: Iterable[Mapping[str, Any]],
    period_start: date,
    period_end: date,
) -> List[Dict[str, Any]]:
    """Return scheduled fixed/EMI expense occurrences due inside a date range."""

    start = _as_date(period_start)
    end = _as_date(period_end)
    if not start or not end or end < start:
        return []

    occurrences: List[Dict[str, Any]] = []
    first_month = date(start.year, start.month, 1)
    last_month = date(end.year, end.month, 1)

    for item in fixed_expenses or []:
        if not isinstance(item, Mapping):
            continue
        amount = _money(item.get("amount"))
        if amount <= 0:
            continue
        try:
            start_year = int(item.get("start_year", 0))
            start_month = int(item.get("start_month", 0))
            fixed_start = date(start_year, start_month, 1)
        except (TypeError, ValueError):
            continue

        item_type = str(item.get("type") or "fixed").strip().lower()
        frequency = str(item.get("frequency") or "monthly").strip().lower()
        interval = 1 if item_type == "emi" else _FREQUENCY_MONTHS.get(frequency, 1)
        total_months = int(_money(item.get("total_months", 0))) if item_type == "emi" else 0
        day_of_month = _billing_start_day(item.get("day_of_month", 1), max_day=31)
        category = _clean_category(item.get("category")) or DEFAULT_FIXED_CATEGORY

        cursor = max(first_month, fixed_start)
        while cursor <= last_month:
            months_since = _month_diff(fixed_start, cursor)
            if months_since >= 0:
                applies = False
                if item_type == "emi":
                    applies = total_months > 0 and months_since < total_months
                else:
                    applies = months_since % interval == 0
                if applies:
                    due = date(cursor.year, cursor.month, min(day_of_month, monthrange(cursor.year, cursor.month)[1]))
                    if start <= due <= end:
                        occurrences.append(
                            {
                                "id": str(item.get("id") or ""),
                                "name": str(item.get("name") or "Scheduled expense"),
                                "category": category,
                                "amount": round(amount, 2),
                                "date": due.isoformat(),
                                "type": item_type,
                                "frequency": frequency,
                                "interval_months": interval,
                            }
                        )
            cursor = _add_months(cursor, 1)

    return occurrences


def _income_for_period(data: Mapping[str, Any], period_start: date, period_end: date, ref_today: date) -> Dict[str, float]:
    salary = _salary_for_date(data.get("income", {}), ref_today)
    extra = _extra_income_for_period(data.get("extra_income", []), period_start, period_end)
    return {"salary": salary, "extra_income": extra, "total": round(salary + extra, 2)}


def _salary_for_date(income_data: Any, target_date: date) -> float:
    if not isinstance(income_data, Mapping):
        return 0.0
    history = income_data.get("salary_history", [])
    if isinstance(history, list) and history:
        applicable = []
        for row in history:
            if not isinstance(row, Mapping):
                continue
            effective = _as_date(row.get("effective_from"))
            if effective and effective <= target_date:
                applicable.append((effective, _money(row.get("amount"))))
        if applicable:
            applicable.sort(key=lambda item: item[0])
            return round(applicable[-1][1], 2)
    return round(_money(income_data.get("monthly_salary", 0)), 2)


def _extra_income_for_period(extra_income: Any, period_start: date, period_end: date) -> float:
    total = 0.0
    for item in extra_income or []:
        if not isinstance(item, Mapping):
            continue
        amount = _money(item.get("amount"))
        if amount <= 0:
            continue
        income_type = str(item.get("type") or "one-time").strip().lower()
        if income_type == "recurring":
            total += _recurring_amount_in_period(item, amount, period_start, period_end)
        else:
            income_date = _as_date(item.get("date"))
            if income_date and period_start <= income_date <= period_end:
                total += amount
    return round(total, 2)


def _recurring_amount_in_period(item: Mapping[str, Any], amount: float, period_start: date, period_end: date) -> float:
    start_date = _as_date(item.get("start_date") or item.get("date"))
    if not start_date:
        return 0.0
    end_date = _as_date(item.get("end_date")) or period_end
    if end_date < period_start:
        return 0.0

    frequency = str(item.get("frequency") or "monthly").strip().lower()
    interval = _FREQUENCY_MONTHS.get(frequency, 1)
    cursor = date(max(period_start, start_date).year, max(period_start, start_date).month, 1)
    last_month = date(period_end.year, period_end.month, 1)
    total = 0.0
    while cursor <= last_month:
        months_since = _month_diff(date(start_date.year, start_date.month, 1), cursor)
        if months_since >= 0 and months_since % interval == 0:
            occurrence = date(cursor.year, cursor.month, min(start_date.day, monthrange(cursor.year, cursor.month)[1]))
            if start_date <= occurrence <= end_date and period_start <= occurrence <= period_end:
                total += amount
        cursor = _add_months(cursor, 1)
    return total


def _spend_by_category(data: Mapping[str, Any], period_start: date, period_end: date, include_fixed: bool) -> Dict[str, Any]:
    totals: Dict[str, float] = defaultdict(float)
    counts: Dict[str, int] = defaultdict(int)
    for expense in data.get("expenses", []) or []:
        if not isinstance(expense, Mapping):
            continue
        spent_on = _as_date(expense.get("date"))
        if not spent_on or not (period_start <= spent_on <= period_end):
            continue
        amount = _money(expense.get("amount"))
        if amount <= 0:
            continue
        category = _clean_category(expense.get("category")) or "Other"
        totals[category] += amount
        counts[category] += 1

    fixed_occurrences: List[Dict[str, Any]] = []
    fixed_counts: Dict[str, int] = defaultdict(int)
    if include_fixed:
        fixed_occurrences = fixed_expense_occurrences(data.get("fixed_expenses", []), period_start, period_end)
        for occurrence in fixed_occurrences:
            category = occurrence["category"]
            totals[category] += occurrence["amount"]
            counts[category] += 1
            fixed_counts[category] += 1

    return {
        "totals": {category: round(amount, 2) for category, amount in totals.items()},
        "counts": dict(counts),
        "fixed_counts": dict(fixed_counts),
        "fixed_occurrences": fixed_occurrences,
    }


def _fixed_sinking_suggestions(fixed_expenses: Any) -> Dict[str, Dict[str, Any]]:
    suggestions: Dict[str, Dict[str, Any]] = {}
    for item in fixed_expenses or []:
        if not isinstance(item, Mapping):
            continue
        item_type = str(item.get("type") or "fixed").strip().lower()
        if item_type == "emi":
            continue
        frequency = str(item.get("frequency") or "monthly").strip().lower()
        interval = _FREQUENCY_MONTHS.get(frequency, 1)
        if interval <= 1:
            continue
        amount = _money(item.get("amount"))
        if amount <= 0:
            continue
        category = _clean_category(item.get("category")) or DEFAULT_FIXED_CATEGORY
        suggestion = suggestions.setdefault(
            category,
            {
                "amount": 0.0,
                "kind": "scheduled-fixed",
                "items": [],
            },
        )
        monthly_amount = round(amount / interval, 2)
        suggestion["amount"] += monthly_amount
        suggestion["items"].append(
            {
                "name": str(item.get("name") or "Scheduled expense"),
                "amount": round(amount, 2),
                "monthly_amount": monthly_amount,
                "frequency": frequency,
                "interval_months": interval,
            }
        )
    for suggestion in suggestions.values():
        suggestion["amount"] = round(suggestion["amount"], 2)
    return suggestions


def _assignment_for_category(
    *,
    category: str,
    period_start: date,
    period_end: date,
    config: Mapping[str, Any],
    setting: Mapping[str, Any],
    budget_limits: Mapping[str, float],
    fixed_sinking: Mapping[str, Dict[str, Any]],
    fixed_due_by_category: Mapping[str, float],
) -> Dict[str, Any]:
    period_key = _month_key(period_start)
    override = config["monthly_assignments"].get(period_key, {}).get(category)
    explicit = _explicit_monthly_budget(category, setting, budget_limits)
    sinking = _sinking_fund_for_category(category, setting, fixed_sinking, config, period_start)
    scheduled_assignment = 0.0
    source = explicit["source"]

    if override is not None:
        assigned = _money(override)
        source = "monthly-assignment"
    else:
        base = explicit["amount"]
        sinking_amount = sinking["amount"] if sinking else 0.0
        if (
            config["auto_assign_fixed_expenses"]
            and base <= 0
            and sinking_amount <= 0
            and fixed_due_by_category.get(category, 0.0) > 0
        ):
            scheduled_assignment = fixed_due_by_category.get(category, 0.0)
            source = "scheduled-fixed-expense"
        assigned = base + sinking_amount + scheduled_assignment
        if sinking and source == "none":
            source = sinking["kind"]

    return {
        "assigned": round(max(assigned, 0.0), 2),
        "base_monthly_budget": round(explicit["amount"], 2),
        "sinking_fund_amount": round(sinking["amount"] if sinking else 0.0, 2),
        "scheduled_assignment": round(scheduled_assignment, 2),
        "source": source,
        "group": setting.get("group") if isinstance(setting, Mapping) else "",
        "sinking_fund": sinking,
    }


def _explicit_monthly_budget(category: str, setting: Mapping[str, Any], budget_limits: Mapping[str, float]) -> Dict[str, Any]:
    for key in ("monthly_budget", "budget", "limit"):
        if key in setting:
            return {"amount": _money(setting.get(key)), "source": "envelope-setting"}
    if category in budget_limits:
        return {"amount": _money(budget_limits.get(category)), "source": "budget-limits"}
    return {"amount": 0.0, "source": "none"}


def _sinking_fund_for_category(
    category: str,
    setting: Mapping[str, Any],
    fixed_sinking: Mapping[str, Dict[str, Any]],
    config: Mapping[str, Any],
    period_start: date,
) -> Optional[Dict[str, Any]]:
    if not isinstance(setting, Mapping):
        setting = {}

    annual_amount = _money(setting.get("annual_amount"))
    if annual_amount > 0:
        due_month = _int_or_none(setting.get("due_month"))
        due_day = _int_or_none(setting.get("due_day")) or 1
        due_date = ""
        if due_month:
            due_year = period_start.year if due_month >= period_start.month else period_start.year + 1
            due_date = date(due_year, due_month, min(due_day, monthrange(due_year, due_month)[1])).isoformat()
        return {
            "kind": "annual",
            "amount": round(annual_amount / 12, 2),
            "target_amount": round(annual_amount, 2),
            "due_date": due_date,
            "items": [],
        }

    target_amount = _money(setting.get("target_amount", setting.get("irregular_amount", 0)))
    due_date = _as_date(setting.get("due_date"))
    if target_amount > 0 and due_date:
        months = max(_month_diff(date(period_start.year, period_start.month, 1), date(due_date.year, due_date.month, 1)) + 1, 1)
        return {
            "kind": "irregular",
            "amount": round(target_amount / months, 2),
            "target_amount": round(target_amount, 2),
            "due_date": due_date.isoformat(),
            "months_until_due": months,
            "items": [],
        }

    if config["auto_sinking_funds"] and _to_bool(setting.get("auto_sinking", True)):
        suggestion = fixed_sinking.get(category)
        if suggestion:
            return {
                "kind": suggestion["kind"],
                "amount": round(suggestion["amount"], 2),
                "target_amount": round(sum(item["amount"] for item in suggestion["items"]), 2),
                "due_date": "",
                "items": suggestion["items"],
            }
    return None


def _unassigned_balance_before(
    *,
    data: Mapping[str, Any],
    config: Mapping[str, Any],
    config_start: date,
    period_start: date,
    budget_limits: Mapping[str, float],
    fixed_sinking: Mapping[str, Dict[str, Any]],
    categories: Iterable[str],
) -> float:
    balance = _money(config.get("unassigned_balance", 0))
    if not config.get("rollover_unassigned"):
        return 0.0

    for start, end in _iter_billing_periods(config_start, _add_months(period_start, -1)):
        period_income = _income_for_period(data, start, end, start)["total"]
        period_spend = _spend_by_category(data, start, end, include_fixed=True)
        fixed_due = _fixed_totals_by_category(period_spend["fixed_occurrences"])
        assigned = 0.0
        for category in categories:
            setting = config["category_settings"].get(category, {})
            assigned += _assignment_for_category(
                category=category,
                period_start=start,
                period_end=end,
                config=config,
                setting=setting,
                budget_limits=budget_limits,
                fixed_sinking=fixed_sinking,
                fixed_due_by_category=fixed_due,
            )["assigned"]
        balance += period_income - assigned
    return round(balance, 2)


def _budget_categories(
    *,
    budget_limits: Mapping[str, float],
    settings: Mapping[str, Mapping[str, Any]],
    spend_totals: Mapping[str, float],
    fixed_sinking: Mapping[str, Dict[str, Any]],
) -> List[str]:
    categories = {
        _clean_category(category)
        for category in budget_limits.keys()
        if _clean_category(category)
    }
    categories.update(_clean_category(category) for category in settings.keys() if _clean_category(category))
    categories.update(_clean_category(category) for category in spend_totals.keys() if _clean_category(category))
    categories.update(_clean_category(category) for category in fixed_sinking.keys() if _clean_category(category))
    return sorted(categories, key=str.casefold)


def _normalize_category_settings(raw_settings: Any) -> Dict[str, Dict[str, Any]]:
    settings: Dict[str, Dict[str, Any]] = {}
    iterable: Iterable[Tuple[Any, Any]]
    if isinstance(raw_settings, Mapping):
        iterable = raw_settings.items()
    elif isinstance(raw_settings, list):
        iterable = ((item.get("category"), item) for item in raw_settings if isinstance(item, Mapping))
    else:
        iterable = []

    for raw_category, raw_value in iterable:
        category = _clean_category(raw_category)
        if not category:
            continue
        if isinstance(raw_value, Mapping):
            setting = dict(raw_value)
        else:
            setting = {"monthly_budget": raw_value}
        setting.pop("category", None)
        if "monthly_budget" in setting:
            setting["monthly_budget"] = _money(setting["monthly_budget"])
        if "budget" in setting:
            setting["budget"] = _money(setting["budget"])
        if "limit" in setting:
            setting["limit"] = _money(setting["limit"])
        if "annual_amount" in setting:
            setting["annual_amount"] = _money(setting["annual_amount"])
        if "target_amount" in setting:
            setting["target_amount"] = _money(setting["target_amount"])
        if "irregular_amount" in setting:
            setting["irregular_amount"] = _money(setting["irregular_amount"])
        if "carryover" in setting:
            setting["carryover"] = _money(setting["carryover"])
        if "starting_balance" in setting:
            setting["starting_balance"] = _money(setting["starting_balance"])
        settings[category] = setting
    return settings


def _normalize_monthly_assignments(raw_assignments: Any) -> Dict[str, Dict[str, float]]:
    if not isinstance(raw_assignments, Mapping):
        return {}
    assignments: Dict[str, Dict[str, float]] = {}
    for raw_month, raw_values in raw_assignments.items():
        month_key = _coerce_month_key(raw_month)
        if not month_key or not isinstance(raw_values, Mapping):
            continue
        clean_values = _clean_money_map(raw_values)
        if clean_values:
            assignments[month_key] = clean_values
    return assignments


def _clean_money_map(raw: Any) -> Dict[str, float]:
    if not isinstance(raw, Mapping):
        return {}
    result = {}
    for key, value in raw.items():
        category = _clean_category(key)
        amount = _money(value)
        if category and amount > 0:
            result[category] = amount
    return result


def _fixed_totals_by_category(occurrences: Iterable[Mapping[str, Any]]) -> Dict[str, float]:
    totals: Dict[str, float] = defaultdict(float)
    for occurrence in occurrences:
        category = _clean_category(occurrence.get("category")) or DEFAULT_FIXED_CATEGORY
        totals[category] += _money(occurrence.get("amount"))
    return {category: round(amount, 2) for category, amount in totals.items()}


def _rollover_enabled(setting: Mapping[str, Any], config: Mapping[str, Any], has_sinking_fund: bool = False) -> bool:
    if "rollover" in setting:
        return _to_bool(setting.get("rollover"))
    if any(key in setting for key in ("annual_amount", "target_amount", "irregular_amount", "due_date")):
        return True
    if has_sinking_fund:
        return True
    return bool(config.get("default_rollover", False))


def _opening_balance(category: str, setting: Mapping[str, Any], config: Mapping[str, Any]) -> float:
    if category in config["carryover"]:
        return _money(config["carryover"][category])
    if "carryover" in setting:
        return _money(setting.get("carryover"))
    if "starting_balance" in setting:
        return _money(setting.get("starting_balance"))
    return 0.0


def _progress_status(assigned: float, spent: float, remaining: float, pct_spent: float) -> str:
    if assigned <= 0 and spent > 0:
        return "no-budget"
    if remaining < 0:
        return "over"
    if pct_spent >= 80:
        return "warning"
    return "ok"


def _status_label(status: str) -> str:
    return {
        "ok": "On track",
        "warning": "Watch",
        "over": "Over",
        "no-budget": "No budget",
    }.get(status, "On track")


def _row_sort_key(row: Mapping[str, Any]) -> Tuple[int, str, str]:
    order = {"over": 0, "warning": 1, "no-budget": 2, "ok": 3}
    return (order.get(str(row.get("status")), 4), str(row.get("group") or ""), str(row.get("category") or "").casefold())


def _iter_billing_periods(start: date, end_start: date) -> Iterable[Tuple[date, date]]:
    if end_start < start:
        return []
    periods = []
    cursor = start
    while cursor <= end_start:
        periods.append((cursor, _add_months(cursor, 1) - timedelta(days=1)))
        cursor = _add_months(cursor, 1)
    return periods


def _period_start_from_month_key(month_key: str, billing_start_day: int) -> date:
    year, month = (int(part) for part in month_key.split("-", 1))
    day = min(_billing_start_day(billing_start_day), monthrange(year, month)[1])
    return date(year, month, day)


def _coerce_month_key(value: Any) -> str:
    if isinstance(value, str):
        value = value.strip()
        if len(value) >= 7:
            value = value[:7]
            try:
                year, month = (int(part) for part in value.split("-", 1))
                date(year, month, 1)
                return f"{year:04d}-{month:02d}"
            except (TypeError, ValueError):
                return ""
    parsed = _as_date(value)
    return _month_key(parsed) if parsed else ""


def _as_date(value: Any) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        raw = value.strip()
        if not raw:
            return None
        try:
            if len(raw) == 7:
                return date.fromisoformat(raw + "-01")
            return date.fromisoformat(raw[:10])
        except ValueError:
            return None
    return None


def _add_months(src: date, months: int) -> date:
    month = src.month - 1 + months
    year = src.year + month // 12
    month = month % 12 + 1
    day = min(src.day, monthrange(year, month)[1])
    return date(year, month, day)


def _month_diff(start: date, end: date) -> int:
    return (end.year - start.year) * 12 + (end.month - start.month)


def _month_key(value: date) -> str:
    return f"{value.year:04d}-{value.month:02d}"


def _money(value: Any) -> float:
    try:
        if value in (None, ""):
            return 0.0
        if isinstance(value, str):
            value = value.replace(",", "").strip()
            if not value:
                return 0.0
        amount = float(value)
        return round(amount, 2) if amount > 0 else 0.0
    except (TypeError, ValueError):
        return 0.0


def _int_or_none(value: Any) -> Optional[int]:
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _to_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "on", "y"}
    return bool(value)


def _billing_start_day(value: Any, max_day: int = 28) -> int:
    try:
        day = int(value)
    except (TypeError, ValueError):
        day = 1
    return min(max(day, 1), max_day)


def _clean_category(value: Any) -> str:
    return str(value or "").strip()
