"""Pure recurring calendar helpers for SpendSight.

The helpers in this module accept SpendSight's existing JSON data dict and
return JSON-safe view models. They do not import Flask, read files, write
files, or mutate the input data.
"""

from __future__ import annotations

from calendar import monthrange
from collections import defaultdict
from datetime import date, datetime, timedelta
import re
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple


DEFAULT_MONTHS_AHEAD = 3
MAX_MONTHS_AHEAD = 24
DEFAULT_FIXED_CATEGORY = "EMI / Finance"
GOAL_CATEGORY = "Savings Goals"

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

_STATUS_ASSUMPTIONS = [
    "Future due items are marked unpaid until matching payment evidence exists.",
    "Past/current due items are marked paid only when explicit paid metadata or a matching expense is found.",
    "Goal contributions are treated as unpaid planned transfers unless goal payment metadata or a linked expense exists.",
    "Cash-after-bills subtracts all scheduled recurring items in the visible range, including savings goals.",
]


def build_recurring_calendar(
    data: Mapping[str, Any],
    ref_today: Optional[date] = None,
    months_ahead: int = DEFAULT_MONTHS_AHEAD,
    start_date: Optional[Any] = None,
    end_date: Optional[Any] = None,
    subscription_insights: Optional[Any] = None,
) -> Dict[str, Any]:
    """Build a bounded recurring calendar view model.

    Sources consumed from SpendSight's JSON data dict:
    - ``fixed_expenses`` from the Income page.
    - detected monthly subscriptions from repeated ``expenses`` or supplied
      ``subscription_insights``.
    - ``goals`` with ``monthly_contribution``.
    - optional trial/renewal metadata under ``recurring_calendar``,
      ``trial_renewals``, ``trials``, ``renewals``, or subscription records.

    The horizon is bounded to ``MAX_MONTHS_AHEAD`` to prevent accidental
    unbounded recurrence expansion.
    """

    ref = _as_date(ref_today) or date.today()
    start = _as_date(start_date) or ref
    end = _resolve_end_date(start, months_ahead, end_date)

    expenses = _normalise_expenses(data.get("expenses", []))
    subscriptions = _subscription_candidates(data, subscription_insights, ref)

    items: List[Dict[str, Any]] = []
    items.extend(_fixed_expense_items(data.get("fixed_expenses", []), start, end, expenses, ref))
    items.extend(_subscription_items(subscriptions, start, end, expenses, ref))
    items.extend(_goal_contribution_items(data, start, end, expenses, ref))
    items.extend(_trial_renewal_items(data, subscriptions, start, end, expenses, ref))

    items = _dedupe_items(items)
    items.sort(key=lambda item: (item["due_on"], item["source_type"], item["name"].casefold(), item["amount"]))

    calendar_rows = monthly_calendar_rows(data, items, start, end, ref_today=ref)

    return {
        "range": {
            "start": start.isoformat(),
            "end": end.isoformat(),
            "months_ahead": _bounded_month_count(months_ahead),
            "max_months_ahead": MAX_MONTHS_AHEAD,
            "is_bounded": True,
        },
        "items": items,
        "calendar_rows": calendar_rows,
        "summary": _calendar_summary(items, calendar_rows, ref),
        "assumptions": list(_STATUS_ASSUMPTIONS),
    }


def monthly_calendar_rows(
    data: Mapping[str, Any],
    items: Sequence[Mapping[str, Any]],
    start_date: Any,
    end_date: Any,
    ref_today: Optional[date] = None,
) -> List[Dict[str, Any]]:
    """Return month rows with day/week buckets and cash-after-bills totals."""

    start = _as_date(start_date)
    end = _as_date(end_date)
    if not start or not end or end < start:
        return []

    ref = _as_date(ref_today) or date.today()
    by_date: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for item in items or []:
        due = _as_date(item.get("due_on"))
        if due and start <= due <= end:
            by_date[due.isoformat()].append(dict(item))

    rows: List[Dict[str, Any]] = []
    cursor = date(start.year, start.month, 1)
    last_month = date(end.year, end.month, 1)

    while cursor <= last_month:
        month_start = cursor
        month_end = date(cursor.year, cursor.month, monthrange(cursor.year, cursor.month)[1])
        visible_start = max(start, month_start)
        visible_end = min(end, month_end)
        month_key = _month_key(cursor)

        month_items = [
            dict(item)
            for item in items or []
            if _item_due_in_range(item, visible_start, visible_end)
        ]
        month_items.sort(key=lambda item: (item["due_on"], item["source_type"], item["name"].casefold()))

        day_rows = []
        day = month_start
        while day <= month_end:
            due_items = by_date.get(day.isoformat(), [])
            due_items.sort(key=lambda item: (item["source_type"], item["name"].casefold(), item["amount"]))
            day_rows.append(
                {
                    "date": day.isoformat(),
                    "day": day.day,
                    "weekday": day.strftime("%a"),
                    "in_range": visible_start <= day <= visible_end,
                    "is_today": day == ref,
                    "items": due_items,
                    "item_count": len(due_items),
                    "scheduled_total": _sum_amount(due_items),
                    "paid_total": _sum_amount(item for item in due_items if item.get("status") == "paid"),
                    "unpaid_total": _sum_amount(item for item in due_items if item.get("status") == "unpaid"),
                }
            )
            day += timedelta(days=1)

        income = _income_for_month(data, month_start)
        scheduled_total = _sum_amount(month_items)
        paid_total = _sum_amount(item for item in month_items if item.get("status") == "paid")
        unpaid_total = _sum_amount(item for item in month_items if item.get("status") == "unpaid")

        rows.append(
            {
                "month": month_key,
                "label": cursor.strftime("%B %Y"),
                "start": visible_start.isoformat(),
                "end": visible_end.isoformat(),
                "is_partial_month": visible_start != month_start or visible_end != month_end,
                "income": income,
                "scheduled_total": scheduled_total,
                "paid_total": paid_total,
                "unpaid_total": unpaid_total,
                "cash_after_bills": round(income["total"] - scheduled_total, 2),
                "cash_after_unpaid_bills": round(income["total"] - unpaid_total, 2),
                "item_count": len(month_items),
                "paid_count": sum(1 for item in month_items if item.get("status") == "paid"),
                "unpaid_count": sum(1 for item in month_items if item.get("status") == "unpaid"),
                "items": month_items,
                "days": day_rows,
                "weeks": _week_rows(day_rows),
            }
        )

        cursor = _add_months(cursor, 1)

    return rows


def fixed_expense_occurrences(
    fixed_expenses: Iterable[Mapping[str, Any]],
    start_date: Any,
    end_date: Any,
) -> List[Dict[str, Any]]:
    """Return fixed/EMI occurrences due inside the inclusive date range."""

    start = _as_date(start_date)
    end = _as_date(end_date)
    if not start or not end or end < start:
        return []

    occurrences: List[Dict[str, Any]] = []
    first_month = date(start.year, start.month, 1)
    last_month = date(end.year, end.month, 1)

    for record in fixed_expenses or []:
        if not isinstance(record, Mapping):
            continue
        amount = _money(record.get("amount"))
        if amount <= 0:
            continue
        fixed_start = _fixed_start_month(record)
        if not fixed_start:
            continue

        item_type = str(record.get("type") or "fixed").strip().lower()
        frequency = str(record.get("frequency") or "monthly").strip().lower()
        interval = 1 if item_type == "emi" else _FREQUENCY_MONTHS.get(frequency, 1)
        total_months = _positive_int(record.get("total_months")) if item_type == "emi" else 0
        end_limit = _as_date(record.get("end_date"))
        day_of_month = _day_of_month(record.get("day_of_month", 1))

        cursor = max(first_month, fixed_start)
        while cursor <= last_month:
            months_since = _month_diff(fixed_start, cursor)
            applies = False
            if months_since >= 0:
                if item_type == "emi":
                    applies = total_months > 0 and months_since < total_months
                else:
                    applies = months_since % interval == 0
            if applies:
                due = date(cursor.year, cursor.month, min(day_of_month, monthrange(cursor.year, cursor.month)[1]))
                if end_limit and due > end_limit:
                    break
                if start <= due <= end:
                    occurrences.append(
                        {
                            "source_id": str(record.get("id") or _slug(record.get("name") or "fixed")),
                            "name": str(record.get("name") or "Scheduled expense"),
                            "amount": round(amount, 2),
                            "category": _clean_text(record.get("category")) or DEFAULT_FIXED_CATEGORY,
                            "payment_method": str(record.get("payment_method") or ""),
                            "due_on": due.isoformat(),
                            "type": item_type,
                            "cadence": frequency if item_type != "emi" else "monthly",
                            "interval_months": interval,
                        }
                    )
            cursor = _add_months(cursor, 1)

    occurrences.sort(key=lambda item: (item["due_on"], item["name"].casefold()))
    return occurrences


def detect_monthly_subscriptions(data: Mapping[str, Any], ref_today: Optional[date] = None) -> Dict[str, Any]:
    """Detect monthly subscriptions from repeated expense history.

    This mirrors SpendSight's current product helper shape closely enough for
    draft wiring while staying independent from ``app.py``.
    """

    ref = _as_date(ref_today) or date.today()
    grouped: Dict[str, List[Dict[str, Any]]] = defaultdict(list)

    for expense in data.get("expenses", []) or []:
        if not isinstance(expense, Mapping):
            continue
        tx_date = _as_date(expense.get("date"))
        amount = _money(expense.get("amount"))
        if not tx_date or amount <= 0:
            continue
        merchant = _merchant_key(expense)
        grouped[_normalise_key(merchant)].append(
            {
                **dict(expense),
                "_date": tx_date,
                "_amount": amount,
                "_name": merchant,
            }
        )

    subscriptions: List[Dict[str, Any]] = []
    for rows in grouped.values():
        rows.sort(key=lambda item: item["_date"])
        months_seen = {(row["_date"].year, row["_date"].month) for row in rows}
        if len(rows) < 3 or len(months_seen) < 3:
            continue
        gaps = [(rows[i]["_date"] - rows[i - 1]["_date"]).days for i in range(1, len(rows))]
        monthly_gaps = [gap for gap in gaps if 25 <= gap <= 35]
        if len(monthly_gaps) < max(2, len(gaps) - 1):
            continue

        latest = rows[-1]
        amounts = [row["_amount"] for row in rows]
        next_due = _add_months(latest["_date"], 1)
        while next_due < ref:
            next_due = _add_months(next_due, 1)

        item: Dict[str, Any] = {
            "id": str(latest.get("subscription_id") or latest.get("recurring_id") or _slug(latest["_name"])),
            "name": latest["_name"],
            "category": str(latest.get("category") or "Other"),
            "subcategory": str(latest.get("subcategory") or latest["_name"]),
            "payment_method": str(latest.get("payment_method") or ""),
            "cadence": "monthly",
            "occurrences": len(rows),
            "latest_amount": round(latest["_amount"], 2),
            "average_amount": round(sum(amounts) / len(amounts), 2),
            "first_seen": rows[0]["_date"].isoformat(),
            "last_seen": latest["_date"].isoformat(),
            "next_due_on": next_due.isoformat(),
            "confidence": round(min(0.98, 0.55 + (len(monthly_gaps) * 0.12)), 2),
        }
        price_change = _price_change(amounts)
        if price_change:
            item["price_change"] = price_change
        subscriptions.append(item)

    subscriptions.sort(key=lambda item: (item["next_due_on"], -item["latest_amount"], item["name"].casefold()))
    return {
        "subscriptions": subscriptions,
        "count": len(subscriptions),
        "monthly_total": round(sum(item["latest_amount"] for item in subscriptions), 2),
        "next_30_days_total": round(
            sum(
                item["latest_amount"]
                for item in subscriptions
                if _as_date(item.get("next_due_on")) and _as_date(item["next_due_on"]) <= ref + timedelta(days=30)
            ),
            2,
        ),
    }


def _fixed_expense_items(
    fixed_expenses: Any,
    start: date,
    end: date,
    expenses: Sequence[Mapping[str, Any]],
    ref: date,
) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    records_by_id: Dict[str, Mapping[str, Any]] = {}
    for record in fixed_expenses or []:
        if isinstance(record, Mapping):
            source_id = str(record.get("id") or _slug(record.get("name") or "fixed"))
            records_by_id[source_id] = record

    for occurrence in fixed_expense_occurrences(fixed_expenses, start, end):
        due = _as_date(occurrence["due_on"])
        if not due:
            continue
        record = records_by_id.get(occurrence["source_id"], {})
        status = _payment_status(
            record,
            source_type="fixed_expense",
            source_id=occurrence["source_id"],
            name=occurrence["name"],
            category=occurrence["category"],
            amount=occurrence["amount"],
            due_on=due,
            expenses=expenses,
            ref=ref,
        )
        items.append(
            _make_item(
                source_type="fixed_expense",
                source_id=occurrence["source_id"],
                name=occurrence["name"],
                category=occurrence["category"],
                amount=occurrence["amount"],
                due_on=due,
                cadence=occurrence["cadence"],
                payment_method=occurrence["payment_method"],
                confidence=1.0,
                metadata={
                    "fixed_type": occurrence["type"],
                    "interval_months": occurrence["interval_months"],
                },
                status=status,
            )
        )
    return items


def _subscription_items(
    subscriptions: Sequence[Mapping[str, Any]],
    start: date,
    end: date,
    expenses: Sequence[Mapping[str, Any]],
    ref: date,
) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []

    for record in subscriptions or []:
        if not isinstance(record, Mapping):
            continue
        name = str(record.get("name") or record.get("merchant") or record.get("subcategory") or "Subscription")
        amount = _first_money(record, "latest_amount", "monthly_amount", "amount", "average_amount")
        if amount <= 0:
            continue
        next_due = _as_date(record.get("next_due_on") or record.get("due_on"))
        if not next_due:
            last_seen = _as_date(record.get("last_seen") or record.get("last_paid_on"))
            next_due = _add_months(last_seen, 1) if last_seen else None
        if not next_due:
            continue

        cadence = str(record.get("cadence") or record.get("frequency") or "monthly").strip().lower()
        interval = _FREQUENCY_MONTHS.get(cadence, 1)
        end_limit = _as_date(record.get("end_date"))
        source_id = str(record.get("id") or record.get("source_id") or _slug(name))
        due = next_due
        while due < start:
            due = _add_months(due, interval)
        while due <= end:
            if end_limit and due > end_limit:
                break
            status = _payment_status(
                record,
                source_type="subscription",
                source_id=source_id,
                name=name,
                category=str(record.get("category") or "Subscriptions"),
                amount=amount,
                due_on=due,
                expenses=expenses,
                ref=ref,
            )
            items.append(
                _make_item(
                    source_type="subscription",
                    source_id=source_id,
                    name=name,
                    category=str(record.get("category") or "Subscriptions"),
                    amount=amount,
                    due_on=due,
                    cadence=cadence or "monthly",
                    payment_method=str(record.get("payment_method") or ""),
                    confidence=_money(record.get("confidence"), default=0.75),
                    metadata={
                        "first_seen": str(record.get("first_seen") or ""),
                        "last_seen": str(record.get("last_seen") or ""),
                        "occurrences": _positive_int(record.get("occurrences")),
                        "price_change": record.get("price_change"),
                    },
                    status=status,
                )
            )
            due = _add_months(due, interval)

    return items


def _goal_contribution_items(
    data: Mapping[str, Any],
    start: date,
    end: date,
    expenses: Sequence[Mapping[str, Any]],
    ref: date,
) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    config = data.get("recurring_calendar", {})
    if not isinstance(config, Mapping):
        config = {}
    default_day = _day_of_month(config.get("goal_contribution_day", 1))

    for goal in data.get("goals", []) or []:
        if not isinstance(goal, Mapping):
            continue
        monthly = _money(goal.get("monthly_contribution"))
        if monthly <= 0:
            continue
        target = _money(goal.get("target_amount"))
        current = _money(goal.get("current_amount"))
        remaining = max(target - current, 0.0) if target > 0 else None
        if remaining == 0.0:
            continue

        name = str(goal.get("name") or "Savings goal")
        source_id = str(goal.get("id") or _slug(name))
        contribution_day = _day_of_month(goal.get("contribution_day", goal.get("day_of_month", default_day)))
        goal_start = _as_date(goal.get("start_date") or goal.get("created_at")) or start
        cursor = max(date(start.year, start.month, 1), date(goal_start.year, goal_start.month, 1))
        last_month = date(end.year, end.month, 1)

        while cursor <= last_month:
            due = date(cursor.year, cursor.month, min(contribution_day, monthrange(cursor.year, cursor.month)[1]))
            if due < goal_start:
                cursor = _add_months(cursor, 1)
                continue
            if start <= due <= end:
                amount = monthly if remaining is None else min(monthly, max(remaining, 0.0))
                if amount <= 0:
                    break
                status = _payment_status(
                    goal,
                    source_type="goal_contribution",
                    source_id=source_id,
                    name=name,
                    category=GOAL_CATEGORY,
                    amount=amount,
                    due_on=due,
                    expenses=expenses,
                    ref=ref,
                )
                items.append(
                    _make_item(
                        source_type="goal_contribution",
                        source_id=source_id,
                        name=name,
                        category=GOAL_CATEGORY,
                        amount=amount,
                        due_on=due,
                        cadence="monthly",
                        payment_method=str(goal.get("payment_method") or ""),
                        confidence=1.0,
                        metadata={
                            "target_amount": round(target, 2),
                            "current_amount": round(current, 2),
                            "remaining_before_calendar": round(remaining, 2) if remaining is not None else None,
                            "priority": str(goal.get("priority") or ""),
                        },
                        status=status,
                    )
                )
                if remaining is not None:
                    remaining = round(max(remaining - amount, 0.0), 2)
            cursor = _add_months(cursor, 1)

    return items


def _trial_renewal_items(
    data: Mapping[str, Any],
    subscriptions: Sequence[Mapping[str, Any]],
    start: date,
    end: date,
    expenses: Sequence[Mapping[str, Any]],
    ref: date,
) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []

    for record, source_name, kind_hint in _trial_renewal_records(data, subscriptions):
        if not isinstance(record, Mapping):
            continue
        source_id = str(record.get("id") or record.get("source_id") or _slug(record.get("name") or source_name))
        base_name = str(record.get("name") or record.get("merchant") or record.get("subcategory") or "Recurring item")
        category = str(record.get("category") or "Subscriptions")
        payment_method = str(record.get("payment_method") or "")

        trial_date = _first_date(record, "trial_ends_on", "trial_end_date", "trial_end", "trial_expires_on", "ends_on")
        if not trial_date and kind_hint == "trial":
            trial_date = _first_date(record, "due_on", "date", "end_date")
        if trial_date and start <= trial_date <= end:
            amount = _first_money(record, "renewal_amount", "trial_amount", "amount", "monthly_amount", "latest_amount")
            status = _payment_status(
                record,
                source_type="trial",
                source_id=source_id,
                name=base_name,
                category=category,
                amount=amount,
                due_on=trial_date,
                expenses=expenses,
                ref=ref,
            )
            items.append(
                _make_item(
                    source_type="trial",
                    source_id=source_id,
                    name=f"{base_name} trial ends",
                    category=category,
                    amount=amount,
                    due_on=trial_date,
                    cadence="one_time",
                    payment_method=payment_method,
                    confidence=1.0,
                    metadata={
                        "cancel_by": _date_iso(record.get("cancel_by") or record.get("cancel_before")),
                        "metadata_source": source_name,
                    },
                    status=status,
                )
            )

        renewal_date = _first_date(record, "renewal_date", "renews_on", "next_renewal_on", "contract_renewal_on")
        if not renewal_date and kind_hint == "renewal":
            renewal_date = _first_date(record, "due_on", "date")
        if renewal_date and start <= renewal_date <= end:
            amount = _first_money(record, "renewal_amount", "amount", "monthly_amount", "latest_amount")
            status = _payment_status(
                record,
                source_type="renewal",
                source_id=source_id,
                name=base_name,
                category=category,
                amount=amount,
                due_on=renewal_date,
                expenses=expenses,
                ref=ref,
            )
            items.append(
                _make_item(
                    source_type="renewal",
                    source_id=source_id,
                    name=f"{base_name} renewal",
                    category=category,
                    amount=amount,
                    due_on=renewal_date,
                    cadence=str(record.get("cadence") or record.get("frequency") or "one_time"),
                    payment_method=payment_method,
                    confidence=1.0,
                    metadata={
                        "cancel_by": _date_iso(record.get("cancel_by") or record.get("cancel_before")),
                        "metadata_source": source_name,
                    },
                    status=status,
                )
            )

    return items


def _trial_renewal_records(
    data: Mapping[str, Any],
    subscriptions: Sequence[Mapping[str, Any]],
) -> Iterable[Tuple[Mapping[str, Any], str, str]]:
    config = data.get("recurring_calendar", {})
    if isinstance(config, Mapping):
        for record in _as_list(config.get("trial_renewals")):
            yield record, "recurring_calendar.trial_renewals", str(record.get("type") or "")

    for key, kind in (
        ("trial_renewals", ""),
        ("trials", "trial"),
        ("renewals", "renewal"),
        ("recurring_payments", ""),
        ("fixed_expenses", ""),
    ):
        for record in _as_list(data.get(key)):
            yield record, key, kind or str(record.get("type") or "")

    for record in subscriptions or []:
        yield record, "detected_subscriptions", str(record.get("type") or "")


def _payment_status(
    record: Mapping[str, Any],
    *,
    source_type: str,
    source_id: str,
    name: str,
    category: str,
    amount: float,
    due_on: date,
    expenses: Sequence[Mapping[str, Any]],
    ref: date,
) -> Dict[str, Any]:
    explicit = _explicit_paid_status(record, due_on)
    if explicit:
        return explicit

    match = _matching_expense(
        expenses,
        source_type=source_type,
        source_id=source_id,
        name=name,
        category=category,
        amount=amount,
        due_on=due_on,
    )
    if match:
        return {
            "status": "paid",
            "is_paid": True,
            "status_assumption": "matched_expense",
            "status_reason": "A SpendSight expense in the same month matched the source, name/category, and amount.",
            "paid_on": match["date"].isoformat(),
            "matched_expense_id": str(match.get("id") or ""),
        }

    if source_type == "goal_contribution":
        reason = "No linked contribution evidence was found for this savings goal."
        assumption = "goal_contribution_assumed_unpaid"
    elif due_on > ref:
        reason = "Future due item; no payment evidence is expected yet."
        assumption = "future_due_unpaid"
    else:
        reason = "Due date has passed or is today, but no matching expense or paid metadata was found."
        assumption = "past_due_no_match"

    return {
        "status": "unpaid",
        "is_paid": False,
        "status_assumption": assumption,
        "status_reason": reason,
        "paid_on": "",
        "matched_expense_id": "",
    }


def _explicit_paid_status(record: Mapping[str, Any], due_on: date) -> Optional[Dict[str, Any]]:
    month = _month_key(due_on)

    if _month_in_values(record.get("paid_months"), month):
        return _explicit_status("paid", "explicit_paid_month", f"Paid month metadata includes {month}.")
    if _date_in_month_values(record.get("paid_dates"), month):
        return _explicit_status("paid", "explicit_paid_date", f"Paid date metadata falls in {month}.")
    if _date_in_month_values(record.get("paid_occurrences"), month):
        return _explicit_status("paid", "explicit_paid_occurrence", f"Paid occurrence metadata falls in {month}.")

    for key in ("payments", "payment_history", "contributions"):
        for payment in _as_list(record.get(key)):
            if not isinstance(payment, Mapping):
                continue
            paid_date = _first_date(payment, "paid_on", "paid_date", "date", "contributed_on")
            if paid_date and _month_key(paid_date) == month:
                status = str(payment.get("status") or "paid").strip().lower()
                if status not in {"unpaid", "missed", "failed", "cancelled", "canceled"}:
                    return _explicit_status("paid", f"explicit_{key}", f"{key} metadata contains a payment in {month}.")

    last_paid = _first_date(record, "last_paid_on", "paid_on", "paid_date")
    if last_paid and _month_key(last_paid) == month:
        return _explicit_status("paid", "explicit_last_paid_on", f"Last paid metadata falls in {month}.")

    payment_status = str(record.get("payment_status") or "").strip().lower()
    if payment_status in {"unpaid", "missed", "failed", "past_due", "past-due"}:
        return _explicit_status("unpaid", "explicit_unpaid_status", "Payment status metadata is unpaid.")

    return None


def _explicit_status(status: str, assumption: str, reason: str) -> Dict[str, Any]:
    return {
        "status": status,
        "is_paid": status == "paid",
        "status_assumption": assumption,
        "status_reason": reason,
        "paid_on": "",
        "matched_expense_id": "",
    }


def _matching_expense(
    expenses: Sequence[Mapping[str, Any]],
    *,
    source_type: str,
    source_id: str,
    name: str,
    category: str,
    amount: float,
    due_on: date,
) -> Optional[Mapping[str, Any]]:
    due_month = _month_key(due_on)
    link_fields = _link_fields_for_source(source_type)
    name_key = _normalise_key(name)
    category_key = _normalise_key(category)

    for expense in expenses:
        exp_date = expense.get("date")
        if not isinstance(exp_date, date) or _month_key(exp_date) != due_month:
            continue
        direct_link = source_id and any(str(expense.get(field) or "") == source_id for field in link_fields)
        if direct_link and _amount_close(expense.get("amount", 0), amount):
            return expense

        if not _amount_close(expense.get("amount", 0), amount):
            continue
        days_from_due = abs((exp_date - due_on).days)
        if days_from_due > 10:
            continue

        text = str(expense.get("text") or "")
        name_match = bool(name_key and name_key in text)
        category_match = bool(category_key and _normalise_key(expense.get("category")) == category_key)
        if name_match or (category_match and source_type in {"fixed_expense", "goal_contribution", "renewal"}):
            return expense

    return None


def _link_fields_for_source(source_type: str) -> Tuple[str, ...]:
    if source_type == "fixed_expense":
        return ("fixed_expense_id", "emi_id", "recurring_id", "recurring_payment_id", "source_id")
    if source_type == "subscription":
        return ("subscription_id", "recurring_id", "recurring_payment_id", "source_id")
    if source_type == "goal_contribution":
        return ("goal_id", "source_id")
    if source_type == "trial":
        return ("trial_id", "recurring_id", "source_id")
    if source_type == "renewal":
        return ("renewal_id", "recurring_id", "source_id")
    return ("source_id",)


def _make_item(
    *,
    source_type: str,
    source_id: str,
    name: str,
    category: str,
    amount: float,
    due_on: date,
    cadence: str,
    payment_method: str,
    confidence: float,
    metadata: Mapping[str, Any],
    status: Mapping[str, Any],
) -> Dict[str, Any]:
    item_id = f"{source_type}:{source_id}:{due_on.isoformat()}"
    return {
        "id": item_id,
        "source_type": source_type,
        "source_id": source_id,
        "name": name,
        "category": category,
        "amount": round(_money(amount), 2),
        "due_on": due_on.isoformat(),
        "month": _month_key(due_on),
        "cadence": cadence,
        "payment_method": payment_method,
        "confidence": round(float(confidence or 0), 2),
        "status": status["status"],
        "is_paid": bool(status["is_paid"]),
        "status_assumption": status["status_assumption"],
        "status_reason": status["status_reason"],
        "paid_on": status.get("paid_on", ""),
        "matched_expense_id": status.get("matched_expense_id", ""),
        "metadata": _json_safe_metadata(metadata),
    }


def _subscription_candidates(data: Mapping[str, Any], supplied: Optional[Any], ref: date) -> List[Dict[str, Any]]:
    candidates: List[Mapping[str, Any]] = []
    for payload in (
        supplied,
        data.get("subscription_insights"),
        data.get("detected_subscriptions"),
        data.get("subscriptions"),
    ):
        candidates.extend(_subscription_list(payload))
    candidates.extend(detect_monthly_subscriptions(data, ref_today=ref)["subscriptions"])

    deduped: Dict[Tuple[str, str], Dict[str, Any]] = {}
    for item in candidates:
        if not isinstance(item, Mapping):
            continue
        name = str(item.get("name") or item.get("merchant") or item.get("subcategory") or "").strip()
        if not name:
            continue
        key = (_normalise_key(name), str(item.get("payment_method") or "").casefold())
        existing = deduped.get(key)
        if not existing:
            deduped[key] = dict(item)
            continue
        if _as_date(item.get("next_due_on")) and not _as_date(existing.get("next_due_on")):
            deduped[key] = dict(item)
        elif _money(item.get("confidence")) > _money(existing.get("confidence")):
            deduped[key] = dict(item)
    return list(deduped.values())


def _subscription_list(payload: Any) -> List[Mapping[str, Any]]:
    if isinstance(payload, Mapping):
        for key in ("subscriptions", "items", "detected"):
            value = payload.get(key)
            if isinstance(value, list):
                return [item for item in value if isinstance(item, Mapping)]
        return []
    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, Mapping)]
    return []


def _calendar_summary(items: Sequence[Mapping[str, Any]], rows: Sequence[Mapping[str, Any]], ref: date) -> Dict[str, Any]:
    scheduled_total = _sum_amount(items)
    paid_total = _sum_amount(item for item in items if item.get("status") == "paid")
    unpaid_items = [item for item in items if item.get("status") == "unpaid"]
    source_counts: Dict[str, int] = defaultdict(int)
    for item in items:
        source_counts[str(item.get("source_type") or "unknown")] += 1

    return {
        "item_count": len(items),
        "scheduled_total": scheduled_total,
        "paid_total": paid_total,
        "unpaid_total": _sum_amount(unpaid_items),
        "past_due_unpaid_total": _sum_amount(
            item for item in unpaid_items if (_as_date(item.get("due_on")) or ref) <= ref
        ),
        "next_30_days_unpaid_total": _sum_amount(
            item
            for item in unpaid_items
            if (_as_date(item.get("due_on")) or ref) <= ref + timedelta(days=30)
        ),
        "cash_after_bills_total": round(sum(_number(row.get("cash_after_bills")) for row in rows), 2),
        "cash_after_unpaid_bills_total": round(
            sum(_number(row.get("cash_after_unpaid_bills")) for row in rows),
            2,
        ),
        "source_counts": dict(sorted(source_counts.items())),
        "assumption_count": len(_STATUS_ASSUMPTIONS),
    }


def _income_for_month(data: Mapping[str, Any], month_start: date) -> Dict[str, float]:
    salary = _salary_for_date(data.get("income", {}), month_start)
    extra = _extra_income_for_month(data.get("extra_income", []), month_start)
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
    return _money(income_data.get("monthly_salary"))


def _extra_income_for_month(extra_income: Any, month_start: date) -> float:
    month_end = date(month_start.year, month_start.month, monthrange(month_start.year, month_start.month)[1])
    total = 0.0
    for item in extra_income or []:
        if not isinstance(item, Mapping):
            continue
        amount = _money(item.get("amount"))
        if amount <= 0:
            continue
        item_type = str(item.get("type") or "one-time").strip().lower()
        if item_type == "recurring":
            start = _as_date(item.get("start_date") or item.get("date"))
            if not start:
                continue
            end = _as_date(item.get("end_date")) or month_end
            if end < month_start:
                continue
            interval = _FREQUENCY_MONTHS.get(str(item.get("frequency") or "monthly").strip().lower(), 1)
            months_since = _month_diff(date(start.year, start.month, 1), month_start)
            if months_since >= 0 and months_since % interval == 0:
                occurrence = date(month_start.year, month_start.month, min(start.day, monthrange(month_start.year, month_start.month)[1]))
                if start <= occurrence <= end:
                    total += amount
        else:
            paid_on = _as_date(item.get("date"))
            if paid_on and month_start <= paid_on <= month_end:
                total += amount
    return round(total, 2)


def _week_rows(day_rows: Sequence[Mapping[str, Any]]) -> List[Dict[str, Any]]:
    if not day_rows:
        return []
    first = _as_date(day_rows[0].get("date"))
    if not first:
        return []

    weeks: List[Dict[str, Any]] = []
    current: List[Dict[str, Any]] = []
    for _ in range(first.weekday()):
        current.append(_blank_day())
    for day in day_rows:
        current.append({**dict(day), "in_month": True})
        if len(current) == 7:
            weeks.append(_week_summary(current))
            current = []
    if current:
        while len(current) < 7:
            current.append(_blank_day())
        weeks.append(_week_summary(current))
    return weeks


def _week_summary(days: Sequence[Mapping[str, Any]]) -> Dict[str, Any]:
    return {
        "days": list(days),
        "scheduled_total": _sum_amount(day.get("scheduled_total") for day in days),
        "paid_total": _sum_amount(day.get("paid_total") for day in days),
        "unpaid_total": _sum_amount(day.get("unpaid_total") for day in days),
        "item_count": sum(int(day.get("item_count") or 0) for day in days),
    }


def _blank_day() -> Dict[str, Any]:
    return {
        "date": "",
        "day": None,
        "weekday": "",
        "in_range": False,
        "in_month": False,
        "is_today": False,
        "items": [],
        "item_count": 0,
        "scheduled_total": 0.0,
        "paid_total": 0.0,
        "unpaid_total": 0.0,
    }


def _normalise_expenses(expenses: Any) -> List[Dict[str, Any]]:
    normalised: List[Dict[str, Any]] = []
    for expense in expenses or []:
        if not isinstance(expense, Mapping):
            continue
        tx_date = _as_date(expense.get("date"))
        if not tx_date:
            continue
        text = " ".join(
            str(expense.get(key) or "")
            for key in ("name", "merchant", "category", "subcategory", "notes", "payment_method")
        )
        normalised.append(
            {
                **dict(expense),
                "id": str(expense.get("id") or ""),
                "date": tx_date,
                "amount": _money(expense.get("amount")),
                "category": str(expense.get("category") or ""),
                "text": _normalise_key(text),
            }
        )
    return normalised


def _fixed_start_month(record: Mapping[str, Any]) -> Optional[date]:
    start_date = _as_date(record.get("start_date"))
    if start_date:
        return date(start_date.year, start_date.month, 1)
    try:
        year = int(record.get("start_year", 0))
        month = int(record.get("start_month", 0))
        return date(year, month, 1)
    except (TypeError, ValueError):
        return None


def _resolve_end_date(start: date, months_ahead: int, end_date: Optional[Any]) -> date:
    default_end = _add_months(start, _bounded_month_count(months_ahead)) - timedelta(days=1)
    requested_end = _as_date(end_date) or default_end
    max_end = _add_months(start, MAX_MONTHS_AHEAD) - timedelta(days=1)
    end = min(requested_end, max_end)
    return end if end >= start else start


def _bounded_month_count(value: Any) -> int:
    try:
        months = int(value)
    except (TypeError, ValueError):
        months = DEFAULT_MONTHS_AHEAD
    return max(1, min(MAX_MONTHS_AHEAD, months))


def _item_due_in_range(item: Mapping[str, Any], start: date, end: date) -> bool:
    due = _as_date(item.get("due_on"))
    return bool(due and start <= due <= end)


def _dedupe_items(items: Sequence[Mapping[str, Any]]) -> List[Dict[str, Any]]:
    deduped: Dict[Tuple[str, str, str, str, float], Dict[str, Any]] = {}
    for item in items:
        key = (
            str(item.get("source_type") or ""),
            str(item.get("source_id") or ""),
            str(item.get("due_on") or ""),
            _normalise_key(item.get("name")),
            _money(item.get("amount")),
        )
        existing = deduped.get(key)
        if not existing:
            deduped[key] = dict(item)
        elif existing.get("status") != "paid" and item.get("status") == "paid":
            deduped[key] = dict(item)
    return list(deduped.values())


def _merchant_key(expense: Mapping[str, Any]) -> str:
    subcategory = str(expense.get("subcategory") or "").strip()
    notes = str(expense.get("notes") or "").strip()
    category = str(expense.get("category") or "Other").strip()
    label = subcategory or notes or category
    return re.sub(r"\s+", " ", label).strip() or "Unknown"


def _price_change(amounts: Sequence[float]) -> Optional[Dict[str, Any]]:
    if len(amounts) < 2:
        return None
    latest = round(float(amounts[-1]), 2)
    previous = None
    for amount in reversed(amounts[:-1]):
        value = round(float(amount), 2)
        if value != latest:
            previous = value
            break
    if previous is None:
        return None
    delta = round(latest - previous, 2)
    return {
        "from_amount": previous,
        "to_amount": latest,
        "delta": abs(delta),
        "direction": "increase" if delta > 0 else "decrease",
    }


def _month_in_values(values: Any, month: str) -> bool:
    for value in _as_list(values):
        if isinstance(value, Mapping):
            value = value.get("month") or value.get("date") or value.get("paid_on")
        text = str(value or "").strip()
        if text[:7] == month:
            return True
    return False


def _date_in_month_values(values: Any, month: str) -> bool:
    for value in _as_list(values):
        if isinstance(value, Mapping):
            parsed = _first_date(value, "paid_on", "paid_date", "date", "due_on")
        else:
            parsed = _as_date(value)
        if parsed and _month_key(parsed) == month:
            return True
    return False


def _first_date(record: Mapping[str, Any], *keys: str) -> Optional[date]:
    for key in keys:
        parsed = _as_date(record.get(key))
        if parsed:
            return parsed
    return None


def _first_money(record: Mapping[str, Any], *keys: str) -> float:
    for key in keys:
        amount = _money(record.get(key))
        if amount > 0:
            return amount
    return 0.0


def _as_list(value: Any) -> List[Any]:
    if isinstance(value, list):
        return value
    if isinstance(value, tuple):
        return list(value)
    if value in (None, ""):
        return []
    return [value]


def _json_safe_metadata(metadata: Mapping[str, Any]) -> Dict[str, Any]:
    safe: Dict[str, Any] = {}
    for key, value in metadata.items():
        if isinstance(value, (str, int, float, bool)) or value is None:
            safe[key] = value
        elif isinstance(value, Mapping):
            safe[key] = _json_safe_metadata(value)
        elif isinstance(value, list):
            safe[key] = [
                _json_safe_metadata(item) if isinstance(item, Mapping) else item
                for item in value
                if isinstance(item, (str, int, float, bool, Mapping)) or item is None
            ]
        else:
            safe[key] = str(value)
    return safe


def _sum_amount(items: Iterable[Any]) -> float:
    return round(sum(_money(item.get("amount") if isinstance(item, Mapping) else item) for item in items), 2)


def _amount_close(actual: Any, expected: Any) -> bool:
    expected_amount = _money(expected)
    actual_amount = _money(actual)
    if expected_amount <= 0:
        return actual_amount <= 0
    return abs(actual_amount - expected_amount) <= max(1.0, expected_amount * 0.05)


def _date_iso(value: Any) -> str:
    parsed = _as_date(value)
    return parsed.isoformat() if parsed else ""


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


def _money(value: Any, default: float = 0.0) -> float:
    try:
        if value in (None, ""):
            return default
        if isinstance(value, str):
            value = value.replace(",", "").strip()
            if not value:
                return default
        amount = float(value)
        return round(amount, 2) if amount > 0 else default
    except (TypeError, ValueError):
        return default


def _number(value: Any, default: float = 0.0) -> float:
    try:
        if value in (None, ""):
            return default
        if isinstance(value, str):
            value = value.replace(",", "").strip()
            if not value:
                return default
        return round(float(value), 2)
    except (TypeError, ValueError):
        return default


def _positive_int(value: Any) -> int:
    try:
        parsed = int(float(str(value).replace(",", "").strip()))
        return parsed if parsed > 0 else 0
    except (TypeError, ValueError):
        return 0


def _day_of_month(value: Any) -> int:
    try:
        day = int(value)
    except (TypeError, ValueError):
        day = 1
    return max(1, min(31, day))


def _clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _normalise_key(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().casefold())


def _slug(value: Any) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", str(value or "").strip().casefold()).strip("-")
    return slug or "item"
