"""Receipt inbox helpers for SpendSight feature drafting.

These helpers are deliberately pure and Flask-free.  They operate on the
existing SpendSight user data dict and store draft receipt review state under
``data["receipt_inbox"]`` without changing the existing expense schema.
"""

from __future__ import annotations

import copy
import re
import uuid
from datetime import date, datetime
from difflib import SequenceMatcher
from typing import Any


INBOX_KEY = "receipt_inbox"

STATUS_NEEDS_REVIEW = "needs_review"
STATUS_DUPLICATE_CANDIDATE = "duplicate_candidate"
STATUS_READY_TO_POST = "ready_to_post"
STATUS_POSTED = "posted"
STATUS_DISMISSED = "dismissed"

VALID_STATUSES = {
    STATUS_NEEDS_REVIEW,
    STATUS_DUPLICATE_CANDIDATE,
    STATUS_READY_TO_POST,
    STATUS_POSTED,
    STATUS_DISMISSED,
}

EXTRACTED_FIELD_DEFAULTS: dict[str, Any] = {
    "merchant": "",
    "date": "",
    "amount": None,
    "subtotal": None,
    "tax": None,
    "tip": None,
    "currency_code": "",
    "category": "",
    "subcategory": "",
    "payment_method": "",
    "notes": "",
    "quantity": None,
    "unit": "",
    "line_items": [],
    "raw_candidates": {},
}

OCR_FIELD_DEFAULTS: dict[str, Any] = {
    "status": "pending",
    "engine": "",
    "language": "eng",
    "raw_text": "",
    "confidence": None,
    "processed_at": "",
    "error": "",
}

ATTACHMENT_FIELD_DEFAULTS: dict[str, Any] = {
    "id": "",
    "filename": "",
    "content_type": "",
    "size_bytes": 0,
    "checksum_sha256": "",
    "storage_path": "",
    "source": "upload",
    "uploaded_at": "",
}

REVIEW_FIELD_DEFAULTS: dict[str, Any] = {
    "decision": "",
    "reviewed_by": "",
    "reviewed_at": "",
    "review_notes": "",
    "expense_id": "",
}

STRING_EXTRACTED_FIELDS = {
    "merchant",
    "date",
    "currency_code",
    "category",
    "subcategory",
    "payment_method",
    "notes",
    "unit",
}


def ensure_receipt_inbox(data: dict[str, Any]) -> list[dict[str, Any]]:
    """Return a normalized receipt inbox list on the SpendSight data dict."""

    if not isinstance(data, dict):
        raise TypeError("SpendSight data must be a dict.")

    inbox = data.get(INBOX_KEY)
    if not isinstance(inbox, list):
        inbox = []
        data[INBOX_KEY] = inbox

    normalized = []
    changed = False
    for item in inbox:
        if not isinstance(item, dict):
            changed = True
            continue
        normalized_item = normalize_inbox_item(item)
        item.clear()
        item.update(normalized_item)
        normalized.append(item)
    if changed or len(normalized) != len(inbox):
        data[INBOX_KEY] = normalized
    return data[INBOX_KEY]


def new_attachment_metadata(
    filename: str,
    *,
    content_type: str = "",
    size_bytes: int | str | None = 0,
    checksum_sha256: str = "",
    storage_path: str = "",
    source: str = "upload",
    uploaded_at: str | datetime | None = None,
    attachment_id: str | None = None,
    **extra: Any,
) -> dict[str, Any]:
    """Create JSON-safe metadata for a stored receipt attachment."""

    metadata = copy.deepcopy(ATTACHMENT_FIELD_DEFAULTS)
    metadata.update(extra)
    metadata.update(
        {
            "id": attachment_id or str(uuid.uuid4()),
            "filename": _clean_string(filename),
            "content_type": _clean_string(content_type),
            "size_bytes": _coerce_int(size_bytes),
            "checksum_sha256": _clean_string(checksum_sha256),
            "storage_path": _clean_string(storage_path),
            "source": _clean_string(source) or "upload",
            "uploaded_at": _now_iso(uploaded_at),
        }
    )
    return metadata


def empty_extracted_fields(overrides: dict[str, Any] | None = None) -> dict[str, Any]:
    """Return OCR-ready extracted transaction placeholders."""

    return normalize_extracted_fields(overrides or {})


def empty_ocr_fields(overrides: dict[str, Any] | None = None) -> dict[str, Any]:
    """Return OCR processing placeholders for a receipt inbox item."""

    return normalize_ocr_fields(overrides or {})


def create_receipt_inbox_item(
    data: dict[str, Any],
    attachment: dict[str, Any],
    *,
    extracted: dict[str, Any] | None = None,
    ocr: dict[str, Any] | None = None,
    status: str | None = None,
    item_id: str | None = None,
    now: str | datetime | None = None,
    duplicate_options: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Append a receipt review item to ``data["receipt_inbox"]`` and return it."""

    inbox = ensure_receipt_inbox(data)
    created_at = _now_iso(now)
    item = {
        "id": item_id or str(uuid.uuid4()),
        "status": status or STATUS_NEEDS_REVIEW,
        "attachment": normalize_attachment_metadata(attachment),
        "ocr": normalize_ocr_fields(ocr or {}),
        "extracted": normalize_extracted_fields(extracted or {}),
        "duplicate_candidates": [],
        "review": copy.deepcopy(REVIEW_FIELD_DEFAULTS),
        "created_at": created_at,
        "updated_at": created_at,
    }
    item["duplicate_candidates"] = find_duplicate_expenses(
        item["extracted"],
        data.get("expenses", []),
        **(duplicate_options or {}),
    )
    if item["duplicate_candidates"] and status is None:
        item["status"] = STATUS_DUPLICATE_CANDIDATE
    item = normalize_inbox_item(item)
    inbox.append(item)
    return item


def normalize_inbox_item(item: dict[str, Any]) -> dict[str, Any]:
    """Normalize one inbox item while preserving unknown JSON-safe fields."""

    normalized = dict(item)
    normalized["id"] = _clean_string(normalized.get("id")) or str(uuid.uuid4())
    normalized["status"] = _normalize_status(normalized.get("status"))
    normalized["attachment"] = normalize_attachment_metadata(normalized.get("attachment") or {})
    normalized["ocr"] = normalize_ocr_fields(normalized.get("ocr") or {})
    normalized["extracted"] = normalize_extracted_fields(normalized.get("extracted") or {})
    normalized["review"] = normalize_review_fields(normalized.get("review") or {})
    normalized["duplicate_candidates"] = [
        dict(candidate)
        for candidate in normalized.get("duplicate_candidates", [])
        if isinstance(candidate, dict)
    ]
    normalized["created_at"] = _clean_string(normalized.get("created_at")) or _now_iso()
    normalized["updated_at"] = _clean_string(normalized.get("updated_at")) or normalized["created_at"]
    return normalized


def normalize_attachment_metadata(metadata: dict[str, Any]) -> dict[str, Any]:
    normalized = copy.deepcopy(ATTACHMENT_FIELD_DEFAULTS)
    if isinstance(metadata, dict):
        normalized.update(metadata)
    normalized["id"] = _clean_string(normalized.get("id")) or str(uuid.uuid4())
    normalized["filename"] = _clean_string(normalized.get("filename"))
    normalized["content_type"] = _clean_string(normalized.get("content_type"))
    normalized["size_bytes"] = _coerce_int(normalized.get("size_bytes"))
    normalized["checksum_sha256"] = _clean_string(normalized.get("checksum_sha256"))
    normalized["storage_path"] = _clean_string(normalized.get("storage_path"))
    normalized["source"] = _clean_string(normalized.get("source")) or "upload"
    normalized["uploaded_at"] = _clean_string(normalized.get("uploaded_at")) or _now_iso()
    return normalized


def normalize_ocr_fields(ocr: dict[str, Any]) -> dict[str, Any]:
    normalized = copy.deepcopy(OCR_FIELD_DEFAULTS)
    if isinstance(ocr, dict):
        normalized.update(ocr)
    for key in ("status", "engine", "language", "raw_text", "processed_at", "error"):
        normalized[key] = _clean_string(normalized.get(key))
    normalized["confidence"] = _coerce_float(normalized.get("confidence"))
    return normalized


def normalize_extracted_fields(extracted: dict[str, Any]) -> dict[str, Any]:
    normalized = copy.deepcopy(EXTRACTED_FIELD_DEFAULTS)
    if isinstance(extracted, dict):
        normalized.update(extracted)

    if normalized.get("amount") in (None, ""):
        normalized["amount"] = normalized.get("total") or normalized.get("total_amount")

    for key in ("amount", "subtotal", "tax", "tip", "quantity"):
        normalized[key] = _coerce_float(normalized.get(key))

    for key in STRING_EXTRACTED_FIELDS:
        normalized[key] = _clean_string(normalized.get(key))

    normalized["date"] = _to_iso_date(normalized.get("date"), fallback=normalized["date"])
    if not isinstance(normalized.get("line_items"), list):
        normalized["line_items"] = []
    else:
        normalized["line_items"] = [
            dict(item) for item in normalized["line_items"] if isinstance(item, dict)
        ]
    if not isinstance(normalized.get("raw_candidates"), dict):
        normalized["raw_candidates"] = {}
    return normalized


def normalize_review_fields(review: dict[str, Any]) -> dict[str, Any]:
    normalized = copy.deepcopy(REVIEW_FIELD_DEFAULTS)
    if isinstance(review, dict):
        normalized.update(review)
    for key in REVIEW_FIELD_DEFAULTS:
        normalized[key] = _clean_string(normalized.get(key))
    return normalized


def find_duplicate_expenses(
    receipt: dict[str, Any],
    expenses: list[dict[str, Any]],
    *,
    amount_tolerance: float = 1.0,
    date_window_days: int = 3,
    min_score: float = 0.72,
    limit: int = 5,
) -> list[dict[str, Any]]:
    """Find likely existing expenses for a receipt extraction.

    The scorer is intentionally transparent and deterministic: amount, date,
    merchant text, and payment method all contribute to the final score.
    """

    extracted = normalize_extracted_fields(receipt.get("extracted", receipt) if isinstance(receipt, dict) else {})
    receipt_amount = _coerce_float(extracted.get("amount"))
    if receipt_amount is None or receipt_amount <= 0:
        return []

    receipt_date = _parse_date(extracted.get("date"))
    receipt_text = _receipt_match_text(extracted)
    receipt_payment = _normalize_match_text(extracted.get("payment_method"))

    candidates: list[dict[str, Any]] = []
    for expense in expenses or []:
        if not isinstance(expense, dict):
            continue
        expense_amount = _coerce_float(expense.get("amount"))
        if expense_amount is None or expense_amount <= 0:
            continue

        amount_score = _amount_match_score(receipt_amount, expense_amount, amount_tolerance)
        if amount_score <= 0:
            continue

        expense_date = _parse_date(expense.get("date"))
        date_score = _date_match_score(receipt_date, expense_date, date_window_days)
        text_score = _text_match_score(receipt_text, _expense_match_text(expense))
        payment_score = 1.0 if receipt_payment and receipt_payment == _normalize_match_text(expense.get("payment_method")) else 0.0

        score = round(
            (amount_score * 0.45)
            + (date_score * 0.30)
            + (text_score * 0.20)
            + (payment_score * 0.05),
            3,
        )
        if score < min_score:
            continue

        matched_on = []
        if amount_score >= 0.7:
            matched_on.append("amount")
        if date_score >= 0.6:
            matched_on.append("date")
        if text_score >= 0.72:
            matched_on.append("merchant")
        if payment_score:
            matched_on.append("payment_method")

        candidates.append(
            {
                "expense_id": _clean_string(expense.get("id")),
                "score": score,
                "matched_on": matched_on,
                "amount_delta": round(abs(receipt_amount - expense_amount), 2),
                "date_delta_days": _date_delta_days(receipt_date, expense_date),
                "expense": _expense_preview(expense),
            }
        )

    candidates.sort(key=lambda candidate: candidate["score"], reverse=True)
    return candidates[:limit]


def refresh_duplicate_candidates(
    data: dict[str, Any],
    *,
    item_id: str | None = None,
    duplicate_options: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Recompute duplicate candidates for one item or the full inbox."""

    inbox = ensure_receipt_inbox(data)
    updated: list[dict[str, Any]] = []
    for item in inbox:
        if item_id and item.get("id") != item_id:
            continue
        item["duplicate_candidates"] = find_duplicate_expenses(
            item.get("extracted", {}),
            data.get("expenses", []),
            **(duplicate_options or {}),
        )
        if item["duplicate_candidates"] and item.get("status") not in {STATUS_POSTED, STATUS_DISMISSED}:
            item["status"] = STATUS_DUPLICATE_CANDIDATE
        item["updated_at"] = _now_iso()
        updated.append(item)
    return updated


def apply_review_update(
    data: dict[str, Any],
    item_id: str,
    *,
    extracted_updates: dict[str, Any] | None = None,
    status: str | None = None,
    review_notes: str | None = None,
    reviewed_by: str | None = None,
    now: str | datetime | None = None,
    duplicate_options: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Update extracted fields and review metadata for an inbox item."""

    item = get_inbox_item(data, item_id)
    current = dict(item.get("extracted", {}))
    current.update(extracted_updates or {})
    item["extracted"] = normalize_extracted_fields(current)

    if status:
        item["status"] = _normalize_status(status)

    review = normalize_review_fields(item.get("review", {}))
    if review_notes is not None:
        review["review_notes"] = _clean_string(review_notes)
    if reviewed_by is not None:
        review["reviewed_by"] = _clean_string(reviewed_by)
        review["reviewed_at"] = _now_iso(now)
    item["review"] = review

    item["duplicate_candidates"] = find_duplicate_expenses(
        item["extracted"],
        data.get("expenses", []),
        **(duplicate_options or {}),
    )
    if item["duplicate_candidates"] and item["status"] not in {STATUS_POSTED, STATUS_DISMISSED}:
        item["status"] = STATUS_DUPLICATE_CANDIDATE
    elif item["status"] == STATUS_DUPLICATE_CANDIDATE:
        item["status"] = STATUS_NEEDS_REVIEW

    item["updated_at"] = _now_iso(now)
    return item


def build_review_inbox(
    data: dict[str, Any],
    *,
    status: str | None = None,
) -> dict[str, Any]:
    """Build a template-friendly review inbox model from SpendSight data."""

    inbox = ensure_receipt_inbox(data)
    summary = {state: 0 for state in VALID_STATUSES}
    for item in inbox:
        summary[item.get("status", STATUS_NEEDS_REVIEW)] = summary.get(item.get("status"), 0) + 1

    selected_status = _clean_string(status)
    items = [
        item
        for item in inbox
        if not selected_status or item.get("status") == selected_status
    ]
    items.sort(key=lambda item: item.get("updated_at") or item.get("created_at") or "", reverse=True)
    return {
        "items": items,
        "summary": {
            "total": len(inbox),
            "open": sum(
                summary.get(state, 0)
                for state in (STATUS_NEEDS_REVIEW, STATUS_DUPLICATE_CANDIDATE, STATUS_READY_TO_POST)
            ),
            **summary,
        },
        "status_filter": selected_status,
    }


def build_expense_from_inbox_item(
    item: dict[str, Any],
    *,
    expense_id: str | None = None,
    now: str | datetime | None = None,
    default_category: str = "Other",
    default_payment_method: str = "Receipt",
) -> dict[str, Any]:
    """Convert a reviewed inbox item into the existing SpendSight expense shape."""

    normalized = normalize_inbox_item(item)
    extracted = normalized["extracted"]
    amount = _coerce_float(extracted.get("amount"))
    if amount is None or amount <= 0:
        raise ValueError("Receipt inbox item needs a positive amount before posting.")

    tx_date = _to_iso_date(extracted.get("date"), fallback=date.today().isoformat())
    category = extracted.get("category") or default_category
    merchant = extracted.get("merchant")
    subcategory = extracted.get("subcategory") or merchant or category
    payment_method = extracted.get("payment_method") or default_payment_method
    attachment = normalized["attachment"]

    note_parts = []
    if extracted.get("notes"):
        note_parts.append(extracted["notes"])
    if merchant and merchant.casefold() not in " ".join(note_parts).casefold():
        note_parts.append(f"Merchant: {merchant}")
    if attachment.get("filename"):
        note_parts.append(f"Receipt: {attachment['filename']}")

    expense = {
        "id": expense_id or str(uuid.uuid4()),
        "amount": round(amount, 2),
        "category": category,
        "subcategory": subcategory,
        "date": tx_date,
        "payment_method": payment_method,
        "notes": "; ".join(note_parts),
        "quantity": _coerce_float(extracted.get("quantity")),
        "unit": extracted.get("unit", ""),
        "created_at": _now_iso(now),
        "source": "receipt_inbox",
        "receipt_inbox_id": normalized["id"],
        "receipt_attachment_id": attachment.get("id", ""),
        "receipt_attachment": {
            "filename": attachment.get("filename", ""),
            "content_type": attachment.get("content_type", ""),
            "storage_path": attachment.get("storage_path", ""),
            "checksum_sha256": attachment.get("checksum_sha256", ""),
        },
    }
    return expense


def mark_item_posted(
    data: dict[str, Any],
    item_id: str,
    expense_id: str,
    *,
    reviewed_by: str = "",
    now: str | datetime | None = None,
) -> dict[str, Any]:
    """Mark an inbox item as posted after the parent app saves the expense."""

    item = get_inbox_item(data, item_id)
    posted_at = _now_iso(now)
    item["status"] = STATUS_POSTED
    item["updated_at"] = posted_at
    review = normalize_review_fields(item.get("review", {}))
    review.update(
        {
            "decision": "posted",
            "expense_id": _clean_string(expense_id),
            "reviewed_by": _clean_string(reviewed_by),
            "reviewed_at": posted_at,
        }
    )
    item["review"] = review
    return item


def dismiss_inbox_item(
    data: dict[str, Any],
    item_id: str,
    *,
    reviewed_by: str = "",
    review_notes: str = "",
    now: str | datetime | None = None,
) -> dict[str, Any]:
    """Mark an inbox item as dismissed without deleting its receipt metadata."""

    item = get_inbox_item(data, item_id)
    dismissed_at = _now_iso(now)
    item["status"] = STATUS_DISMISSED
    item["updated_at"] = dismissed_at
    review = normalize_review_fields(item.get("review", {}))
    review.update(
        {
            "decision": "dismissed",
            "reviewed_by": _clean_string(reviewed_by),
            "reviewed_at": dismissed_at,
            "review_notes": _clean_string(review_notes),
        }
    )
    item["review"] = review
    return item


def get_inbox_item(data: dict[str, Any], item_id: str) -> dict[str, Any]:
    inbox = ensure_receipt_inbox(data)
    for item in inbox:
        if item.get("id") == item_id:
            return item
    raise KeyError(f"Receipt inbox item not found: {item_id}")


def _expense_preview(expense: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": _clean_string(expense.get("id")),
        "date": _clean_string(expense.get("date")),
        "amount": _coerce_float(expense.get("amount")),
        "category": _clean_string(expense.get("category")),
        "subcategory": _clean_string(expense.get("subcategory")),
        "payment_method": _clean_string(expense.get("payment_method")),
        "notes": _clean_string(expense.get("notes")),
    }


def _receipt_match_text(extracted: dict[str, Any]) -> str:
    return _normalize_match_text(
        " ".join(
            str(extracted.get(field, ""))
            for field in ("merchant", "subcategory", "notes", "category")
        )
    )


def _expense_match_text(expense: dict[str, Any]) -> str:
    return _normalize_match_text(
        " ".join(
            str(expense.get(field, ""))
            for field in ("subcategory", "notes", "category")
        )
    )


def _amount_match_score(left: float, right: float, tolerance: float) -> float:
    delta = abs(left - right)
    if delta <= max(tolerance, 0):
        return 1.0
    relative_delta = delta / max(abs(left), abs(right), 1)
    if relative_delta <= 0.05:
        return 0.7
    return 0.0


def _date_match_score(left: date | None, right: date | None, window_days: int) -> float:
    if not left or not right:
        return 0.0
    delta_days = abs((left - right).days)
    if delta_days == 0:
        return 1.0
    if delta_days <= max(window_days, 0):
        return round(1 - (delta_days / (window_days + 1)), 3)
    return 0.0


def _text_match_score(left: str, right: str) -> float:
    if not left or not right:
        return 0.0
    if left in right or right in left:
        return 1.0
    return SequenceMatcher(None, left, right).ratio()


def _date_delta_days(left: date | None, right: date | None) -> int | None:
    if not left or not right:
        return None
    return abs((left - right).days)


def _normalize_status(status: Any) -> str:
    cleaned = _clean_string(status)
    return cleaned if cleaned in VALID_STATUSES else STATUS_NEEDS_REVIEW


def _normalize_match_text(value: Any) -> str:
    text = _clean_string(value).casefold()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _to_iso_date(value: Any, *, fallback: str = "") -> str:
    parsed = _parse_date(value)
    if parsed:
        return parsed.isoformat()
    return _clean_string(fallback)


def _parse_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raw = _clean_string(value)
    if not raw:
        return None
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%d %b %Y", "%d %B %Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    try:
        return date.fromisoformat(raw)
    except ValueError:
        return None


def _now_iso(value: str | datetime | None = None) -> str:
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, str) and value.strip():
        return value.strip()
    return datetime.now().isoformat()


def _clean_string(value: Any) -> str:
    return str(value or "").strip()


def _coerce_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return round(float(str(value).replace(",", "").strip()), 3)
    except (TypeError, ValueError):
        return None


def _coerce_int(value: Any) -> int:
    try:
        return max(0, int(float(str(value).replace(",", "").strip())))
    except (TypeError, ValueError):
        return 0
