"""Pure receipt OCR text extraction helpers for SpendSight feature drafts.

The functions in this module are deliberately Flask-free and side-effect free.
They parse pasted OCR text into JSON-safe fields that can live under
``data["receipts"][i]["extracted"]`` or seed a reviewed SpendSight expense.
"""

from __future__ import annotations

import re
import uuid
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher
from typing import Any, Iterable, Mapping


SPENDSIGHT_DEFAULT_CATEGORIES: dict[str, list[str]] = {
    "Groceries": [
        "Milk",
        "Dal",
        "Rice",
        "Sugar",
        "Atta",
        "Oil",
        "Soap",
        "Shampoo",
        "Vegetables",
        "Fruits",
        "Eggs",
        "Pulses",
        "Ghee",
        "Butter",
        "Tea",
    ],
    "Fast Food": [
        "McDonald's",
        "Wada Pav",
        "Pani Puri",
        "Samosa",
        "Pizza",
        "Chai",
        "Biryani",
        "Thali",
        "Dosa",
        "Idli",
        "Burger",
        "Noodles",
    ],
    "Fuel": ["Petrol", "Diesel", "CNG"],
    "Utilities": ["Electricity", "Water", "Gas Cylinder", "Internet", "Mobile Recharge", "DTH"],
    "Entertainment": ["Netflix", "Amazon Prime", "Hotstar", "Movies", "Games", "Events", "Spotify"],
    "Health": ["Medicine", "Doctor", "Gym", "Supplements", "Lab Test", "Dental"],
    "Transport": ["Ola/Uber", "Local Train", "Bus", "Auto", "Metro", "Parking"],
    "Shopping": ["Clothing", "Electronics", "Household", "Personal Care", "Accessories", "Books"],
    "EMI / Finance": ["Card EMI", "Loan EMI", "Insurance Premium", "SIP", "Rent"],
    "Other": ["Miscellaneous", "Gifts", "Donations", "Fees"],
}


CATEGORY_KEYWORDS: dict[str, list[str]] = {
    "Groceries": [
        "grocery",
        "groceries",
        "supermarket",
        "hypermarket",
        "dmart",
        "d mart",
        "big bazaar",
        "reliance fresh",
        "more supermarket",
        "fresh",
        "kirana",
        "vegetable",
        "fruit",
        "milk",
        "atta",
        "rice",
        "dal",
        "pulses",
        "eggs",
    ],
    "Fast Food": [
        "restaurant",
        "cafe",
        "coffee",
        "food",
        "pizza",
        "burger",
        "dominos",
        "mcdonald",
        "swiggy",
        "zomato",
        "biryani",
        "dosa",
        "idli",
        "chai",
        "thali",
    ],
    "Fuel": [
        "fuel",
        "petrol",
        "diesel",
        "cng",
        "indianoil",
        "indian oil",
        "bharat petroleum",
        "bpcl",
        "hpcl",
        "hindustan petroleum",
    ],
    "Utilities": [
        "electricity",
        "water bill",
        "gas cylinder",
        "broadband",
        "internet",
        "mobile recharge",
        "airtel",
        "jio",
        "vi recharge",
        "dth",
    ],
    "Entertainment": ["movie", "cinema", "pvr", "inox", "netflix", "spotify", "prime video", "hotstar", "game"],
    "Health": ["medical", "pharmacy", "chemist", "medicine", "hospital", "clinic", "doctor", "apollo", "medplus", "lab"],
    "Transport": ["ola", "uber", "taxi", "auto", "metro", "railway", "parking", "bus ticket", "toll"],
    "Shopping": [
        "store",
        "retail",
        "mall",
        "clothing",
        "apparel",
        "electronics",
        "books",
        "bookstore",
        "amazon",
        "flipkart",
        "myntra",
        "lifestyle",
        "shopper",
    ],
    "EMI / Finance": ["emi", "loan", "insurance premium", "sip", "finance charge", "bank fee"],
}


_MERCHANT_NOISE_RE = re.compile(
    r"\b("
    r"tax\s*invoice|invoice|receipt|bill\s*no|cash\s*memo|cardholder\s*copy|merchant\s*copy|"
    r"customer\s*copy|duplicate|gstin|gst\s*no|fssai|cin|pan\s*no|vat|tin|"
    r"date|time|terminal|tid|mid|batch|approval|auth|rrn|stan|sale|void|"
    r"total|subtotal|sub\s*total|amount|balance|change|cash|paid|payment|"
    r"phone|mobile|tel|address|pin\s*code|thank\s*you|visit\s*again"
    r")\b",
    re.IGNORECASE,
)
_ADDRESS_HINT_RE = re.compile(r"\b(road|rd|street|st|lane|nagar|sector|phase|floor|shop|opp|near|city|state)\b", re.I)
_DATE_LABEL_RE = re.compile(r"\b(date|dt|bill\s*date|txn\s*date|trans\s*date|transaction\s*date)\b", re.I)
_NEGATIVE_DATE_LABEL_RE = re.compile(r"\b(expiry|valid|member|card|batch|invoice|bill\s*no|order\s*no)\b", re.I)
_TOTAL_LABEL_RE = re.compile(
    r"\b("
    r"grand\s*total|net\s*(amount|payable)|amount\s*(due|paid|payable)|total\s*(amount|due|paid|payable)?|"
    r"sale\s*amount|purchase\s*amount|card\s*sale|paid\s*amount"
    r")\b",
    re.I,
)
_LOW_PRIORITY_AMOUNT_RE = re.compile(
    r"\b(sub\s*total|subtotal|taxable|cgst|sgst|igst|gst|vat|service\s*tax|discount|round\s*off|"
    r"change|balance|tender|cash\s*received|cash\s*tendered|qty|quantity|items?|mrp|rate|unit\s*price)\b",
    re.I,
)
_PAYMENT_SKIP_AMOUNT_RE = re.compile(r"\b(card\s*(no|number)?|approval|auth|batch|tid|mid|rrn|stan|trace)\b", re.I)
_TIME_RE = re.compile(r"(?<!\d)\d{1,2}:\d{2}(?::\d{2})?(?!\d)")
_NUMBER_RE = re.compile(
    r"(?P<currency>rs\.?|inr|usd|eur|gbp|aud|cad|sgd|s\$|ca\$|r\$|\$|\u20b9)?"
    r"\s*"
    r"(?P<amount>[+-]?(?:\d{1,3}(?:,\d{2,3})+|\d+)(?:\.\d{1,2})?)",
    re.I,
)
_NUMERIC_DATE_RE = re.compile(r"(?<!\d)(\d{1,4})[\/.\-](\d{1,2})[\/.\-](\d{1,4})(?!\d)")
_SHORT_NUMERIC_DATE_RE = re.compile(r"(?<!\d)(\d{1,2})[\/.\-](\d{1,2})(?!\d)")
_MONTH_NAME_RE = re.compile(
    r"(?<!\w)(\d{1,2})(?:st|nd|rd|th)?[\s\-/,]+"
    r"(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|"
    r"aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
    r"[\s\-/,]+(\d{2,4})(?!\w)",
    re.I,
)
_MONTH_FIRST_RE = re.compile(
    r"(?<!\w)"
    r"(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|"
    r"aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
    r"[\s\-/,]+(\d{1,2})(?:st|nd|rd|th)?(?:,)?[\s\-/,]+(\d{2,4})(?!\w)",
    re.I,
)

_MONTHS = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}


def parse_receipt_ocr_text(
    ocr_text: str,
    data: Mapping[str, Any] | None = None,
    *,
    date_order: str = "DMY",
    ref_date: date | datetime | str | None = None,
    default_currency_code: str | None = None,
    duplicate_date_window_days: int = 0,
) -> dict[str, Any]:
    """Parse pasted receipt OCR text into SpendSight-compatible extracted fields.

    ``data`` is optional. When supplied, custom categories are used for category
    suggestions and existing ``data["expenses"]`` are scored as duplicate
    matches. The input dict is never mutated.
    """

    raw_text = str(ocr_text or "").strip()
    lines = _normalized_lines(raw_text)
    reference_date = _as_date(ref_date) or date.today()
    currency_code = detect_currency_code(raw_text, data=data, default=default_currency_code)

    merchant_candidates = _find_merchant_candidates(lines)
    date_candidates = _find_date_candidates(lines, date_order=date_order, ref_date=reference_date)
    amount_candidates = _find_amount_candidates(lines)

    merchant = merchant_candidates[0]["value"] if merchant_candidates else ""
    iso_date = date_candidates[0]["value"] if date_candidates else ""
    amount = amount_candidates[0]["value"] if amount_candidates else None
    payment_method = detect_payment_method(raw_text)

    categories = build_category_map(data)
    category_suggestion = suggest_category(merchant, raw_text, categories=categories)
    category = category_suggestion["category"]
    subcategory = merchant or category

    duplicate_keys = build_duplicate_key_candidates(
        merchant_candidates=merchant_candidates,
        date_candidates=date_candidates,
        amount_candidates=amount_candidates,
        date_window_days=duplicate_date_window_days,
    )

    confidence = _build_confidence(
        merchant_candidates=merchant_candidates,
        date_candidates=date_candidates,
        amount_candidates=amount_candidates,
        category_suggestion=category_suggestion,
    )
    signals = _build_confidence_signals(
        lines=lines,
        merchant_candidates=merchant_candidates,
        date_candidates=date_candidates,
        amount_candidates=amount_candidates,
        duplicate_keys=duplicate_keys,
        payment_method=payment_method,
    )

    result = {
        "merchant": merchant,
        "date": iso_date,
        "amount": amount,
        "category": category,
        "subcategory": subcategory,
        "payment_method": payment_method,
        "currency_code": currency_code,
        "notes": f"OCR receipt: {merchant}" if merchant else "OCR receipt",
        "ocr_text": raw_text,
        "category_suggestion": category_suggestion,
        "confidence": confidence,
        "confidence_signals": signals,
        "duplicate_key_candidates": duplicate_keys,
        "raw_candidates": {
            "merchant": merchant_candidates,
            "date": date_candidates,
            "amount": amount_candidates,
        },
    }

    if data is not None:
        result["duplicate_matches"] = find_duplicate_expense_matches(result, data.get("expenses", []))

    return result


def build_spendsight_receipt_payload(
    ocr_text: str,
    data: Mapping[str, Any] | None = None,
    *,
    receipt_id: str | None = None,
    original_filename: str = "",
    stored_filename: str = "",
    content_type: str = "",
    uploaded_at: str | datetime | None = None,
    status: str = "needs_review",
    **parse_options: Any,
) -> dict[str, Any]:
    """Return a JSON-safe receipt row compatible with ``data["receipts"]``."""

    return {
        "id": receipt_id or str(uuid.uuid4()),
        "original_filename": _clean_string(original_filename),
        "stored_filename": _clean_string(stored_filename),
        "content_type": _clean_string(content_type),
        "uploaded_at": _iso_timestamp(uploaded_at),
        "status": _clean_string(status) or "needs_review",
        "extracted": parse_receipt_ocr_text(ocr_text, data=data, **parse_options),
        "expense_id": "",
    }


def build_expense_candidate(
    extracted: Mapping[str, Any],
    *,
    expense_id: str | None = None,
    now: str | datetime | None = None,
    default_category: str = "Other",
    default_payment_method: str = "Cash",
) -> dict[str, Any]:
    """Build an expense-shaped draft from extracted receipt fields."""

    amount = _coerce_money(extracted.get("amount"))
    merchant = _clean_string(extracted.get("merchant"))
    category = _clean_string(extracted.get("category")) or default_category
    payment_method = _clean_string(extracted.get("payment_method")) or default_payment_method
    notes = _clean_string(extracted.get("notes")) or (f"OCR receipt: {merchant}" if merchant else "OCR receipt")

    return {
        "id": expense_id or str(uuid.uuid4()),
        "amount": round(amount or 0.0, 2),
        "category": category,
        "subcategory": _clean_string(extracted.get("subcategory")) or merchant or category,
        "date": _clean_string(extracted.get("date")) or date.today().isoformat(),
        "payment_method": payment_method,
        "notes": notes,
        "quantity": None,
        "unit": "",
        "source": "receipt_ocr",
        "review_status": "needs_review",
        "created_at": _iso_timestamp(now),
    }


def build_duplicate_key_candidates(
    *,
    merchant_candidates: Iterable[Mapping[str, Any]] = (),
    date_candidates: Iterable[Mapping[str, Any]] = (),
    amount_candidates: Iterable[Mapping[str, Any]] = (),
    date_window_days: int = 0,
) -> list[dict[str, Any]]:
    """Build JSON-safe probes mirroring SpendSight's duplicate key shape.

    The ``key`` field is a list equivalent to the app's tuple:
    ``[date, rounded_amount, merchant.casefold()]``.
    """

    merchants = [
        _clean_string(candidate.get("value"))
        for candidate in sorted(merchant_candidates, key=lambda item: item.get("score", 0), reverse=True)
        if _clean_string(candidate.get("value"))
    ][:3]
    dates = [
        _clean_string(candidate.get("value"))
        for candidate in sorted(date_candidates, key=lambda item: item.get("score", 0), reverse=True)
        if _clean_string(candidate.get("value"))
    ][:3]
    amounts = [
        _coerce_money(candidate.get("value"))
        for candidate in sorted(amount_candidates, key=lambda item: item.get("score", 0), reverse=True)
        if _coerce_money(candidate.get("value")) is not None
    ][:3]

    expanded_dates: list[tuple[str, str]] = []
    for value in dates[:1]:
        parsed = _parse_iso_date(value)
        if not parsed:
            continue
        expanded_dates.append((value, "exact"))
        for offset in range(1, max(0, int(date_window_days)) + 1):
            expanded_dates.append(((parsed - timedelta(days=offset)).isoformat(), f"-{offset}d"))
            expanded_dates.append(((parsed + timedelta(days=offset)).isoformat(), f"+{offset}d"))
    for value in dates[1:]:
        if value not in {item[0] for item in expanded_dates}:
            expanded_dates.append((value, "alternate"))

    candidates: list[dict[str, Any]] = []
    seen: set[tuple[str, float, str]] = set()
    for merchant_index, merchant in enumerate(merchants):
        merchant_key = _merchant_duplicate_key(merchant)
        if not merchant_key:
            continue
        for date_value, date_basis in expanded_dates:
            for amount in amounts:
                if amount is None:
                    continue
                rounded_amount = round(amount, 2)
                key = (date_value, rounded_amount, merchant_key)
                if key in seen:
                    continue
                seen.add(key)
                score = max(0.1, 0.98 - (merchant_index * 0.08))
                if date_basis != "exact":
                    score -= 0.08
                candidates.append(
                    {
                        "key": [date_value, rounded_amount, merchant_key],
                        "date": date_value,
                        "amount": rounded_amount,
                        "merchant": merchant,
                        "score": round(score, 3),
                        "basis": ["date", "amount", "merchant", date_basis],
                        "expense_probe": {
                            "date": date_value,
                            "amount": rounded_amount,
                            "subcategory": merchant,
                            "notes": merchant,
                        },
                    }
                )
    candidates.sort(key=lambda item: item["score"], reverse=True)
    return candidates[:12]


def find_duplicate_expense_matches(
    extracted: Mapping[str, Any],
    expenses: Iterable[Mapping[str, Any]],
    *,
    amount_tolerance: float = 1.0,
    date_window_days: int = 3,
    merchant_threshold: float = 0.72,
    min_score: float = 0.70,
    limit: int = 5,
) -> list[dict[str, Any]]:
    """Score existing SpendSight expenses against parsed receipt fields."""

    amount = _coerce_money(extracted.get("amount"))
    merchant = _clean_string(extracted.get("merchant") or extracted.get("subcategory"))
    receipt_date = _parse_iso_date(extracted.get("date"))
    if amount is None or amount <= 0:
        return []

    matches: list[dict[str, Any]] = []
    merchant_key = _normalize_match_text(merchant)
    for expense in expenses or []:
        if not isinstance(expense, Mapping):
            continue
        expense_amount = _coerce_money(expense.get("amount"))
        if expense_amount is None or expense_amount <= 0:
            continue

        amount_delta = abs(amount - expense_amount)
        if amount_delta <= max(0.0, amount_tolerance):
            amount_score = 1.0
        elif amount_delta / max(amount, expense_amount, 1.0) <= 0.05:
            amount_score = 0.72
        else:
            amount_score = 0.0
        if amount_score <= 0:
            continue

        expense_date = _parse_iso_date(expense.get("date"))
        date_delta = abs((receipt_date - expense_date).days) if receipt_date and expense_date else None
        if date_delta == 0:
            date_score = 1.0
        elif date_delta is not None and date_delta <= max(0, date_window_days):
            date_score = 1.0 - (date_delta / (date_window_days + 1))
        else:
            date_score = 0.0

        expense_text = _normalize_match_text(
            " ".join(str(expense.get(field, "")) for field in ("subcategory", "notes", "category"))
        )
        merchant_score = _text_score(merchant_key, expense_text)
        if merchant_score < merchant_threshold and date_score <= 0:
            continue

        score = round((amount_score * 0.45) + (date_score * 0.35) + (merchant_score * 0.20), 3)
        if score < min_score:
            continue

        matched_on = []
        if amount_score >= 0.72:
            matched_on.append("amount")
        if date_score > 0:
            matched_on.append("date")
        if merchant_score >= merchant_threshold:
            matched_on.append("merchant")

        matches.append(
            {
                "expense_id": _clean_string(expense.get("id")),
                "score": score,
                "matched_on": matched_on,
                "amount_delta": round(amount_delta, 2),
                "date_delta_days": date_delta,
                "expense": {
                    "id": _clean_string(expense.get("id")),
                    "date": _clean_string(expense.get("date")),
                    "amount": expense_amount,
                    "category": _clean_string(expense.get("category")),
                    "subcategory": _clean_string(expense.get("subcategory")),
                    "payment_method": _clean_string(expense.get("payment_method")),
                    "notes": _clean_string(expense.get("notes")),
                },
            }
        )

    matches.sort(key=lambda item: item["score"], reverse=True)
    return matches[:limit]


def suggest_category(
    merchant: str,
    ocr_text: str = "",
    *,
    categories: Mapping[str, Iterable[str]] | None = None,
) -> dict[str, Any]:
    """Suggest a SpendSight category from merchant and OCR text keywords."""

    category_map = {key: list(value) for key, value in (categories or SPENDSIGHT_DEFAULT_CATEGORIES).items()}
    haystack = _normalize_match_text(f"{merchant} {ocr_text}")
    scores: dict[str, float] = {category: 0.0 for category in category_map}
    reasons: dict[str, list[str]] = {category: [] for category in category_map}

    for category, keywords in CATEGORY_KEYWORDS.items():
        if category not in scores:
            continue
        for keyword in keywords:
            if _phrase_in_text(keyword, haystack):
                scores[category] += 2.0 if keyword in _normalize_match_text(merchant) else 1.0
                reasons[category].append(keyword)

    for category, subcategories in category_map.items():
        category_key = _normalize_match_text(category)
        if category_key and _phrase_in_text(category_key, haystack):
            scores[category] += 1.2
            reasons[category].append(category)
        for subcategory in subcategories:
            sub_key = _normalize_match_text(subcategory)
            if sub_key and _phrase_in_text(sub_key, haystack):
                scores[category] += 1.6
                reasons[category].append(str(subcategory))

    best_category = max(scores, key=lambda item: scores[item]) if scores else "Other"
    best_score = scores.get(best_category, 0.0)
    if best_score <= 0:
        return {
            "category": "Other",
            "confidence": 0.25,
            "matched_keywords": [],
            "reason": "No category keyword matched.",
        }

    confidence = min(0.95, 0.35 + (best_score * 0.12))
    return {
        "category": best_category,
        "confidence": round(confidence, 3),
        "matched_keywords": sorted(set(reasons[best_category]), key=str.casefold)[:8],
        "reason": f"Matched receipt text to {best_category}.",
    }


def build_category_map(data: Mapping[str, Any] | None = None) -> dict[str, list[str]]:
    """Merge SpendSight defaults with optional ``data["custom_categories"]``."""

    categories = {category: list(subcategories) for category, subcategories in SPENDSIGHT_DEFAULT_CATEGORIES.items()}
    if not isinstance(data, Mapping):
        return categories

    custom = data.get("custom_categories", {})
    if not isinstance(custom, Mapping):
        return categories
    for category, subcategories in custom.items():
        clean_category = _clean_string(category)
        if not clean_category:
            continue
        bucket = categories.setdefault(clean_category, [])
        if isinstance(subcategories, Iterable) and not isinstance(subcategories, (str, bytes)):
            for subcategory in subcategories:
                clean_subcategory = _clean_string(subcategory)
                if clean_subcategory and clean_subcategory not in bucket:
                    bucket.append(clean_subcategory)
    return categories


def detect_currency_code(
    ocr_text: str,
    *,
    data: Mapping[str, Any] | None = None,
    default: str | None = None,
) -> str:
    text = str(ocr_text or "").casefold()
    if re.search(r"(\binr\b|\brs\.?\b|\u20b9)", text):
        return "INR"
    if re.search(r"(\busd\b|\$)", text):
        return "USD"
    if re.search(r"\bgbp\b", text):
        return "GBP"
    if re.search(r"\beur\b", text):
        return "EUR"
    if default:
        return _clean_string(default).upper()
    if isinstance(data, Mapping) and data.get("currency_code"):
        return _clean_string(data.get("currency_code")).upper()
    return "INR"


def detect_payment_method(ocr_text: str) -> str:
    text = str(ocr_text or "").casefold()
    if re.search(r"\b(upi|bhim|vpa|paytm|phonepe|gpay|google\s*pay)\b", text):
        return "UPI"
    if re.search(r"\b(visa|master\s*card|mastercard|amex|rupay|cardholder|card\s*no|auth\s*code|tid|mid)\b", text):
        return "Card"
    if re.search(r"\b(cash|cashier|cash\s*tendered)\b", text):
        return "Cash"
    return ""


def _find_merchant_candidates(lines: list[str]) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    search_lines = lines[:14] if len(lines) > 14 else lines
    for index, line in enumerate(search_lines):
        value = _clean_merchant_line(line)
        if not value:
            continue
        alpha_count = len(re.findall(r"[A-Za-z]", value))
        if alpha_count < 3:
            continue
        if _MERCHANT_NOISE_RE.search(value):
            continue
        digit_count = len(re.findall(r"\d", value))
        if digit_count > alpha_count:
            continue

        score = 0.55
        score += max(0.0, 0.25 - (index * 0.025))
        if value.isupper():
            score += 0.04
        if re.search(r"\b(ltd|limited|pvt|private|store|mart|cafe|foods?|restaurant|pharmacy|fuel|book)\b", value, re.I):
            score += 0.10
        if _ADDRESS_HINT_RE.search(value):
            score -= 0.18
        if digit_count:
            score -= min(0.2, digit_count * 0.03)

        candidates.append(
            {
                "value": value,
                "score": round(min(max(score, 0.05), 0.98), 3),
                "line_index": index,
                "line": line,
            }
        )
    candidates.sort(key=lambda item: item["score"], reverse=True)
    return candidates[:5]


def _find_date_candidates(lines: list[str], *, date_order: str, ref_date: date) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    seen: set[str] = set()

    for index, line in enumerate(lines):
        spans: list[tuple[int, int]] = []
        for match in _NUMERIC_DATE_RE.finditer(line):
            parsed = _parse_numeric_date(match.groups(), date_order=date_order)
            if parsed:
                spans.append(match.span())
                _append_date_candidate(candidates, seen, parsed, match.group(0), line, index)

        for match in _MONTH_NAME_RE.finditer(line):
            parsed = _safe_date(_expand_year(match.group(3)), _MONTHS[match.group(2).casefold()[:3]], int(match.group(1)))
            if parsed:
                spans.append(match.span())
                _append_date_candidate(candidates, seen, parsed, match.group(0), line, index)

        for match in _MONTH_FIRST_RE.finditer(line):
            parsed = _safe_date(_expand_year(match.group(3)), _MONTHS[match.group(1).casefold()[:3]], int(match.group(2)))
            if parsed:
                spans.append(match.span())
                _append_date_candidate(candidates, seen, parsed, match.group(0), line, index)

        for match in _SHORT_NUMERIC_DATE_RE.finditer(line):
            if any(_span_overlaps(match.span(), span) for span in spans):
                continue
            if not _DATE_LABEL_RE.search(line):
                continue
            parsed = _parse_short_date(match.groups(), date_order=date_order, ref_date=ref_date)
            if parsed:
                _append_date_candidate(candidates, seen, parsed, match.group(0), line, index)

    candidates.sort(key=lambda item: item["score"], reverse=True)
    return candidates[:5]


def _append_date_candidate(
    candidates: list[dict[str, Any]],
    seen: set[str],
    parsed: date,
    raw: str,
    line: str,
    index: int,
) -> None:
    value = parsed.isoformat()
    if value in seen:
        return
    score = 0.62
    if _DATE_LABEL_RE.search(line):
        score += 0.24
    if _NEGATIVE_DATE_LABEL_RE.search(line):
        score -= 0.25
    score += max(0.0, 0.10 - (index * 0.008))
    seen.add(value)
    candidates.append(
        {
            "value": value,
            "raw": raw,
            "score": round(min(max(score, 0.05), 0.98), 3),
            "line_index": index,
            "line": line,
        }
    )


def _find_amount_candidates(lines: list[str]) -> list[dict[str, Any]]:
    all_candidates: list[dict[str, Any]] = []
    for index, line in enumerate(lines):
        skip_spans = [match.span() for match in _NUMERIC_DATE_RE.finditer(line)]
        if _DATE_LABEL_RE.search(line) and not re.search(r"\b(amount|total|paid|payable|due)\b", line, re.I):
            skip_spans.extend(match.span() for match in _SHORT_NUMERIC_DATE_RE.finditer(line))
        skip_spans.extend(match.span() for match in _TIME_RE.finditer(line))

        for match in _NUMBER_RE.finditer(line):
            if any(_span_overlaps(match.span("amount"), span) for span in skip_spans):
                continue
            value = _coerce_money(match.group("amount"))
            if value is None or value <= 0:
                continue
            raw = match.group(0).strip()
            if not raw:
                continue
            if _should_skip_amount_line(line, raw):
                continue

            score = 0.30
            if _TOTAL_LABEL_RE.search(line):
                score += 0.44
            if re.search(r"\b(amount|paid|payable|due)\b", line, re.I):
                score += 0.18
            if match.group("currency"):
                score += 0.16
            if "." in match.group("amount"):
                score += 0.06
            if index >= max(0, len(lines) - 8):
                score += 0.06
            if _LOW_PRIORITY_AMOUNT_RE.search(line):
                score -= 0.30
            if _PAYMENT_SKIP_AMOUNT_RE.search(line) and not _TOTAL_LABEL_RE.search(line):
                score -= 0.30

            all_candidates.append(
                {
                    "value": round(value, 2),
                    "raw": raw,
                    "score": round(min(max(score, 0.05), 0.98), 3),
                    "line_index": index,
                    "line": line,
                }
            )

    if not all_candidates:
        return []

    max_amount = max(candidate["value"] for candidate in all_candidates)
    for candidate in all_candidates:
        if candidate["value"] == max_amount:
            candidate["score"] = round(min(0.98, candidate["score"] + 0.10), 3)
    all_candidates.sort(key=lambda item: (item["score"], item["value"]), reverse=True)

    deduped: list[dict[str, Any]] = []
    seen: set[tuple[float, int]] = set()
    for candidate in all_candidates:
        key = (candidate["value"], candidate["line_index"])
        if key in seen:
            continue
        seen.add(key)
        deduped.append(candidate)
    return deduped[:8]


def _should_skip_amount_line(line: str, raw_amount: str) -> bool:
    if _PAYMENT_SKIP_AMOUNT_RE.search(line) and not (_TOTAL_LABEL_RE.search(line) or re.search(r"\bamount\b", line, re.I)):
        return True
    if re.search(r"\b(gstin|gst\s*no|fssai|phone|mobile|tel|invoice\s*no|bill\s*no|order\s*no)\b", line, re.I):
        return True
    if _ADDRESS_HINT_RE.search(line) and not re.search(r"\b(amount|total|paid|payable|due)\b", line, re.I):
        return True
    if re.search(r"\b(items?|qty|quantity)\b", line, re.I) and not re.search(r"\b(amount|total|payable|due)\b", line, re.I):
        return True
    digits = re.sub(r"\D", "", raw_amount)
    if len(digits) >= 10 and "." not in raw_amount and "," not in raw_amount:
        return True
    return False


def _build_confidence(
    *,
    merchant_candidates: list[dict[str, Any]],
    date_candidates: list[dict[str, Any]],
    amount_candidates: list[dict[str, Any]],
    category_suggestion: Mapping[str, Any],
) -> dict[str, float]:
    merchant = merchant_candidates[0]["score"] if merchant_candidates else 0.0
    tx_date = date_candidates[0]["score"] if date_candidates else 0.0
    amount = amount_candidates[0]["score"] if amount_candidates else 0.0
    category = float(category_suggestion.get("confidence", 0.0) or 0.0)
    overall = (merchant * 0.25) + (tx_date * 0.25) + (amount * 0.35) + (category * 0.15)
    return {
        "overall": round(overall, 3),
        "merchant": round(merchant, 3),
        "date": round(tx_date, 3),
        "amount": round(amount, 3),
        "category": round(category, 3),
    }


def _build_confidence_signals(
    *,
    lines: list[str],
    merchant_candidates: list[dict[str, Any]],
    date_candidates: list[dict[str, Any]],
    amount_candidates: list[dict[str, Any]],
    duplicate_keys: list[dict[str, Any]],
    payment_method: str,
) -> dict[str, Any]:
    return {
        "line_count": len(lines),
        "merchant_candidate_count": len(merchant_candidates),
        "date_candidate_count": len(date_candidates),
        "amount_candidate_count": len(amount_candidates),
        "has_labeled_date": bool(date_candidates and _DATE_LABEL_RE.search(date_candidates[0].get("line", ""))),
        "has_labeled_total": bool(amount_candidates and _TOTAL_LABEL_RE.search(amount_candidates[0].get("line", ""))),
        "has_currency_marker": bool(amount_candidates and re.search(r"(rs\.?|inr|usd|\$|\u20b9)", amount_candidates[0].get("raw", ""), re.I)),
        "payment_method_detected": payment_method,
        "duplicate_key_candidate_count": len(duplicate_keys),
        "missing_fields": [
            field
            for field, present in {
                "merchant": bool(merchant_candidates),
                "date": bool(date_candidates),
                "amount": bool(amount_candidates),
            }.items()
            if not present
        ],
    }


def _normalized_lines(ocr_text: str) -> list[str]:
    lines = []
    for raw_line in str(ocr_text or "").replace("\r", "\n").split("\n"):
        line = re.sub(r"\s+", " ", raw_line).strip(" \t|")
        if line:
            lines.append(line)
    return lines


def _clean_merchant_line(line: str) -> str:
    value = re.sub(r"^[^A-Za-z0-9]+|[^A-Za-z0-9)&.' -]+$", "", line.strip())
    value = re.sub(r"\s+", " ", value)
    value = re.sub(r"^(m/s|m\/s|merchant name)\s*[:\-]?\s*", "", value, flags=re.I)
    return value.strip(" -:")


def _parse_numeric_date(parts: tuple[str, str, str], *, date_order: str) -> date | None:
    first, second, third = (int(part) for part in parts)
    order = date_order.upper()

    if len(parts[0]) == 4:
        return _safe_date(first, second, third)
    if len(parts[2]) == 4 or len(parts[2]) == 2:
        year = _expand_year(parts[2])
        if first > 12 and second <= 12:
            return _safe_date(year, second, first)
        if second > 12 and first <= 12:
            return _safe_date(year, first, second)
        if order == "MDY":
            return _safe_date(year, first, second)
        return _safe_date(year, second, first)
    return None


def _parse_short_date(parts: tuple[str, str], *, date_order: str, ref_date: date) -> date | None:
    first, second = (int(part) for part in parts)
    if first > 12 and second <= 12:
        day, month = first, second
    elif second > 12 and first <= 12:
        month, day = first, second
    elif date_order.upper() == "MDY":
        month, day = first, second
    else:
        day, month = first, second
    return _safe_date(ref_date.year, month, day)


def _expand_year(value: str) -> int:
    year = int(value)
    if year < 100:
        return 2000 + year if year < 80 else 1900 + year
    return year


def _safe_date(year: int, month: int, day: int) -> date | None:
    try:
        return date(year, month, day)
    except ValueError:
        return None


def _parse_iso_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raw = _clean_string(value)
    if not raw:
        return None
    try:
        return date.fromisoformat(raw[:10])
    except ValueError:
        return None


def _as_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str) and value.strip():
        try:
            return date.fromisoformat(value.strip()[:10])
        except ValueError:
            return None
    return None


def _coerce_money(value: Any) -> float | None:
    if value in (None, ""):
        return None
    raw = str(value).strip()
    if not raw:
        return None
    raw = raw.replace(",", "")
    raw = re.sub(r"^(rs\.?|inr|usd|eur|gbp|aud|cad|sgd|s\$|ca\$|r\$|\$|\u20b9)\s*", "", raw, flags=re.I)
    try:
        return round(abs(float(raw)), 2)
    except (TypeError, ValueError):
        return None


def _span_overlaps(left: tuple[int, int], right: tuple[int, int]) -> bool:
    return left[0] < right[1] and right[0] < left[1]


def _merchant_duplicate_key(value: Any) -> str:
    return re.sub(r"\s+", " ", _clean_string(value)).casefold()


def _normalize_match_text(value: Any) -> str:
    text = _clean_string(value).casefold()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _phrase_in_text(phrase: str, haystack: str) -> bool:
    normalized = _normalize_match_text(phrase)
    if not normalized:
        return False
    return re.search(rf"(?<![a-z0-9]){re.escape(normalized)}(?![a-z0-9])", haystack) is not None


def _text_score(left: str, right: str) -> float:
    if not left or not right:
        return 0.0
    if left in right or right in left:
        return 1.0
    return SequenceMatcher(None, left, right).ratio()


def _clean_string(value: Any) -> str:
    return str(value or "").strip()


def _iso_timestamp(value: str | datetime | None = None) -> str:
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, str) and value.strip():
        return value.strip()
    return datetime.now().isoformat()
