"""Draft OCR extraction helpers for SpendSight receipts."""

from .helpers import (
    SPENDSIGHT_DEFAULT_CATEGORIES,
    build_category_map,
    build_duplicate_key_candidates,
    build_expense_candidate,
    build_spendsight_receipt_payload,
    detect_currency_code,
    detect_payment_method,
    find_duplicate_expense_matches,
    parse_receipt_ocr_text,
    suggest_category,
)

__all__ = [
    "SPENDSIGHT_DEFAULT_CATEGORIES",
    "build_category_map",
    "build_duplicate_key_candidates",
    "build_expense_candidate",
    "build_spendsight_receipt_payload",
    "detect_currency_code",
    "detect_payment_method",
    "find_duplicate_expense_matches",
    "parse_receipt_ocr_text",
    "suggest_category",
]
