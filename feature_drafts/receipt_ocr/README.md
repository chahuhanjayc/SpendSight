# Receipt OCR Extraction Draft

This draft is isolated from the live Flask app. It parses pasted OCR text with pure Python heuristics and returns JSON-safe fields that fit the current SpendSight receipt and expense data shapes.

## Helper API

```python
from feature_drafts.receipt_ocr import (
    build_expense_candidate,
    build_spendsight_receipt_payload,
    parse_receipt_ocr_text,
)
```

`parse_receipt_ocr_text(ocr_text, data=None)` returns:

- `merchant`, `date` as ISO `YYYY-MM-DD`, `amount`
- `category`, `subcategory`, `category_suggestion`
- `payment_method`, `currency_code`, `notes`, `ocr_text`
- `confidence` and `confidence_signals`
- `raw_candidates` for merchant/date/amount review UI
- `duplicate_key_candidates`, where each `key` mirrors SpendSight's duplicate tuple as a JSON list: `[date, rounded_amount, merchant.casefold()]`
- `duplicate_matches` when a SpendSight data dict is supplied

`build_spendsight_receipt_payload(...)` wraps the extraction in the current `data["receipts"]` row shape:

```python
{
    "id": "...",
    "original_filename": "...",
    "stored_filename": "...",
    "content_type": "...",
    "uploaded_at": "...",
    "status": "needs_review",
    "extracted": parse_receipt_ocr_text(...),
    "expense_id": "",
}
```

`build_expense_candidate(extracted)` builds an expense-shaped draft with `source="receipt_ocr"` and `review_status="needs_review"`.

## Flask Wiring Instructions

When promoting this draft, import the helper in `app.py`:

```python
from feature_drafts.receipt_ocr import (
    build_expense_candidate,
    build_spendsight_receipt_payload,
    parse_receipt_ocr_text,
)
```

For the existing `/receipts` upload flow, after saving the file and reading pasted OCR text from a form field, replace the manual `extracted = {...}` assembly with:

```python
extracted = parse_receipt_ocr_text(
    request.form.get("ocr_text", ""),
    data,
    default_currency_code=data.get("currency_code", "INR"),
)
```

Then keep the current receipt row shape and assign `extracted` to `receipt["extracted"]`.

For a cleaner route body, use:

```python
receipt = build_spendsight_receipt_payload(
    request.form.get("ocr_text", ""),
    data,
    receipt_id=receipt_id,
    original_filename=filename,
    stored_filename=stored_name,
    content_type=upload.mimetype,
)
data.setdefault("receipts", []).append(receipt)
```

In `create_expense_from_receipt`, the reviewed form values should still win. The helper can provide defaults:

```python
candidate = build_expense_candidate(receipt.get("extracted", {}))
candidate.update({
    "amount": round(amount, 2),
    "date": tx_date,
    "category": category,
    "subcategory": subcategory,
    "payment_method": request.form.get("payment_method", candidate["payment_method"]).strip() or candidate["payment_method"],
    "receipt_id": receipt_id,
    "source": "receipt",
})
apply_transaction_rules(candidate, data)
data.setdefault("expenses", []).append(candidate)
```

No template changes are required to store the extraction, but a review UI can optionally show `raw_candidates`, `confidence`, and `duplicate_key_candidates` from `receipt["extracted"]`.
