# Receipts Inbox Draft

This draft keeps all product code outside the live Flask app. The helpers store review state under `data["receipt_inbox"]`, which is compatible with SpendSight's current JSON user-data dict because unknown top-level keys are preserved.

## Data Model

Each inbox item is a JSON-safe dict:

- `id`, `status`, `created_at`, `updated_at`
- `attachment`: filename, content type, size, checksum, storage path, upload source
- `ocr`: status, engine, language, raw text, confidence, processed timestamp, error
- `extracted`: OCR-ready transaction fields such as merchant, date, amount, category, subcategory, payment method, notes, quantity, unit, line items
- `duplicate_candidates`: scored matches against existing `data["expenses"]`
- `review`: decision, reviewer, reviewed timestamp, notes, posted expense id

Statuses are `needs_review`, `duplicate_candidate`, `ready_to_post`, `posted`, and `dismissed`.

## Parent Wiring

Recommended import in `app.py` when this is promoted:

```python
from feature_drafts.receipts_inbox import (
    apply_review_update,
    build_expense_from_inbox_item,
    build_review_inbox,
    create_receipt_inbox_item,
    dismiss_inbox_item,
    get_inbox_item,
    mark_item_posted,
    new_attachment_metadata,
)
```

Add `receipt_inbox` to `_default_data()` and the list-normalization keys in `_normalize_data()` when the feature is no longer a draft. The helpers already call `ensure_receipt_inbox(data)`, so this is optional during experimentation.

Route shape:

```python
@app.route("/receipts")
@login_required
def receipts_inbox():
    data = load_data()
    return render_spendsight_template(
        "receipts_inbox.html",
        inbox=build_review_inbox(data, status=request.args.get("status")),
        categories=get_all_categories(data),
        payment_methods=data.get("payment_methods", DEFAULT_PAYMENT_METHODS),
    )

@app.route("/receipts/upload", methods=["POST"])
@login_required
def receipts_inbox_upload():
    data = load_data()
    upload = request.files["receipt_file"]
    attachment = new_attachment_metadata(
        upload.filename,
        content_type=upload.mimetype,
        size_bytes=request.content_length,
        storage_path="",  # fill after saving the file
        source=request.form.get("source", "upload"),
    )
    create_receipt_inbox_item(data, attachment)
    save_data(data)
    return redirect(url_for("receipts_inbox"))

@app.route("/receipts/<item_id>/update", methods=["POST"])
@login_required
def receipts_inbox_update(item_id):
    data = load_data()
    apply_review_update(
        data,
        item_id,
        extracted_updates={
            "merchant": request.form.get("merchant", ""),
            "date": request.form.get("date", ""),
            "amount": request.form.get("amount", ""),
            "category": request.form.get("category", ""),
            "subcategory": request.form.get("subcategory", ""),
            "payment_method": request.form.get("payment_method", ""),
            "notes": request.form.get("notes", ""),
        },
        status=request.form.get("status"),
        reviewed_by=current_user.id,
    )
    save_data(data)
    return redirect(url_for("receipts_inbox"))

@app.route("/receipts/<item_id>/post", methods=["POST"])
@login_required
def receipts_inbox_post(item_id):
    data = load_data()
    item = get_inbox_item(data, item_id)
    expense = build_expense_from_inbox_item(item)
    apply_transaction_rules(expense, data)
    data["expenses"].append(expense)
    mark_item_posted(data, item_id, expense["id"], reviewed_by=current_user.id)
    save_data(data)
    return redirect(url_for("receipts_inbox"))

@app.route("/receipts/<item_id>/dismiss", methods=["POST"])
@login_required
def receipts_inbox_dismiss(item_id):
    data = load_data()
    dismiss_inbox_item(data, item_id, reviewed_by=current_user.id)
    save_data(data)
    return redirect(url_for("receipts_inbox"))
```

The template draft expects endpoint names `receipts_inbox_upload`, `receipts_inbox_update`, `receipts_inbox_post`, and `receipts_inbox_dismiss`. Either copy `template_draft.html` to `templates/receipts_inbox.html` when promoting or register this folder as a blueprint template folder.

Navigation can sit under the current Main group after Import CSV:

```html
<a href="{{ url_for('receipts_inbox') }}" class="{{ 'active' if request.endpoint.startswith('receipts_inbox') }}">
  <i class="bi bi-receipt-cutoff"></i> Receipts Inbox
</a>
```
