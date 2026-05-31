# Envelope Budget Draft

This draft stays outside the live Flask app. It adds pure helpers and a Jinja
template draft that can be wired into SpendSight when the parent is ready.

## Files

- `helpers.py` builds the envelope budget view model from SpendSight's existing
  JSON data dict.
- `template_draft.html` is a route template draft using the existing base layout,
  Bootstrap classes, `format_inr`, and CSRF token conventions.
- `tests/test_helpers.py` covers left-to-assign, rollover, sinking funds, fixed
  expense occurrences, and monthly assignment updates.

## Data Shape

The helpers read existing keys first:

- `budget_limits` as monthly category budgets.
- `expenses` as spending.
- `income` and `extra_income` as the assignable income pool.
- `fixed_expenses` as scheduled spending.
- `billing_start_day` for SpendSight billing periods.

Optional draft config lives under `data["envelope_budget"]`:

```python
{
    "start_month": "2026-01",
    "default_rollover": False,
    "rollover_unassigned": True,
    "unassigned_balance": 0,
    "carryover": {"Groceries": 500},
    "monthly_assignments": {"2026-05": {"Groceries": 12000}},
    "category_settings": {
        "Groceries": {"monthly_budget": 12000, "rollover": True},
        "Insurance": {"annual_amount": 24000, "due_month": 12, "rollover": True}
    }
}
```

## Parent Wiring

Route sketch:

```python
from feature_drafts.envelope_budget import build_envelope_budget, build_monthly_assignment_update

@app.route("/envelope-budget", methods=["GET", "POST"])
@login_required
def envelope_budget():
    data = load_data()
    if request.method == "POST":
        action = request.form.get("action")
        if action == "save_assignments":
            assignments = {
                key.removeprefix("assignment__"): value
                for key, value in request.form.items()
                if key.startswith("assignment__")
            }
            data["envelope_budget"] = build_monthly_assignment_update(
                data,
                assignments,
                period_key=request.form.get("period_key"),
            )
            save_data(data)
            flash("Envelope assignments saved.", "success")
            return redirect(url_for("envelope_budget"))
        if action == "save_settings":
            config = dict(data.get("envelope_budget") or {})
            config["start_month"] = request.form.get("start_month")
            config["default_rollover"] = bool(request.form.get("default_rollover"))
            config["rollover_unassigned"] = bool(request.form.get("rollover_unassigned"))
            data["envelope_budget"] = config
            save_data(data)
            return redirect(url_for("envelope_budget"))

    return render_spendsight_template(
        "envelope_budget.html",
        budget=build_envelope_budget(data),
    )
```

Template wiring:

1. Move or copy `feature_drafts/envelope_budget/template_draft.html` to
   `templates/envelope_budget.html`.
2. Add a sidebar nav entry in `templates/base.html`, near Planning:
   `url_for('envelope_budget')`, icon `bi-wallet2`, label `Envelope Budget`.
3. Keep using `render_spendsight_template` so `format_inr`, `currency_symbol`,
   users, CSRF, and existing globals are available.
