# Recurring Calendar Draft

This draft is isolated from the live Flask app. It adds pure Python helpers
that accept SpendSight's existing JSON data dict and return a JSON-safe
recurring calendar view model. The helpers do not read files, write files,
import Flask, or mutate the input dict.

## Files

- `helpers.py` builds bounded upcoming recurring items, month rows, payment
  status assumptions, and cash-after-bills totals.
- `tests/test_helpers.py` covers fixed expenses, EMI bounds, detected
  subscriptions, goal contributions, trial/renewal metadata, and bounded
  horizons.

## Data Consumed

Existing keys:

- `fixed_expenses`: SpendSight Income page EMI/fixed expense rows.
- `expenses`: existing spend rows used as payment evidence and subscription
  detection history.
- `goals`: monthly savings goal contributions.
- `income` and `extra_income`: monthly cash source totals.

Optional metadata keys:

- `data["recurring_calendar"]["goal_contribution_day"]`
- `data["recurring_calendar"]["trial_renewals"]`
- `data["trial_renewals"]`, `data["trials"]`, `data["renewals"]`
- subscription insights passed directly to `build_recurring_calendar(...,
  subscription_insights=...)` or stored under `subscription_insights`,
  `detected_subscriptions`, or `subscriptions`.

Supported paid evidence includes matching expenses, `paid_months`,
`paid_dates`, `paid_occurrences`, `payment_history`, `payments`,
`contributions`, and `last_paid_on`.

## Parent Wiring

Import in `app.py` when promoted:

```python
from feature_drafts.recurring_calendar import build_recurring_calendar
```

Route sketch:

```python
@app.route("/recurring-calendar")
@login_required
def recurring_calendar():
    data = load_data()
    months = request.args.get("months", default=3, type=int)
    calendar = build_recurring_calendar(data, months_ahead=months)
    return render_spendsight_template(
        "recurring_calendar.html",
        calendar=calendar,
    )
```

Optional JSON route for early integration:

```python
@app.route("/api/recurring-calendar")
@login_required
def recurring_calendar_api():
    data = load_data()
    months = request.args.get("months", default=3, type=int)
    return jsonify(build_recurring_calendar(data, months_ahead=months))
```

Navigation entry in `templates/base.html`, near Planning:

```html
<a href="{{ url_for('recurring_calendar') }}" class="{{ 'active' if request.endpoint == 'recurring_calendar' }}">
  <i class="bi bi-calendar2-week"></i> Recurring Calendar
</a>
```

Template expectation:

- Create `templates/recurring_calendar.html` when the feature is promoted.
- Render `calendar.summary` for totals.
- Iterate `calendar.calendar_rows` for each month.
- Each month row includes `days`, `weeks`, `items`, `scheduled_total`,
  `paid_total`, `unpaid_total`, `cash_after_bills`, and
  `cash_after_unpaid_bills`.
