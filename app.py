
"""
SpendSight - Personal Finance Tracker
Local Flask web app — data stored in user_data/data_<user_id>.json
Run: python app.py  then open http://localhost:5000
"""

from flask import Flask, render_template, request, redirect, url_for, jsonify, flash, session, Response, g
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps
import json
import os
from datetime import date, datetime, timedelta
from calendar import monthrange
import uuid
from collections import defaultdict
import webbrowser
import threading
import secrets
import urllib.parse
import urllib.request
import csv
import io

app = Flask(__name__)

# ── Secret key: env var → persistent file → auto-generate ────────────────────
_KEY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".secret_key")
def _load_secret_key():
    if os.environ.get("SECRET_KEY"):
        return os.environ["SECRET_KEY"]
    if os.path.exists(_KEY_FILE):
        with open(_KEY_FILE, "r") as _f:
            key = _f.read().strip()
            if key:
                return key
    key = secrets.token_hex(32)
    with open(_KEY_FILE, "w") as _f:
        _f.write(key)
    return key

app.secret_key = _load_secret_key()

USERS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "users.json")
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_data")
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

# ── Login rate limiting (in-memory, no extra package) ────────────────────────
# Tracks failed attempts per IP: { ip: {"count": int, "locked_until": datetime|None} }
_login_attempts: dict = {}
_MAX_ATTEMPTS  = 5          # lock after 5 consecutive failures
_LOCKOUT_SECS  = 60         # locked for 60 seconds

def _get_ip() -> str:
    return request.headers.get("X-Forwarded-For", request.remote_addr or "unknown").split(",")[0].strip()

def _check_rate_limit() -> tuple[bool, int]:
    """Return (is_blocked, seconds_remaining). Cleans up expired locks."""
    ip  = _get_ip()
    rec = _login_attempts.get(ip)
    if not rec:
        return False, 0
    if rec["locked_until"] and datetime.now() < rec["locked_until"]:
        secs = int((rec["locked_until"] - datetime.now()).total_seconds()) + 1
        return True, secs
    if rec["locked_until"] and datetime.now() >= rec["locked_until"]:
        _login_attempts.pop(ip, None)   # lock expired, reset
    return False, 0

def _record_failure():
    ip  = _get_ip()
    rec = _login_attempts.setdefault(ip, {"count": 0, "locked_until": None})
    rec["count"] += 1
    if rec["count"] >= _MAX_ATTEMPTS:
        rec["locked_until"] = datetime.now() + timedelta(seconds=_LOCKOUT_SECS)

def _clear_rate_limit():
    _login_attempts.pop(_get_ip(), None)

def _validate_password_strength(pw: str) -> str | None:
    """Return an error message if password doesn't meet complexity rules, else None."""
    import re
    if len(pw) < 8:
        return "Password must be at least 8 characters."
    if not re.search(r"[A-Z]", pw):
        return "Password must contain at least one uppercase letter."
    if not re.search(r"[a-z]", pw):
        return "Password must contain at least one lowercase letter."
    if not re.search(r"\d", pw):
        return "Password must contain at least one number."
    if not re.search(r"[^A-Za-z0-9]", pw):
        return "Password must contain at least one symbol (!@#$%^&* etc.)."
    return None

CLOUD_TOKENS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cloud_tokens.json")
REDIRECT_BASE = "http://localhost:5000"

# ── Developer OAuth credentials (fill these in once after registering your apps) ──
# Google Drive:  console.cloud.google.com  → OAuth 2.0 → Desktop app
# OneDrive:      portal.azure.com          → App registrations → Mobile/Desktop
# Dropbox:       dropbox.com/developers    → Create app
# Redirect URI for all three: http://localhost:5000/cloud/callback/<service>
CLOUD_CREDENTIALS = {
    "gdrive": {
        "client_id":     "YOUR_GOOGLE_CLIENT_ID",
        "client_secret": "YOUR_GOOGLE_CLIENT_SECRET",
    },
    "onedrive": {
        "client_id": "YOUR_MICROSOFT_CLIENT_ID",
    },
    "dropbox": {
        "app_key":    "YOUR_DROPBOX_APP_KEY",
        "app_secret": "YOUR_DROPBOX_APP_SECRET",
    },
}

def load_cloud_tokens():
    if not os.path.exists(CLOUD_TOKENS_FILE):
        return {}
    with open(CLOUD_TOKENS_FILE) as f:
        return json.load(f)

def save_cloud_tokens(tokens):
    with open(CLOUD_TOKENS_FILE, "w") as f:
        json.dump(tokens, f, indent=2)


def _is_hashed(pw: str) -> bool:
    """Return True if the password string is already a werkzeug hash."""
    return pw.startswith(("pbkdf2:", "scrypt:", "argon2:"))

def load_users():
    if not os.path.exists(USERS_FILE):
        users = [{"id": "admin", "password": generate_password_hash("admin123"), "role": "admin"}]
        save_users(users)
        return users
    with open(USERS_FILE, "r", encoding="utf-8") as f:
        users = json.load(f)
    # ── Migrate any plaintext passwords to hashed on first load ──────────────
    migrated = False
    for u in users:
        if not _is_hashed(u.get("password", "")):
            u["password"] = generate_password_hash(u["password"])
            migrated = True
    if migrated:
        save_users(users)
    return users

def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, indent=2, ensure_ascii=False)

def render_spendsight_template(template_name, **kwargs):
    # Inject user info for admin to switch views
    users = load_users()
    current_view_id = session.get("view_user_id", current_user.id if current_user.is_authenticated else "admin")
    
    # Ensure currency_symbol and currency_code are available to prevent template errors
    from flask import g
    currency_symbol = getattr(g, "currency_symbol", "₹")
    currency_code = getattr(g, "currency_code", "INR")

    return render_template(template_name,
                           all_users=users,
                           current_view_id=current_view_id,
                           currency_symbol=currency_symbol,
                           currency_code=currency_code,
                           unit_options=UNIT_OPTIONS,
                           **kwargs)

def get_current_data_file():
    user_id = current_user.id if current_user.is_authenticated else "admin"
    if current_user.is_authenticated and current_user.role == "admin":
        user_id = session.get("view_user_id", current_user.id)
    return os.path.join(DATA_DIR, f"data_{user_id}.json")

# ── Login Configuration ───────────────────────────────────────────────────────

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"

class User(UserMixin):
    def __init__(self, id, role="user"):
        self.id = id
        self.role = role

@login_manager.user_loader
def load_user(user_id):
    users = load_users()
    user_data = next((u for u in users if u["id"] == user_id), None)
    if user_data:
        return User(user_data["id"], user_data.get("role", "user"))
    return None

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not current_user.is_authenticated or current_user.role != "admin":
            flash("You do not have permission to access this page.", "danger")
            return redirect(url_for("dashboard"))
        return f(*args, **kwargs)
    return decorated_function

@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        # ── Rate-limit check ─────────────────────────────────────────────────
        blocked, secs = _check_rate_limit()
        if blocked:
            flash(f"Too many failed attempts. Please wait {secs} second{'s' if secs != 1 else ''} before trying again.", "danger")
            return render_template("login.html")

        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        users     = load_users()
        user_data = next((u for u in users if u["id"].lower() == username.lower()), None)

        if user_data and check_password_hash(user_data["password"], password):
            _clear_rate_limit()                      # reset on success
            user = User(user_data["id"], user_data.get("role", "user"))
            login_user(user)
            session["view_user_id"] = user.id
            return redirect(url_for("dashboard"))
        else:
            _record_failure()                        # count the failed attempt
            blocked, secs = _check_rate_limit()
            if blocked:
                flash(f"Too many failed attempts. Locked for {secs} seconds.", "danger")
            else:
                flash("Invalid username or password.", "danger")

    return render_template("login.html")

@app.route("/logout")
@login_required
def logout():
    logout_user()
    session.pop("view_user_id", None)
    return redirect(url_for("login"))

@app.route("/admin")
@login_required
@admin_required
def admin_dashboard():
    users = load_users()
    return render_spendsight_template("admin.html", users=users)

@app.route("/admin/user/add", methods=["POST"])
@login_required
@admin_required
def admin_add_user():
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "").strip()
    role     = request.form.get("role", "user")

    if not username or not password:
        flash("Username and password are required.", "danger")
        return redirect(url_for("admin_dashboard"))

    pw_error = _validate_password_strength(password)
    if pw_error:
        flash(pw_error, "danger")
        return redirect(url_for("admin_dashboard"))

    users = load_users()
    if any(u["id"].lower() == username.lower() for u in users):
        flash(f"User '{username}' already exists.", "danger")
        return redirect(url_for("admin_dashboard"))

    users.append({
        "id": username,
        "password": generate_password_hash(password),
        "role": role
    })
    save_users(users)
    flash(f"User '{username}' added successfully.", "success")
    return redirect(url_for("admin_dashboard"))

@app.route("/admin/user/delete/<uid>", methods=["POST"])
@login_required
@admin_required
def admin_delete_user(uid):
    if uid == current_user.id:
        flash("You cannot delete yourself.", "danger")
        return redirect(url_for("admin_dashboard"))

    users = load_users()
    admins = [u for u in users if u.get("role") == "admin"]
    user_to_delete = next((u for u in users if u["id"] == uid), None)

    if user_to_delete and user_to_delete.get("role") == "admin" and len(admins) <= 1:
        flash("Cannot delete the last admin user.", "danger")
        return redirect(url_for("admin_dashboard"))

    users = [u for u in users if u["id"] != uid]
    save_users(users)
    if session.get("view_user_id") == uid:
        session["view_user_id"] = current_user.id
    flash(f"User '{uid}' deleted.", "info")
    return redirect(url_for("admin_dashboard"))

@app.route("/admin/user/reset_password/<uid>", methods=["POST"])
@login_required
@admin_required
def admin_reset_password(uid):
    new_password = request.form.get("new_password", "").strip()
    if not new_password:
        flash("New password cannot be empty.", "danger")
        return redirect(url_for("admin_dashboard"))

    pw_error = _validate_password_strength(new_password)
    if pw_error:
        flash(pw_error, "danger")
        return redirect(url_for("admin_dashboard"))

    users = load_users()
    user = next((u for u in users if u["id"] == uid), None)
    if user:
        user["password"] = generate_password_hash(new_password)
        save_users(users)
        flash(f"Password for '{uid}' updated.", "success")
    else:
        flash("User not found.", "danger")

    return redirect(url_for("admin_dashboard"))

@app.route("/admin/switch-user", methods=["POST"])
@login_required
@admin_required
def admin_switch_user():
    target_user_id = request.form.get("user_id", "").strip()
    users          = load_users()
    valid_ids      = [u["id"] for u in users]
    if target_user_id not in valid_ids:
        flash("Invalid user selection.", "danger")
        return redirect(request.referrer or url_for("dashboard"))
    session["view_user_id"] = target_user_id
    if target_user_id == current_user.id:
        flash("Switched back to your own data.", "info")
    else:
        flash(f"Now viewing data for: {target_user_id}", "info")
    return redirect(request.referrer or url_for("dashboard"))

# ── Default data ──────────────────────────────────────────────────────────────

DEFAULT_CATEGORIES = {
    "Groceries":    ["Milk", "Dal", "Rice", "Sugar", "Atta", "Oil", "Soap",
                     "Shampoo", "Vegetables", "Fruits", "Eggs", "Pulses", "Ghee", "Butter", "Tea"],       
    "Fast Food":    ["McDonald's", "Wada Pav", "Pani Puri", "Samosa", "Pizza",
                     "Chai", "Biryani", "Thali", "Dosa", "Idli", "Burger", "Noodles"],
    "Fuel":         ["Petrol", "Diesel", "CNG"],
    "Utilities":    ["Electricity", "Water", "Gas Cylinder", "Internet", "Mobile Recharge", "DTH"],       
    "Entertainment":["Netflix", "Amazon Prime", "Hotstar", "Movies", "Games", "Events", "Spotify"],       
    "Health":       ["Medicine", "Doctor", "Gym", "Supplements", "Lab Test", "Dental"],
    "Transport":    ["Ola/Uber", "Local Train", "Bus", "Auto", "Metro", "Parking"],
    "Shopping":     ["Clothing", "Electronics", "Household", "Personal Care", "Accessories", "Books"],    
    "EMI / Finance":["Card EMI", "Loan EMI", "Insurance Premium", "SIP", "Rent"],
    "Other":        ["Miscellaneous", "Gifts", "Donations", "Fees"],
}

CATEGORY_COLORS = {
    "Groceries":     "#0E9F6E",
    "Fast Food":     "#E3A008",
    "Fuel":          "#1A56DB",
    "Utilities":     "#4B5563",
    "Entertainment": "#7E3AF2",
    "Health":        "#F05252",
    "Transport":     "#FF8C00",
    "Shopping":      "#0694A2",
    "EMI / Finance": "#C81E1E",
    "Other":         "#374151",
}

DEFAULT_PAYMENT_METHODS = ["Cash", "Card 1", "Card 2", "Card 3"]
UNIT_OPTIONS = ["pcs", "kg", "g", "L", "ml", "pack", "dozen", "box"]

CURRENCIES = [
    {"code": "INR", "symbol": "₹",   "name": "Indian Rupee"},
    {"code": "USD", "symbol": "$",    "name": "US Dollar"},
    {"code": "EUR", "symbol": "€",    "name": "Euro"},
    {"code": "GBP", "symbol": "£",    "name": "British Pound"},
    {"code": "AED", "symbol": "د.إ",  "name": "UAE Dirham"},
    {"code": "SGD", "symbol": "S$",   "name": "Singapore Dollar"},
    {"code": "CAD", "symbol": "CA$",  "name": "Canadian Dollar"},
    {"code": "AUD", "symbol": "A$",   "name": "Australian Dollar"},
    {"code": "JPY", "symbol": "¥",    "name": "Japanese Yen"},
    {"code": "MYR", "symbol": "RM",   "name": "Malaysian Ringgit"},
    {"code": "BRL", "symbol": "R$",   "name": "Brazilian Real"},
    {"code": "NGN", "symbol": "₦",    "name": "Nigerian Naira"},
    {"code": "ZAR", "symbol": "R",    "name": "South African Rand"},
    {"code": "PKR", "symbol": "Rs",   "name": "Pakistani Rupee"},
    {"code": "BDT", "symbol": "৳",    "name": "Bangladeshi Taka"},
]

# ── Data helpers ──────────────────────────────────────────────────────────────

def load_data():
    data_file = get_current_data_file()
    if not os.path.exists(data_file):
        data = {
            "expenses": [],
            "templates": [],
            "custom_categories": {},
            "payment_methods": DEFAULT_PAYMENT_METHODS,
            "billing_start_day": 1,
            "income": {"monthly_salary": 0, "salary_history": []},
            "extra_income": [],
            "fixed_expenses": [],
            "recurring_payments": [],
            "budget_limits": {},
        }
        # Special migration for admin: if old expenses.json exists, move it to data_admin.json
        old_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "expenses.json")
        if "data_admin.json" in data_file and os.path.exists(old_file):
             with open(old_file, "r", encoding="utf-8") as f:
                 old_data = json.load(f)
             for k, v in old_data.items():
                 if k != "users": data[k] = v
             save_data(data)
        else:
             save_data(data)
        return data

    with open(data_file, "r", encoding="utf-8") as f:
        d = json.load(f)

    d.setdefault("expenses", [])
    d.setdefault("templates", [])
    d.setdefault("custom_categories", {})
    d.setdefault("payment_methods", DEFAULT_PAYMENT_METHODS)
    d.setdefault("billing_start_day", 1)
    d.setdefault("income", {"monthly_salary": 0, "salary_history": []})
    d.setdefault("extra_income", [])
    d.setdefault("fixed_expenses", [])
    d.setdefault("recurring_payments", [])
    d.setdefault("currency_code", "INR")
    d.setdefault("budget_limits", {})

    for e in d["expenses"]:
        e.setdefault("quantity", None)
        e.setdefault("unit", "")

    # Ensure salary history migration
    income = d.get("income", {"monthly_salary": 0, "salary_history": []})
    if "salary_history" not in income:
        sal = float(income.get("monthly_salary", 0))
        if sal > 0:
            effective = income.get("salary_updated", date.today().isoformat())
            income["salary_history"] = [{"amount": round(sal, 2), "effective_from": effective, "added_on": effective}]
        else:
            income["salary_history"] = []
        d["income"] = income
    return d

def save_data(data):
    data_file = get_current_data_file()
    with open(data_file, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False, default=str)

def get_salary_for_date(income_data, target_date):
    history = income_data.get("salary_history", [])
    if not history:
        return float(income_data.get("monthly_salary", 0))
    target_str = target_date.isoformat() if hasattr(target_date, "isoformat") else str(target_date)       
    applicable = [h for h in history if h["effective_from"] <= target_str]
    if not applicable:
        return 0.0
    return float(sorted(applicable, key=lambda h: h["effective_from"])[-1]["amount"])

def get_all_categories(data):
    cats = {k: list(v) for k, v in DEFAULT_CATEGORIES.items()}
    for cat, subs in data.get("custom_categories", {}).items():
        if cat in cats:
            for s in subs:
                if s not in cats[cat]:
                    cats[cat].append(s)
        else:
            cats[cat] = list(subs)

    sorted_cats = {}
    for cat in sorted(cats.keys(), key=str.casefold):
        sorted_cats[cat] = sorted({s for s in cats[cat] if s}, key=str.casefold)
    return sorted_cats

def today_str():
    return date.today().isoformat()

def parse_quantity_value(raw_value):
    try:
        if raw_value in (None, ""):
            return None
        qty = float(str(raw_value).replace(",", "").strip())
        return round(qty, 3) if qty > 0 else None
    except (ValueError, TypeError):
        return None

def clean_unit_value(raw_value):
    return (raw_value or "").strip()

def format_quantity_display(quantity, unit):
    qty = parse_quantity_value(quantity)
    unit = clean_unit_value(unit)
    if not qty:
        return "-"
    if abs(qty - round(qty)) < 1e-9:
        qty_text = str(int(round(qty)))
    else:
        qty_text = f"{qty:.3f}".rstrip("0").rstrip(".")
    return f"{qty_text} {unit}".strip()

def build_purchase_audit_rows(expenses, category_filter="Groceries"):
    today = date.today()
    grouped = defaultdict(list)

    for e in expenses:
        category = (e.get("category") or "Other").strip() or "Other"
        if category_filter and category_filter != "All" and category != category_filter:
            continue
        subcategory = (e.get("subcategory") or category).strip() or category
        try:
            dt = date.fromisoformat(e.get("date", today.isoformat()))
        except Exception:
            continue
        grouped[(category, subcategory)].append({
            "date_obj": dt,
            "amount": float(e.get("amount", 0) or 0),
            "quantity": parse_quantity_value(e.get("quantity")),
            "unit": clean_unit_value(e.get("unit")),
        })

    rows = []
    for (category, subcategory), entries in grouped.items():
        entries.sort(key=lambda x: x["date_obj"])
        purchase_count = len(entries)
        last_entry = entries[-1]
        last_purchase = last_entry["date_obj"]
        days_since_last = (today - last_purchase).days

        gaps = []
        for prev, curr in zip(entries, entries[1:]):
            gap = (curr["date_obj"] - prev["date_obj"]).days
            if gap > 0:
                gaps.append(gap)
        avg_gap = round(sum(gaps) / len(gaps), 1) if gaps else None
        next_expected = last_purchase + timedelta(days=int(round(avg_gap))) if avg_gap else None

        qty_entries = [x for x in entries if x["quantity"] and x["unit"]]
        units = sorted({x["unit"] for x in qty_entries}, key=str.casefold)
        single_unit = units[0] if len(units) == 1 else ""
        total_qty = sum(x["quantity"] for x in qty_entries) if single_unit else None

        total_prev_qty = 0.0
        total_gap_days = 0
        same_unit_entries = [x for x in qty_entries if x["unit"] == single_unit] if single_unit else []
        for prev, curr in zip(same_unit_entries, same_unit_entries[1:]):
            gap = (curr["date_obj"] - prev["date_obj"]).days
            if gap > 0 and prev["quantity"]:
                total_prev_qty += prev["quantity"]
                total_gap_days += gap

        daily_usage = round(total_prev_qty / total_gap_days, 3) if total_gap_days and total_prev_qty else None
        last_qty = same_unit_entries[-1]["quantity"] if same_unit_entries else None
        stock_days = round(last_qty / daily_usage, 1) if daily_usage and last_qty else None
        stock_until = last_purchase + timedelta(days=int(round(stock_days))) if stock_days else None
        remaining_days = (stock_until - today).days if stock_until else None

        pattern_status = "Need more history"
        if avg_gap:
            if days_since_last > avg_gap * 1.5:
                pattern_status = "Delayed repurchase"
            elif gaps and gaps[-1] < avg_gap * 0.6:
                pattern_status = "Bought earlier than usual"
            else:
                pattern_status = "On pattern"
        elif purchase_count == 1:
            pattern_status = "First recorded purchase"

        usage_status = "No quantity data"
        if stock_days:
            if remaining_days is not None and remaining_days < 0:
                usage_status = "Past expected depletion"
            elif remaining_days is not None and remaining_days <= 3:
                usage_status = "Running low soon"
            else:
                usage_status = "Stock coverage looks healthy"
        elif qty_entries and not single_unit:
            usage_status = "Mixed units"

        insight_bits = []
        if avg_gap:
            insight_bits.append(f"Avg gap {avg_gap:.1f} days")
        if daily_usage and single_unit:
            insight_bits.append(f"Usage {daily_usage:.3f} {single_unit}/day")
        if stock_days and single_unit:
            insight_bits.append(f"Current buy may last {stock_days:.1f} days")
        if not insight_bits:
            insight_bits.append("Add quantity and unit on future purchases for stock estimates")

        rows.append({
            "category": category,
            "subcategory": subcategory,
            "purchase_count": purchase_count,
            "total_spent": round(sum(x["amount"] for x in entries), 2),
            "last_purchase": last_purchase.isoformat(),
            "days_since_last": days_since_last,
            "avg_gap_days": avg_gap,
            "next_expected": next_expected.isoformat() if next_expected else "",
            "pattern_status": pattern_status,
            "quantity_display": format_quantity_display(last_qty, single_unit) if last_qty and single_unit else "-",
            "total_quantity_display": format_quantity_display(total_qty, single_unit) if total_qty and single_unit else (", ".join(units) if units else "-"),
            "daily_usage_display": f"{daily_usage:.3f} {single_unit}/day" if daily_usage and single_unit else "-",
            "stock_days": stock_days,
            "stock_until": stock_until.isoformat() if stock_until else "",
            "usage_status": usage_status,
            "insight": " | ".join(insight_bits),
            "needs_attention": pattern_status != "On pattern" or usage_status in {"Past expected depletion", "Running low soon"},
            "has_usage_model": bool(stock_days),
        })

    rows.sort(key=lambda r: (not r["needs_attention"], r["category"].casefold(), r["subcategory"].casefold()))
    summary = {
        "tracked_items": len(rows),
        "attention_count": sum(1 for r in rows if r["needs_attention"]),
        "usage_ready_count": sum(1 for r in rows if r["has_usage_model"]),
        "total_spent": round(sum(r["total_spent"] for r in rows), 2),
    }
    return rows, summary

def get_billing_period(billing_start_day=1):
    today = date.today()
    day = min(billing_start_day, 28)
    if today.day >= day:
        start = today.replace(day=day)
        if start.month == 12:
            end = date(start.year + 1, 1, day) - timedelta(days=1)
        else:
            end = date(start.year, start.month + 1, day) - timedelta(days=1)
    else:
        if today.month == 1:
            start = date(today.year - 1, 12, day)
        else:
            start = today.replace(month=today.month - 1, day=day)
        end = today.replace(day=day) - timedelta(days=1)
    return start, end


def billing_period_label(bsd):
    s, e = get_billing_period(bsd)
    return f"{s.strftime('%b %d')} – {e.strftime('%b %d')}"

def is_emi_active_in_month(emi, year, month):
    sy = int(emi.get('start_year', 0))
    sm = int(emi.get('start_month', 0))
    if emi.get('type') == 'fixed':
        return (year > sy) or (year == sy and month >= sm)
    tm = int(emi.get('total_months', 0))
    start_val = sy * 12 + (sm - 1)
    end_val   = start_val + tm - 1
    cur_val   = year * 12 + (month - 1)
    return start_val <= cur_val <= end_val

def get_month_extra_income(extra_income_list, year, month):
    """Sum extra income entries that apply to the given calendar month.
    Handles both one-time (date match) and recurring (frequency) entries.
    """
    total = 0.0
    for ei in extra_income_list:
        try:
            ei_type = ei.get("type", "one-time")
            if ei_type == "recurring":
                start = date.fromisoformat(ei.get("start_date", ""))
                if (year, month) < (start.year, start.month):
                    continue
                end_s = ei.get("end_date", "").strip()
                if end_s:
                    end_d = date.fromisoformat(end_s)
                    if (year, month) > (end_d.year, end_d.month):
                        continue
                months_since = (year - start.year) * 12 + (month - start.month)
                freq = ei.get("frequency", "monthly")
                if freq == "monthly":
                    applies = True
                elif freq == "quarterly":
                    applies = months_since % 3 == 0
                elif freq == "half_yearly":
                    applies = months_since % 6 == 0
                elif freq == "yearly":
                    applies = months_since % 12 == 0
                else:
                    applies = True
                if applies:
                    total += float(ei["amount"])
            else:
                d = date.fromisoformat(ei["date"])
                if d.year == year and d.month == month:
                    total += float(ei["amount"])
        except Exception:
            pass
    return total


def get_billing_period_for_n_ago(n, bsd, ref_today=None):
    """Get the billing period that is n periods ago (0 = current)."""
    ref = ref_today or date.today()
    day = min(bsd, 28)
    curr_start, curr_end = get_billing_period(bsd)
    if n == 0:
        return curr_start, curr_end
    month = curr_start.month - n
    year  = curr_start.year
    while month <= 0:
        month += 12
        year  -= 1
    start = date(year, month, min(day, monthrange(year, month)[1]))
    next_m = month + 1
    next_y = year
    if next_m > 12:
        next_m = 1
        next_y += 1
    end = date(next_y, next_m, min(day, monthrange(next_y, next_m)[1])) - timedelta(days=1)
    return start, end


def get_n_billing_months_ago(n, billing_start_day):
    """Return (start_date, end_date) spanning N billing months back from current period end."""
    curr_start, curr_end = get_billing_period(billing_start_day)
    day = min(billing_start_day, 28)
    month = curr_start.month - n
    year  = curr_start.year
    while month <= 0:
        month += 12
        year  -= 1
    max_day = monthrange(year, month)[1]
    start   = date(year, month, min(day, max_day))
    return start, curr_end


def get_emi_status(emi, ref_date=None):
    """Return status dict for an EMI or Fixed Expense.
    Keys: paid, remaining, total, is_active, end_year, end_month,
          progress_pct, amount_due, type
    """
    ref      = ref_date or date.today()
    emi_type = emi.get("type", "emi")
    sy, sm   = int(emi.get("start_year", 0)), int(emi.get("start_month", 0))

    if emi_type == "emi":
        total     = int(emi.get("total_months", 0))
        elapsed   = (ref.year - sy) * 12 + (ref.month - sm) + 1
        paid      = min(max(elapsed, 0), total)
        remaining = total - paid
        is_active = paid < total

        offset = total - 1
        ey = sy + (sm - 1 + offset) // 12
        em = (sm - 1 + offset) % 12 + 1

        return {
            "paid":         paid,
            "remaining":    remaining,
            "total":        total,
            "is_active":    is_active,
            "end_year":     ey,
            "end_month":    em,
            "months_left":  remaining,
            "progress_pct": round(paid / total * 100) if total else 100,
            "amount_due":   float(emi["amount"]) if is_active else 0.0,
            "type":         "emi",
        }
    else:
        is_active      = (ref.year > sy) or (ref.year == sy and ref.month >= sm)
        due_this_month = is_emi_active_in_month(emi, ref.year, ref.month)
        return {
            "paid":         0,
            "remaining":    0,
            "total":        0,
            "is_active":    is_active,
            "end_year":     None,
            "end_month":    None,
            "months_left":  None,
            "progress_pct": 0,
            "amount_due":   float(emi["amount"]) if due_this_month else 0.0,
            "type":         "fixed",
            "frequency":    emi.get("frequency", "monthly"),
            "day_of_month": emi.get("day_of_month", 1),
        }

@app.before_request
def inject_currency():
    if not current_user.is_authenticated:
        g.currency_symbol = "₹"
        g.currency_code = "INR"
        return
    data = load_data()
    cc = data.get("currency_code", "INR")
    g.currency_code = cc
    g.currency_symbol = next((c["symbol"] for c in CURRENCIES if c["code"] == cc), "₹")

def fmtINR(v):
    symbol = getattr(g, "currency_symbol", "₹")
    code   = getattr(g, "currency_code", "INR")
    if code != "INR":
        return f"{symbol}{v:,.0f}"
    # Indian lakh formatting: last 3 digits, then groups of 2
    v = int(round(v))
    if v < 0:
        return f"-{fmtINR(-v)}"
    s = str(v)
    if len(s) <= 3:
        return f"{symbol}{s}"
    last3 = s[-3:]
    rest  = s[:-3]
    groups = []
    while len(rest) > 2:
        groups.insert(0, rest[-2:])
        rest = rest[:-2]
    if rest:
        groups.insert(0, rest)
    return f"{symbol}{','.join(groups)},{last3}"

app.jinja_env.globals.update(
    format_inr=fmtINR,
    today_str=today_str,
    get_category_color=lambda c: CATEGORY_COLORS.get(c, "#4B5563"),
    billing_period_label=billing_period_label,
)
@app.route("/")
@login_required
def dashboard():
    data = load_data()
    expenses = data["expenses"]
    bsd = int(data.get("billing_start_day", 1))
    today = today_str()
    today_dt = date.today()
    bill_start, bill_end = get_billing_period(bsd)
    bill_start_s, bill_end_s = bill_start.isoformat(), bill_end.isoformat()
    today_expenses = [e for e in expenses if e["date"] == today]
    billing_expenses = [e for e in expenses if bill_start_s <= e["date"] <= bill_end_s]
    latest_today_expenses = sorted(
        today_expenses,
        key=lambda x: (x["date"], x.get("created_at", "")),
        reverse=True
    )[:2]
    today_total = sum(e["amount"] for e in today_expenses)
    billing_total = sum(e["amount"] for e in billing_expenses)
    cash_cycle_expenses = [e for e in billing_expenses if (e.get("payment_method") or "").strip().lower() == "cash"]
    card_cycle_expenses = [e for e in billing_expenses if (e.get("payment_method") or "").strip().lower() != "cash"]
    cash_cycle_total = sum(e["amount"] for e in cash_cycle_expenses)
    card_cycle_total = sum(e["amount"] for e in card_cycle_expenses)
    cash_cycle_count = len(cash_cycle_expenses)
    card_cycle_count = len(card_cycle_expenses)
    biggest_cycle_expenses = sorted(
        billing_expenses,
        key=lambda x: (x["amount"], x["date"], x.get("created_at", "")),
        reverse=True
    )[:2]
    cat_totals = {}
    for e in billing_expenses: 
        cat = e["category"]
        cat_totals[cat] = cat_totals.get(cat, 0) + e["amount"]
    top_categories = sorted(cat_totals.items(), key=lambda x: x[1], reverse=True)[:5]
    pm_totals = {}
    for e in billing_expenses:
        pm = e["payment_method"]
        pm_totals[pm] = pm_totals.get(pm, 0) + e["amount"]
    recent = sorted(expenses, key=lambda x: x["date"], reverse=True)[:10]
    
    elapsed = (date.today() - bill_start).days + 1
    total_days = (bill_end - bill_start).days + 1
    daily_avg = billing_total / max(1, elapsed)
    projected = daily_avg * total_days

    prev_month_totals = []
    for offset in (1, 2):
        month = today_dt.month - offset
        year = today_dt.year
        while month <= 0:
            month += 12
            year -= 1
        month_prefix = f"{year}-{month:02d}"
        month_total = sum(
            e["amount"] for e in expenses
            if e["date"].startswith(month_prefix)
        )
        prev_month_totals.append({
            "label": f"{date(year, month, 1).strftime('%b %Y')}",
            "total": round(month_total, 2),
        })

    budget_status  = get_budget_status(data)
    smart_insights = get_smart_insights(data)

    return render_spendsight_template("dashboard.html", today_total=today_total, billing_total=billing_total,
                                   card_cycle_total=card_cycle_total,
                                   cash_cycle_total=cash_cycle_total,
                                   card_cycle_count=card_cycle_count,
                                   cash_cycle_count=cash_cycle_count,
                                   latest_today_expenses=latest_today_expenses,
                                   biggest_cycle_expenses=biggest_cycle_expenses,
                                   prev_month_totals=prev_month_totals,
                                   projected=projected,
                                   daily_avg=daily_avg,
                                   billing_label=billing_period_label(bsd), days_elapsed=elapsed,
                                   days_total=total_days, top_categories=top_categories,
                                   pm_totals=pm_totals, recent=recent, today_expenses=today_expenses,
                                   budget_status=budget_status, smart_insights=smart_insights)

@app.route("/add", methods=["GET", "POST"])
@login_required
def add_expense():
    if request.method == "POST":
        data       = load_data()
        categories = get_all_categories(data)
        try:
            amt = float(request.form.get("amount", 0))
        except ValueError:
            flash("Invalid amount.", "danger")
            return redirect(url_for("dashboard"))

        category    = request.form.get("category", "").strip()
        new_cat     = request.form.get("new_category", "").strip()
        if category == "__new__" and new_cat:
            category = new_cat

        subcategory = request.form.get("subcategory", "").strip()
        new_sub     = request.form.get("new_subcategory", "").strip()
        date_str    = request.form.get("date", today_str())
        pm          = request.form.get("payment_method", "Cash")
        notes       = request.form.get("notes", "").strip()
        quantity    = parse_quantity_value(request.form.get("quantity", ""))
        unit        = clean_unit_value(request.form.get("unit", ""))

        final_sub = new_sub if (subcategory == "__new__" and new_sub) else subcategory
        if not final_sub: final_sub = category
        
        if new_cat or new_sub:
            cc = data.setdefault("custom_categories", {})
            cat_key = category or "Others"
            cc.setdefault(cat_key, [])
            if new_sub and new_sub not in cc[cat_key]:
                cc[cat_key].append(new_sub)

        expense = {
            "id":             str(uuid.uuid4()),
            "amount":         round(amt, 2),
            "category":       category or "Others",
            "subcategory":    final_sub,
            "date":           date_str,
            "payment_method": pm,
            "notes":          notes,
            "quantity":       quantity,
            "unit":           unit,
            "created_at":     datetime.now().isoformat(),
        }
        data["expenses"].append(expense)
        save_data(data)
        flash(f"✓ {g.currency_symbol}{amt:.0f} added", "success")
        return redirect(url_for("dashboard"))

    return redirect(url_for("dashboard"))

    if request.method == "POST":
        try:
            amt = float(request.form.get("amount", "").replace(",", "").strip())
            if amt <= 0:
                raise ValueError
        except ValueError:
            flash("Please enter a valid positive amount.", "danger")
            return redirect(url_for("add_expense"))

        category    = request.form.get("category", "").strip()
        subcategory = request.form.get("subcategory", "").strip()
        new_sub     = request.form.get("new_subcategory", "").strip()

        # Handle new custom subcategory
        if new_sub:
            subcategory = new_sub
            cc = data.setdefault("custom_categories", {})
            cc.setdefault(category, [])
            if new_sub not in cc[category]:
                cc[category].append(new_sub)

        expense = {
            "id":             str(uuid.uuid4()),
            "amount":         round(amt, 2),
            "category":       category,
            "subcategory":    subcategory,
            "date":           request.form.get("date", today_str()),
            "payment_method": request.form.get("payment_method", "Cash"),
            "notes":          request.form.get("notes", "").strip(),
            "created_at":     datetime.now().isoformat(),
        }
        data["expenses"].append(expense)

        # Save as quick template if requested
        if request.form.get("save_as_template"):
            tname = request.form.get("template_name", subcategory).strip() or subcategory
            data.setdefault("templates", []).append({
                "id":             str(uuid.uuid4()),
                "name":           tname,
                "category":       category,
                "subcategory":    subcategory,
                "amount":         round(amt, 2),
                "payment_method": expense["payment_method"],
            })

        save_data(data)
        symbol = getattr(g, "currency_symbol", "₹")
        flash(f"✓ {symbol}{amt:.0f} added — {subcategory or category}", "success")

        if request.form.get("add_another"):
            return redirect(url_for("add_expense"))
        return redirect(url_for("dashboard"))

    return render_spendsight_template(
        "add_expense.html",
        categories=categories,
        payment_methods=data.get("payment_methods", DEFAULT_PAYMENT_METHODS),
        today=today_str(),
        category_colors=CATEGORY_COLORS,
    )

@app.route("/add_bulk", methods=["POST"])
@login_required
def add_bulk():
    data = load_data()
    req_data = request.get_json()
    
    if not req_data or "expenses" not in req_data:
        return jsonify({"error": "No data provided"}), 400
        
    new_expenses = req_data["expenses"]
    count = 0
    total_amt = 0
    
    for item in new_expenses:
        try:
            amt = float(item.get("amount", 0))
            if amt <= 0: continue

            category = item.get("category", "Others").strip() or "Others"
            subcategory = item.get("subcategory", "").strip()

            cc = data.setdefault("custom_categories", {})
            if category not in get_all_categories(data):
                cc.setdefault(category, [])
            if subcategory:
                cc.setdefault(category, [])
                if subcategory not in cc[category]:
                    cc[category].append(subcategory)
            
            expense = {
                "id":             str(uuid.uuid4()),
                "amount":         round(amt, 2),
                "category":       category,
                "subcategory":    subcategory,
                "date":           item.get("date", today_str()),
                "payment_method": item.get("payment_method", "Cash"),
                "notes":          item.get("notes", "Scanned"),
                "quantity":       parse_quantity_value(item.get("quantity")),
                "unit":           clean_unit_value(item.get("unit", "")),
                "created_at":     datetime.now().isoformat(),
            }
            data["expenses"].append(expense)
            count += 1
            total_amt += amt
        except (ValueError, TypeError):
            continue
            
    if count > 0:
        save_data(data)
        flash(f"✓ {count} expenses added successfully (Total: {g.currency_symbol}{total_amt:.0f})", "success")
        return jsonify({"success": True, "count": count}), 200
    else:
        return jsonify({"error": "No valid expenses found"}), 400

@app.route("/view")
@login_required
def view_expenses():
    data = load_data()
    expenses   = data["expenses"]
    categories = get_all_categories(data)

    date_from = request.args.get("date_from", "")
    date_to   = request.args.get("date_to",   "")
    category  = request.args.get("category",  "")
    payment   = request.args.get("payment",   "")
    search    = request.args.get("search",    "").lower()
    try:
        page     = int(request.args.get("page", 1))
        per_page = request.args.get("per_page", "50")
    except (ValueError, TypeError):
        page, per_page = 1, "50"

    filtered = expenses
    if date_from: filtered = [e for e in filtered if e["date"] >= date_from]
    if date_to:   filtered = [e for e in filtered if e["date"] <= date_to]
    if category:  filtered = [e for e in filtered if e["category"] == category]
    if payment:   filtered = [e for e in filtered if e["payment_method"] == payment]
    if search:
        filtered = [e for e in filtered if
                    search in e.get("notes", e.get("note", "")).lower() or
                    search in e.get("category", "").lower() or
                    search in e.get("subcategory", "").lower()]

    filtered = sorted(filtered, key=lambda x: (x["date"], x.get("created_at", "")), reverse=True)
    total_count  = len(filtered)
    total_amount = sum(e["amount"] for e in filtered)

    if per_page == "all":
        per_page_val = total_count if total_count > 0 else 50
    else:
        try:    per_page_val = int(per_page)
        except: per_page_val = 50

    start_idx          = (page - 1) * per_page_val
    paginated_expenses = filtered[start_idx : start_idx + per_page_val]
    total_pages        = max(1, (total_count + per_page_val - 1) // per_page_val)
    page_total         = sum(e["amount"] for e in paginated_expenses)

    bsd = int(data.get("billing_start_day", 1))
    bill_start, bill_end = get_billing_period(bsd)
    billing_cycle_total = sum(
        e["amount"] for e in expenses
        if bill_start.isoformat() <= e["date"] <= bill_end.isoformat()
    )

    return render_spendsight_template(
        "view_expenses.html",
        expenses=paginated_expenses,
        total=page_total,
        filtered_total=total_amount,
        billing_cycle_total=billing_cycle_total,
        billing_label=billing_period_label(bsd),
        total_count=total_count,
        categories=list(categories.keys()),
        payment_methods=data["payment_methods"],
        filters={"date_from": date_from, "date_to": date_to,
                 "category": category, "payment": payment, "search": search,
                 "page": page, "per_page": per_page},
        total_pages=total_pages,
        current_page=page,
    )

def filter_expense_records(expenses, date_from="", date_to="", category="", payment="", search=""):
    filtered = list(expenses)
    if date_from:
        filtered = [e for e in filtered if e["date"] >= date_from]
    if date_to:
        filtered = [e for e in filtered if e["date"] <= date_to]
    if category:
        filtered = [e for e in filtered if e["category"] == category]
    if payment:
        filtered = [e for e in filtered if e["payment_method"] == payment]
    if search:
        search_l = search.lower()
        filtered = [
            e for e in filtered
            if search_l in e.get("notes", e.get("note", "")).lower()
            or search_l in e.get("category", "").lower()
            or search_l in e.get("subcategory", "").lower()
        ]
    return filtered

@app.route("/edit-bulk/category", methods=["GET", "POST"])
@login_required
def edit_bulk_category():
    data = load_data()
    categories = get_all_categories(data)
    payment_methods = data.get("payment_methods", DEFAULT_PAYMENT_METHODS)

    date_from = request.values.get("date_from", "")
    date_to   = request.values.get("date_to", "")
    category  = request.values.get("category", "").strip()
    payment   = request.values.get("payment", "").strip()
    search    = request.values.get("search", "").strip()
    per_page  = request.values.get("per_page", "50")
    page      = request.values.get("page", "1")

    if not category:
        flash("Select a category first to bulk edit matching entries.", "danger")
        return redirect(url_for("view_expenses"))

    matching = filter_expense_records(data["expenses"], date_from, date_to, category, payment, search)
    matching = sorted(matching, key=lambda x: (x["date"], x.get("created_at", "")), reverse=True)
    if not matching:
        flash("No matching expenses found for that bulk edit.", "danger")
        return redirect(url_for("view_expenses", date_from=date_from, date_to=date_to, category=category, payment=payment, search=search, per_page=per_page, page=page))

    if request.method == "POST":
        target_category = request.form.get("target_category", "").strip()
        new_category    = request.form.get("new_category", "").strip()
        payment_mode    = request.form.get("payment_mode", "keep").strip()
        target_payment  = request.form.get("target_payment", "").strip()
        subcat_mode     = request.form.get("subcategory_mode", "keep").strip()
        target_subcat   = request.form.get("target_subcategory", "").strip()

        final_category = new_category if target_category == "__new__" and new_category else target_category
        if not final_category:
            flash("Please choose a category.", "danger")
            return redirect(url_for("edit_bulk_category", date_from=date_from, date_to=date_to, category=category, payment=payment, search=search, per_page=per_page, page=page))

        if final_category not in categories:
            data.setdefault("custom_categories", {}).setdefault(final_category, [])
            categories = get_all_categories(data)

        if subcat_mode == "replace" and target_subcat:
            cc = data.setdefault("custom_categories", {})
            cc.setdefault(final_category, [])
            if target_subcat not in cc[final_category]:
                cc[final_category].append(target_subcat)

        match_ids = {e["id"] for e in matching}
        updated = 0
        for expense in data["expenses"]:
            if expense["id"] not in match_ids:
                continue
            expense["category"] = final_category
            if payment_mode == "replace" and target_payment:
                expense["payment_method"] = target_payment
            if subcat_mode == "replace" and target_subcat:
                expense["subcategory"] = target_subcat
            updated += 1

        save_data(data)
        flash(f"✓ Updated {updated} expense{'s' if updated != 1 else ''} in bulk.", "success")
        return redirect(url_for("view_expenses", date_from=date_from, date_to=date_to, category=final_category, payment=(target_payment if payment_mode == 'replace' else payment), search=search, per_page=per_page, page=1))

    subcategories = sorted({e.get("subcategory", "").strip() for e in matching if e.get("subcategory", "").strip()})
    return render_spendsight_template(
        "edit_bulk_category.html",
        categories=categories,
        payment_methods=payment_methods,
        current_category=category,
        match_count=len(matching),
        matching_preview=matching[:12],
        subcategories=subcategories,
        filters={
            "date_from": date_from,
            "date_to": date_to,
            "category": category,
            "payment": payment,
            "search": search,
            "per_page": per_page,
            "page": page,
        },
    )

@app.route("/purchase-audit")
@login_required
def purchase_audit():
    data = load_data()
    categories = list(get_all_categories(data).keys())
    category_filter = request.args.get("category", "Groceries").strip() or "Groceries"
    category_options = ["All"] + categories
    if category_filter not in category_options:
        category_filter = "Groceries" if "Groceries" in category_options else (category_options[0] if category_options else "All")

    audit_rows, audit_summary = build_purchase_audit_rows(data.get("expenses", []), category_filter)
    return render_spendsight_template(
        "purchase_audit.html",
        category_filter=category_filter,
        category_options=category_options,
        audit_rows=audit_rows,
        audit_summary=audit_summary,
    )

@app.route("/edit/<expense_id>", methods=["GET", "POST"])
@login_required
def edit_expense(expense_id):
    data    = load_data()
    expense = next((e for e in data["expenses"] if e["id"] == expense_id), None)
    if not expense:
        flash("Expense not found.", "danger")
        return redirect(url_for("view_expenses"))

    categories = get_all_categories(data)

    if request.method == "POST":
        try:
            amt = float(request.form.get("amount", "").replace(",", "").strip())
            if amt <= 0:
                raise ValueError
        except ValueError:
            flash("Please enter a valid positive amount.", "danger")
            return redirect(url_for("edit_expense", expense_id=expense_id))

        category    = request.form.get("category", "").strip()
        subcategory = request.form.get("subcategory", "").strip()
        new_sub     = request.form.get("new_subcategory", "").strip()

        if new_sub:
            subcategory = new_sub
            cc = data.setdefault("custom_categories", {})
            cc.setdefault(category, [])
            if new_sub not in cc[category]:
                cc[category].append(new_sub)

        expense["amount"]         = round(amt, 2)
        expense["category"]       = category
        expense["subcategory"]    = subcategory
        expense["date"]           = request.form.get("date", today_str())
        expense["payment_method"] = request.form.get("payment_method", "Cash")
        expense["notes"]          = request.form.get("notes", "").strip()
        expense["quantity"]       = parse_quantity_value(request.form.get("quantity", ""))
        expense["unit"]           = clean_unit_value(request.form.get("unit", ""))

        save_data(data)
        flash("✓ Expense updated successfully.", "success")
        return redirect(url_for("view_expenses"))

    return render_spendsight_template(
        "edit_expense.html",
        expense=expense,
        categories=categories,
        payment_methods=data.get("payment_methods", DEFAULT_PAYMENT_METHODS),
        category_colors=CATEGORY_COLORS,
    )

@app.route("/delete/<expense_id>", methods=["POST"])
@login_required
def delete_expense(expense_id):
    data = load_data()
    data["expenses"] = [e for e in data["expenses"] if e["id"] != expense_id]
    save_data(data)
    flash("Expense deleted.", "info")
    return redirect(url_for("view_expenses"))

@app.route("/analytics")
@login_required
def analytics():
    data   = load_data()
    months = sorted(set(e["date"][:7] for e in data["expenses"]), reverse=True)
    if not months:
        months = [date.today().strftime("%Y-%m")]
    return render_spendsight_template(
        "analytics.html",
        months=months,
        current_month=date.today().strftime("%Y-%m"),
    )

@app.route("/api/period-stats")
@login_required
def api_period_stats():
    data   = load_data()
    period = request.args.get("period", "billing")
    today  = date.today()
    bsd    = int(data.get("billing_start_day", 1))

    if period in ("billing", "month"):
        ps, pe = get_billing_period(bsd)
    elif period in ("3months", "quarter"):
        ps, pe = get_n_billing_months_ago(3, bsd)
    elif period in ("6months", "half"):
        ps, pe = get_n_billing_months_ago(6, bsd)
    elif period == "12months":
        ps, pe = get_n_billing_months_ago(12, bsd)
    elif period == "year":
        ps = today.replace(month=1, day=1)
        pe = today
    elif period == "custom":
        try:
            ps = date.fromisoformat(request.args.get("from", today.isoformat()))
            pe = date.fromisoformat(request.args.get("to",   today.isoformat()))
        except ValueError:
            ps = pe = today
    else:
        ps, pe = get_billing_period(bsd)

    ps_s, pe_s = ps.isoformat(), pe.isoformat()
    expenses = [e for e in data["expenses"] if ps_s <= e["date"] <= pe_s]

    cat_totals = defaultdict(float)
    pm_totals  = defaultdict(float)
    daily      = defaultdict(float)
    for e in expenses:
        cat_totals[e["category"]] += e["amount"]
        pm_totals[e["payment_method"]] += e["amount"]
        daily[e["date"]] += e["amount"]

    # Include active EMIs per billing-month boundaries
    bsd_day = min(int(data.get("billing_start_day", 1)), 28)
    billing_curr = ps
    while billing_curr <= pe:
        for emi in data.get("fixed_expenses", []):
            if is_emi_active_in_month(emi, billing_curr.year, billing_curr.month):
                cat_totals[emi.get("category", "EMI / Finance")] += float(emi["amount"])
                pm_totals["EMI"] += float(emi["amount"])
        next_m = billing_curr.month + 1
        next_y = billing_curr.year
        if next_m > 12:
            next_m = 1
            next_y += 1
        next_day = min(bsd_day, monthrange(next_y, next_m)[1])
        billing_curr = date(next_y, next_m, next_day)

    total     = sum(cat_totals.values())
    days_span = max((pe - ps).days + 1, 1)

    cat_labels  = sorted(cat_totals, key=cat_totals.get, reverse=True)
    cat_amounts = [round(cat_totals[c], 2) for c in cat_labels]
    cat_colors  = [CATEGORY_COLORS.get(c, "#9CA3AF") for c in cat_labels]
    cat_pct     = [round(cat_totals[c] / total * 100, 1) if total else 0 for c in cat_labels]

    day_labels, day_amounts = [], []
    d = ps
    while d <= pe:
        day_labels.append(d.strftime("%d %b"))
        day_amounts.append(round(daily.get(d.isoformat(), 0), 2))
        d += timedelta(days=1)

    def _lbl(s, e):
        if s.month == e.month and s.year == e.year:
            return f"{s.strftime('%b')} {s.day}–{e.day}, {e.year}"
        if s.year == e.year:
            return f"{s.strftime('%b')} {s.day} – {e.strftime('%b')} {e.day}, {e.year}"
        return f"{s.strftime('%b %d, %Y')} – {e.strftime('%b %d, %Y')}"

    if period in ("billing", "month"):   label = billing_period_label(bsd)
    elif period in ("3months","quarter"): label = f"3 Months ({_lbl(ps,pe)})"
    elif period in ("6months","half"):    label = f"6 Months ({_lbl(ps,pe)})"
    elif period == "12months":            label = f"12 Months ({_lbl(ps,pe)})"
    elif period == "year":                label = f"Calendar Year {today.year} (Jan 1 – {today.strftime('%b %d')})"
    else:                                 label = f"{ps_s} to {pe_s}"

    return jsonify({
        "label":       label,
        "start":       ps_s,
        "end":         pe_s,
        "total":       round(total, 2),
        "count":       len(expenses),
        "daily_avg":   round(total / days_span, 2),
        "cat_labels":  cat_labels,
        "cat_amounts": cat_amounts,
        "cat_colors":  cat_colors,
        "cat_pct":     cat_pct,
        "pm_labels":   list(pm_totals.keys()),
        "pm_amounts":  [round(v, 2) for v in pm_totals.values()],
        "day_labels":  day_labels,
        "day_amounts": day_amounts,
    })

@app.route("/templates", methods=["GET", "POST"])
@login_required
def manage_templates():
    data = load_data()
    if request.method == "POST":
        action = request.form.get("action")
        if action == "add":
            new_t = {
                "id": str(uuid.uuid4()), "name": request.form.get("name"), "amount": float(request.form.get("amount", 0)),
                "category": request.form.get("category"), "subcategory": request.form.get("subcategory", ""),
                "payment_method": request.form.get("payment_method", "Cash"), "notes": request.form.get("notes", "")
            }
            data["templates"].append(new_t)
            save_data(data)
        elif action == "delete":
            tid = request.form.get("template_id")
            data["templates"] = [t for t in data["templates"] if t["id"] != tid]
            save_data(data)
        return redirect(url_for("manage_templates"))
    return render_spendsight_template("manage_templates.html", templates=data["templates"], 
                                   categories=get_all_categories(data), payment_methods=data["payment_methods"])

@app.route("/settings", methods=["GET", "POST"])
@login_required
def settings():
    data = load_data()
    if request.method == "POST":
        # ── Payment Methods: handle renames ───────────────────────────────────
        originals = request.form.getlist("pm_original")
        currents  = request.form.getlist("pm_current")

        for old_name, new_name in zip(originals, currents):
            old_name = old_name.strip()
            new_name = new_name.strip()
            if old_name and new_name and old_name != new_name:
                for e in data["expenses"]:
                    if e.get("payment_method") == old_name:
                        e["payment_method"] = new_name
                for t in data.get("templates", []):
                    if t.get("payment_method") == old_name:
                        t["payment_method"] = new_name
                for f in data.get("fixed_expenses", []):
                    if f.get("payment_method") == old_name:
                        f["payment_method"] = new_name

        new_pms = [p.strip() for p in currents if p.strip()]
        data["payment_methods"] = new_pms if new_pms else DEFAULT_PAYMENT_METHODS

        # ── Billing Start Day ─────────────────────────────────────────────────
        try:
            bsd = int(request.form.get("billing_start_day", 1))
            bsd = max(1, min(28, bsd))
        except (ValueError, TypeError):
            bsd = 1
        data["billing_start_day"] = bsd

        # ── Currency ──────────────────────────────────────────────────────────
        currency_code = request.form.get("currency_code", "INR")
        valid_codes   = [c["code"] for c in CURRENCIES]
        if currency_code in valid_codes:
            data["currency_code"] = currency_code

        save_data(data)
        flash("Settings saved and data updated!", "success")
        return redirect(url_for("settings"))

    bsd = int(data.get("billing_start_day", 1))
    preview_start, preview_end = get_billing_period(bsd)
    billing_preview = f"{preview_start.strftime('%b %d')} – {preview_end.strftime('%b %d, %Y')}"

    tokens = load_cloud_tokens()
    cloud_connected = {
        "gdrive":   "gdrive"   in tokens,
        "onedrive": "onedrive" in tokens,
        "dropbox":  "dropbox"  in tokens,
    }

    all_categories  = list(get_all_categories(data).keys())
    budget_limits   = data.get("budget_limits", {})

    return render_spendsight_template(
        "settings.html",
        payment_methods=data.get("payment_methods", DEFAULT_PAYMENT_METHODS),
        billing_start_day=bsd,
        billing_preview=billing_preview,
        cloud_connected=cloud_connected,
        currencies=CURRENCIES,
        all_categories=all_categories,
        budget_limits=budget_limits,
    )


@app.route("/income", methods=["GET", "POST"])
@login_required
def income():
    data = load_data()

    if request.method == "POST":
        action = request.form.get("action", "")

        if action == "set_salary":
            try:
                sal = float(request.form.get("salary", "0").replace(",", "").strip())
                sal = max(0.0, sal)
            except ValueError:
                sal = 0.0
            effective_from = request.form.get("effective_from", "").strip() or date.today().isoformat()
            history = data["income"].setdefault("salary_history", [])
            history.append({"amount": round(sal, 2), "effective_from": effective_from, "added_on": date.today().isoformat()})
            history.sort(key=lambda h: h["effective_from"])
            data["income"]["monthly_salary"] = get_salary_for_date(data["income"], date.today())
            data["income"]["salary_updated"]  = effective_from
            save_data(data)
            flash("Salary saved.", "success")

        elif action == "update_salary":
            try:
                amt = float(request.form.get("monthly_salary", "0").replace(",", ""))
                data["income"]["salary_history"].append({"amount": amt, "effective_from": today_str(), "added_on": today_str()})
                save_data(data)
            except: flash("Invalid salary amount", "danger")

        elif action == "add_extra":
            try:
                amt = float(request.form.get("amount", "0").replace(",", "").strip())
                if amt <= 0: raise ValueError
            except ValueError:
                flash("Please enter a valid amount.", "danger")
                return redirect(url_for("income"))
            ei_type = request.form.get("ei_type", "one-time")
            entry = {"id": str(uuid.uuid4()), "amount": round(amt, 2),
                     "description": request.form.get("description", "").strip() or "Extra Income",
                     "type": ei_type}
            if ei_type == "recurring":
                entry["start_date"] = request.form.get("start_date", date.today().isoformat())
                entry["end_date"]   = request.form.get("end_date", "").strip()
                entry["frequency"]  = request.form.get("frequency", "monthly")
                entry["date"]       = entry["start_date"]
            else:
                entry["date"] = request.form.get("date", date.today().isoformat())
            data["extra_income"].append(entry)
            save_data(data)
            flash("Extra income added.", "success")

        elif action == "add_emi":
            try:
                amt         = float(request.form.get("amount", "0").replace(",", "").strip())
                emi_type    = request.form.get("type", "emi")
                start_year  = int(request.form.get("start_year",  date.today().year))
                start_month = int(request.form.get("start_month", date.today().month))
                category    = request.form.get("category", "").strip()
                new_cat     = request.form.get("new_category", "").strip()
                if category == "__new__" and new_cat:
                    category = new_cat
                    data.setdefault("custom_categories", {})[category] = []
                if emi_type == "emi":
                    total_months = int(request.form.get("total_months", "0"))
                    if amt <= 0 or total_months <= 0: raise ValueError
                    frequency, day_of_month = "monthly", 1
                else:
                    total_months = 0
                    frequency    = request.form.get("frequency", "monthly")
                    day_of_month = int(request.form.get("day_of_month", "1"))
                    if amt <= 0: raise ValueError
            except (ValueError, TypeError):
                flash("Please fill in all fields correctly.", "danger")
                return redirect(url_for("income"))
            emi = {"id": str(uuid.uuid4()), "name": request.form.get("name", "").strip() or "Expense",
                   "amount": round(amt, 2), "type": emi_type, "frequency": frequency,
                   "day_of_month": day_of_month, "start_year": start_year, "start_month": start_month,
                   "total_months": total_months, "category": category or "EMI / Finance",
                   "payment_method": request.form.get("payment_method", "Cash")}
            data["fixed_expenses"].append(emi)
            save_data(data)
            flash(f"{'EMI' if emi_type == 'emi' else 'Fixed Expense'} '{emi['name']}' added.", "success")

        return redirect(url_for("income"))

    # GET
    today_dt = date.today()
    salary   = get_salary_for_date(data["income"], today_dt)
    bsd      = int(data.get("billing_start_day", 1))
    bill_start, bill_end = get_billing_period(bsd)
    extra_this_cycle = sum(
        float(ei["amount"]) for ei in data["extra_income"]
        if bill_start.isoformat() <= ei.get("date", "") <= bill_end.isoformat()
    )

    emis = []
    for emi in data["fixed_expenses"]:
        st = get_emi_status(emi, today_dt)
        emis.append({**emi, "status": st})

    def emi_sort_key(x):
        st = x["status"]
        etype = 0 if st.get("type") == "fixed" else 1
        if st.get("type") == "fixed":
            return (not st["is_active"], etype, int(x.get("start_year", 0)), int(x.get("start_month", 0)))
        ey = st["end_year"] if st["end_year"] is not None else 9999
        em = st["end_month"] if st["end_month"] is not None else 12
        return (not st["is_active"], etype, ey, em)
    emis.sort(key=emi_sort_key)

    active_emi_total = sum(
        float(e["amount"]) for e in data["fixed_expenses"]
        if is_emi_active_in_month(e, today_dt.year, today_dt.month)
    )
    extra_sorted = sorted(data["extra_income"], key=lambda x: x.get("date", ""), reverse=True)[:20]
    salary_history = sorted(data["income"].get("salary_history", []),
                             key=lambda h: h["effective_from"], reverse=True)

    return render_spendsight_template(
        "income.html",
        salary=salary,
        salary_updated=data["income"].get("salary_updated", ""),
        salary_history=salary_history,
        extra_income=extra_sorted,
        all_extra_income=data["extra_income"],
        extra_this_cycle=extra_this_cycle,
        emis=emis,
        active_emi_total=active_emi_total,
        billing_label=billing_period_label(bsd),
        today=today_dt.isoformat(),
        today_year=today_dt.year,
        today_month=today_dt.month,
        categories=list(get_all_categories(data).keys()),
        pm_list=data.get("payment_methods", DEFAULT_PAYMENT_METHODS),
    )

@app.route("/templates/add", methods=["POST"])
@login_required
def add_template():
    data = load_data()
    try:    amount = float(request.form.get("amount", 0))
    except: amount = 0.0
    template = {
        "id":             str(uuid.uuid4()),
        "name":           request.form.get("name", "").strip(),
        "category":       request.form.get("category", "").strip(),
        "subcategory":    request.form.get("subcategory", "").strip(),
        "amount":         round(amount, 2),
        "payment_method": request.form.get("payment_method", "Cash"),
    }
    data["templates"].append(template)
    save_data(data)
    flash(f"Template '{template['name']}' saved!", "success")
    return redirect(url_for("manage_templates"))


@app.route("/add-from-template/<template_id>")
@login_required
def add_from_template(template_id):
    data = load_data()
    tmpl = next((t for t in data["templates"] if t["id"] == template_id), None)
    if not tmpl:
        flash("Template not found.", "danger")
        return redirect(url_for("dashboard"))
    expense = {
        "id":             str(uuid.uuid4()),
        "amount":         tmpl["amount"],
        "category":       tmpl["category"],
        "subcategory":    tmpl["subcategory"],
        "date":           today_str(),
        "payment_method": tmpl["payment_method"],
        "notes":          f"Quick add: {tmpl['name']}",
        "created_at":     datetime.now().isoformat(),
    }
    data["expenses"].append(expense)
    save_data(data)
    flash(f"✓ {tmpl['name']} — added!", "success")
    return redirect(url_for("dashboard"))


@app.route("/templates/delete/<template_id>", methods=["POST"])
@login_required
def delete_template(template_id):
    data = load_data()
    data["templates"] = [t for t in data["templates"] if t["id"] != template_id]
    save_data(data)
    flash("Template deleted.", "info")
    return redirect(url_for("manage_templates"))


@app.route("/income/extra/delete/<eid>", methods=["POST"])
@login_required
def delete_extra_income(eid):
    data = load_data()
    data["extra_income"] = [e for e in data["extra_income"] if e["id"] != eid]
    save_data(data)
    flash("Extra income entry deleted.", "success")
    return redirect(url_for("income"))


@app.route("/income/emi/delete/<eid>", methods=["POST"])
@login_required
def delete_emi(eid):
    data = load_data()
    data["fixed_expenses"] = [e for e in data["fixed_expenses"] if e["id"] != eid]
    save_data(data)
    flash("EMI deleted.", "success")
    return redirect(url_for("income"))


@app.route("/income/emi/edit/<eid>", methods=["GET", "POST"])
@login_required
def edit_emi(eid):
    data = load_data()
    emi = next((e for e in data["fixed_expenses"] if e["id"] == eid), None)
    if not emi:
        flash("EMI not found.", "danger")
        return redirect(url_for("income"))

    if request.method == "POST":
        try:
            amt = float(request.form.get("amount", "0").replace(",", "").strip())
            emi_type = request.form.get("type", "emi")
            start_year  = int(request.form.get("start_year", date.today().year))
            start_month = int(request.form.get("start_month", date.today().month))

            category = request.form.get("category", "").strip()
            new_cat  = request.form.get("new_category", "").strip()
            if category == "__new__" and new_cat:
                category = new_cat
                cc = data.setdefault("custom_categories", {})
                if category not in cc:
                    cc[category] = []

            if emi_type == "emi":
                total_months = int(request.form.get("total_months", "0"))
                if amt <= 0 or total_months <= 0:
                    raise ValueError
                frequency    = "monthly"
                day_of_month = 1
            else:
                total_months = 0
                frequency    = request.form.get("frequency", "monthly")
                day_of_month = int(request.form.get("day_of_month", "1"))
                if amt <= 0:
                    raise ValueError

            emi["name"]          = request.form.get("name", "").strip() or "Expense"
            emi["amount"]        = round(amt, 2)
            emi["type"]          = emi_type
            emi["frequency"]     = frequency
            emi["day_of_month"]  = day_of_month
            emi["start_year"]    = start_year
            emi["start_month"]   = start_month
            emi["total_months"]  = total_months
            emi["category"]      = category or "EMI / Finance"
            emi["payment_method"] = request.form.get("payment_method", "Cash")

            save_data(data)
            flash(f"{'EMI' if emi_type == 'emi' else 'Fixed Expense'} '{emi['name']}' updated.", "success")
            return redirect(url_for("income"))

        except (ValueError, TypeError):
            flash("Please fill in all fields correctly.", "danger")
            return redirect(url_for("edit_emi", eid=eid))

    categories = list(get_all_categories(data).keys())
    pm_list    = data.get("payment_methods", DEFAULT_PAYMENT_METHODS)
    return render_spendsight_template(
        "edit_emi.html",
        emi=emi,
        categories=categories,
        pm_list=pm_list,
    )


# ── Budget helpers ────────────────────────────────────────────────────────────

def get_budget_status(data):
    """Return per-category spend vs limit for the current billing period.
    Returns list of dicts sorted by pct descending (most over-budget first).
    """
    bsd    = int(data.get("billing_start_day", 1))
    bs, be = get_billing_period(bsd)
    bs_s, be_s = bs.isoformat(), be.isoformat()
    limits = data.get("budget_limits", {})

    # Tally actual spend this billing period (regular expenses)
    spent = defaultdict(float)
    for e in data["expenses"]:
        if bs_s <= e["date"] <= be_s:
            spent[e["category"]] += e["amount"]

    # Also count active fixed expenses / EMIs
    for emi in data.get("fixed_expenses", []):
        if is_emi_active_in_month(emi, bs.year, bs.month):
            spent[emi.get("category", "EMI / Finance")] += float(emi["amount"])

    result = []
    for cat, limit in limits.items():
        if not limit or float(limit) <= 0:
            continue
        s   = round(spent.get(cat, 0), 2)
        lim = round(float(limit), 2)
        pct = round(s / lim * 100, 1) if lim else 0
        result.append({
            "category": cat,
            "spent":    s,
            "limit":    lim,
            "pct":      pct,
            "status":   "over" if pct >= 100 else ("warning" if pct >= 80 else "ok"),
            "color":    CATEGORY_COLORS.get(cat, "#9CA3AF"),
        })

    # Also add categories with spend but no limit (for reference)
    for cat, s in spent.items():
        if cat not in limits:
            result.append({
                "category": cat,
                "spent":    round(s, 2),
                "limit":    0,
                "pct":      0,
                "status":   "no-limit",
                "color":    CATEGORY_COLORS.get(cat, "#9CA3AF"),
            })

    return sorted(result, key=lambda x: (-x["pct"], -x["spent"]))


def get_smart_insights(data):
    """Generate 2-4 plain-language spending insights for the dashboard."""
    insights = []
    today     = date.today()
    bsd       = int(data.get("billing_start_day", 1))
    bs, be    = get_billing_period(bsd)
    bs_s, be_s = bs.isoformat(), be.isoformat()

    prev_bs, prev_be = get_billing_period_for_n_ago(1, bsd)
    prev_s, prev_e   = prev_bs.isoformat(), prev_be.isoformat()

    expenses      = data["expenses"]
    curr_expenses = [e for e in expenses if bs_s  <= e["date"] <= be_s]
    prev_expenses = [e for e in expenses if prev_s <= e["date"] <= prev_e]

    curr_total = sum(e["amount"] for e in curr_expenses)
    prev_total = sum(e["amount"] for e in prev_expenses)

    # 1. Billing-period pace — on track or overspending?
    days_total   = (be - bs).days + 1
    days_elapsed = max(1, (today - bs).days + 1)
    days_left    = (be - today).days
    expected_by_now = (curr_total / days_elapsed) * days_total if days_elapsed else 0

    # Get income for context
    income_data = data.get("income", {})
    salary      = get_salary_for_date(income_data, today)
    emi_total   = sum(
        float(f["amount"]) for f in data.get("fixed_expenses", [])
        if is_emi_active_in_month(f, bs.year, bs.month)
    )
    budget = salary - emi_total if salary > 0 else 0

    if budget > 0 and days_left >= 0:
        remaining = budget - curr_total
        daily_budget = budget / days_total
        daily_actual = curr_total / days_elapsed
        if daily_actual > daily_budget * 1.2:
            overspend = round((daily_actual - daily_budget) * days_total)
            insights.append({
                "type": "danger",
                "icon": "bi-exclamation-triangle-fill",
                "text": f"You're spending <strong>{fmtINR(daily_actual)}/day</strong> vs your <strong>{fmtINR(daily_budget)}/day</strong> budget — projected to overspend by <strong>{fmtINR(overspend)}</strong> this cycle.",
            })
        elif remaining > 0 and days_left > 0:
            insights.append({
                "type": "success",
                "icon": "bi-check-circle-fill",
                "text": f"On track — <strong>{fmtINR(remaining)}</strong> left for the next <strong>{days_left} day{'s' if days_left != 1 else ''}</strong> ({fmtINR(remaining / days_left)}/day).",
            })

    # 2. Biggest category jump vs last billing period
    curr_cats = defaultdict(float)
    prev_cats = defaultdict(float)
    for e in curr_expenses: curr_cats[e["category"]] += e["amount"]
    for e in prev_expenses: prev_cats[e["category"]] += e["amount"]

    biggest_jump = None
    biggest_pct  = 0
    for cat, curr_amt in curr_cats.items():
        prev_amt = prev_cats.get(cat, 0)
        if prev_amt > 0 and curr_amt > prev_amt:
            pct = (curr_amt - prev_amt) / prev_amt * 100
            if pct > biggest_pct and curr_amt > 500:
                biggest_pct  = pct
                biggest_jump = (cat, curr_amt, prev_amt, pct)

    if biggest_jump and biggest_jump[3] >= 25:
        cat, c, p, pct = biggest_jump
        insights.append({
            "type": "warning",
            "icon": "bi-arrow-up-circle-fill",
            "text": f"<strong>{cat}</strong> spending is up <strong>{pct:.0f}%</strong> vs last period ({fmtINR(p)} → {fmtINR(c)}).",
        })

    # 3. Budget limit alerts
    limits = data.get("budget_limits", {})
    over_budget = []
    near_budget = []
    for cat, limit in limits.items():
        if not limit: continue
        s   = curr_cats.get(cat, 0)
        pct = s / float(limit) * 100
        if pct >= 100:
            over_budget.append((cat, s, float(limit)))
        elif pct >= 80:
            near_budget.append((cat, s, float(limit), pct))

    for cat, s, lim in over_budget[:2]:
        insights.append({
            "type": "danger",
            "icon": "bi-slash-circle-fill",
            "text": f"<strong>{cat}</strong> is over budget — spent <strong>{fmtINR(s)}</strong> of your <strong>{fmtINR(lim)}</strong> limit.",
        })
    for cat, s, lim, pct in near_budget[:1]:
        insights.append({
            "type": "warning",
            "icon": "bi-exclamation-circle-fill",
            "text": f"<strong>{cat}</strong> is at <strong>{pct:.0f}%</strong> of its budget ({fmtINR(s)} of {fmtINR(lim)}).",
        })

    # 4. EMIs due in next 7 days
    due_soon = []
    for emi in data.get("fixed_expenses", []):
        if not is_emi_active_in_month(emi, today.year, today.month):
            continue
        day = int(emi.get("day_of_month", 1))
        try:
            due_date = date(today.year, today.month, min(day, monthrange(today.year, today.month)[1]))
        except Exception:
            continue
        if 0 <= (due_date - today).days <= 7:
            due_soon.append((emi["name"], float(emi["amount"]), due_date))

    if due_soon:
        total_due = sum(a for _, a, _ in due_soon)
        names     = ", ".join(n for n, _, _ in due_soon[:3])
        insights.append({
            "type": "info",
            "icon": "bi-calendar-event-fill",
            "text": f"<strong>{len(due_soon)} payment{'s' if len(due_soon)>1 else ''}</strong> due in the next 7 days: {names} — total <strong>{fmtINR(total_due)}</strong>.",
        })

    return insights[:4]   # cap at 4


# ── Supporting API endpoints ──────────────────────────────────────────────────

@app.route("/api/budget-status")
@login_required
def api_budget_status():
    data = load_data()
    return jsonify(get_budget_status(data))


@app.route("/api/budget-limits", methods=["POST"])
@login_required
def api_save_budget_limits():
    """Save budget limits submitted as JSON: {"category": amount, ...}"""
    data   = load_data()
    limits = request.get_json(force=True) or {}
    cleaned = {}
    for cat, val in limits.items():
        try:
            amt = float(val)
            cleaned[cat] = round(amt, 2) if amt > 0 else 0
        except (ValueError, TypeError):
            cleaned[cat] = 0
    data["budget_limits"] = cleaned
    save_data(data)
    return jsonify({"ok": True})


@app.route("/api/init-data")
@login_required
def api_init_data():
    """Used by pages that need category/payment lists without a full page load."""
    data = load_data()
    return jsonify({
        "categories":      list(get_all_categories(data).keys()),
        "payment_methods": data.get("payment_methods", DEFAULT_PAYMENT_METHODS),
    })


@app.route("/api/subcategories/<category>")
@login_required
def api_subcategories(category):
    """Return subcategory list for a given category."""
    data = load_data()
    cats = get_all_categories(data)
    return jsonify(cats.get(category, []))


@app.route("/api/billing-preview")
@login_required
def api_billing_preview():
    """Live preview of the billing period for a given start day (settings page)."""
    try:
        day = int(request.args.get("day", 1))
        day = max(1, min(28, day))
    except (ValueError, TypeError):
        day = 1
    start, end = get_billing_period(day)
    return jsonify({
        "preview": f"{start.strftime('%b %d')} – {end.strftime('%b %d, %Y')}",
        "start":   start.isoformat(),
        "end":     end.isoformat(),
    })


@app.route("/api/billing-period")
@login_required
def api_billing_period():
    """Full analytics data for the current billing period (dashboard charts)."""
    data = load_data()
    bsd  = int(data.get("billing_start_day", 1))
    bill_start, bill_end = get_billing_period(bsd)
    bs, be = bill_start.isoformat(), bill_end.isoformat()

    expenses = [e for e in data["expenses"] if bs <= e["date"] <= be]

    cat_totals = defaultdict(float)
    pm_totals  = defaultdict(float)
    daily      = defaultdict(float)
    for e in expenses:
        cat_totals[e["category"]] += e["amount"]
        pm_totals[e["payment_method"]] += e["amount"]
        daily[e["date"]] += e["amount"]

    for emi in data.get("fixed_expenses", []):
        if is_emi_active_in_month(emi, bill_start.year, bill_start.month):
            cat_totals[emi.get("category", "EMI / Finance")] += float(emi["amount"])
            pm_totals[emi.get("payment_method", "Cash")]     += float(emi["amount"])

    day_labels, day_amounts = [], []
    d = bill_start
    while d <= bill_end:
        day_labels.append(d.strftime("%d %b"))
        day_amounts.append(round(daily.get(d.isoformat(), 0), 2))
        d += timedelta(days=1)

    labels  = list(cat_totals.keys())
    amounts = [round(cat_totals[l], 2) for l in labels]
    return jsonify({
        "label":       billing_period_label(bsd),
        "start":       bs,
        "end":         be,
        "total":       round(sum(amounts), 2),
        "cat_labels":  labels,
        "cat_amounts": amounts,
        "cat_colors":  [CATEGORY_COLORS.get(l, "#9CA3AF") for l in labels],
        "pm_labels":   list(pm_totals.keys()),
        "pm_amounts":  [round(v, 2) for v in pm_totals.values()],
        "day_labels":  day_labels,
        "day_amounts": day_amounts,
    })


@app.route("/api/transactions-detail")
@login_required
def api_transactions_detail():
    """Return transactions for a period, optionally filtered by category or payment method."""
    data     = load_data()
    start_s  = request.args.get("start")
    end_s    = request.args.get("end")
    category = request.args.get("category")
    payment  = request.args.get("payment")

    if not start_s or not end_s:
        return jsonify({"expenses": [], "emis": []})

    expenses = [e for e in data["expenses"] if start_s <= e["date"] <= end_s]
    if category:
        expenses = [e for e in expenses if e["category"] == category]
    if payment:
        if payment == "EMI":
            expenses = []
        else:
            expenses = [e for e in expenses if e["payment_method"] == payment]

    matching_emis = []
    seen_emis     = set()
    try:
        ps      = date.fromisoformat(start_s)
        pe      = date.fromisoformat(end_s)
        bsd_day = min(int(data.get("billing_start_day", 1)), 28)
        cur     = ps
        while cur <= pe:
            month_key = f"{cur.year}-{cur.month:02d}"
            for emi in data.get("fixed_expenses", []):
                emi_id     = emi.get("id")
                unique_key = f"{emi_id}_{month_key}"
                if unique_key not in seen_emis and is_emi_active_in_month(emi, cur.year, cur.month):
                    cat_match = not category or emi.get("category") == category
                    pm_match  = (payment == "EMI") or (not payment)
                    if cat_match and pm_match:
                        st = get_emi_status(emi)
                        matching_emis.append({
                            "name":         emi["name"],
                            "amount":       emi["amount"],
                            "month":        month_key,
                            "is_active":    st["is_active"],
                            "paid":         st["paid"],
                            "total":        st["total"],
                            "remaining":    st["remaining"],
                            "progress_pct": st["progress_pct"],
                            "type":         st["type"],
                            "frequency":    st.get("frequency", "monthly"),
                            "day_of_month": st.get("day_of_month", 1),
                        })
                        seen_emis.add(unique_key)
            nm = cur.month + 1
            ny = cur.year
            if nm > 12:
                nm, ny = 1, ny + 1
            cur = date(ny, nm, min(bsd_day, monthrange(ny, nm)[1]))
    except Exception:
        pass

    return jsonify({
        "expenses": sorted(expenses, key=lambda x: x["date"], reverse=True),
        "emis":     matching_emis,
    })


@app.route("/api/analytics/radar")
@login_required
def api_radar():
    """Radar chart: category axes vs payment-method series."""
    data      = load_data()
    month_str = request.args.get("month", "") or date.today().strftime("%Y-%m")
    year, mon = map(int, month_str.split("-"))

    expenses = [e for e in data["expenses"] if e["date"].startswith(month_str)]
    cats     = sorted(set(e["category"] for e in expenses))
    pms      = sorted(set(e["payment_method"] for e in expenses))

    fixed_this_month = [f for f in data.get("fixed_expenses", [])
                        if is_emi_active_in_month(f, year, mon)]
    if fixed_this_month:
        for f in fixed_this_month:
            c = f.get("category", "EMI / Finance")
            if c not in cats:
                cats.append(c)
        if "EMI" not in pms:
            pms.append("EMI")
        cats.sort()
        pms.sort()

    matrix = {pm: {cat: 0.0 for cat in cats} for pm in pms}
    for e in expenses:
        if e["payment_method"] in matrix and e["category"] in matrix[e["payment_method"]]:
            matrix[e["payment_method"]][e["category"]] += e["amount"]
    for f in fixed_this_month:
        if "EMI" in matrix:
            matrix["EMI"][f.get("category", "EMI / Finance")] += float(f["amount"])

    datasets = [
        {"label": pm, "data": [round(matrix[pm][c], 2) for c in cats]}
        for pm in pms
    ]
    return jsonify({"labels": cats, "datasets": datasets})


# ── Analytics API routes ───────────────────────────────────────────────────────

@app.route("/api/analytics/monthly")
@login_required
def api_monthly():
    data = load_data()
    month_str = request.args.get("month", "") or date.today().strftime("%Y-%m")
    year, mon = map(int, month_str.split("-"))
    num_days  = monthrange(year, mon)[1]
    expenses  = [e for e in data["expenses"] if e["date"].startswith(month_str)]

    cat_totals = defaultdict(float)
    for e in expenses:
        cat_totals[e["category"]] += e["amount"]
    for emi in data.get("fixed_expenses", []):
        if is_emi_active_in_month(emi, year, mon):
            cat_totals[emi.get("category", "EMI / Finance")] += float(emi["amount"])

    sorted_cats = sorted(cat_totals.items(), key=lambda x: x[1], reverse=True)
    labels  = [item[0] for item in sorted_cats]
    amounts = [round(item[1], 2) for item in sorted_cats]
    colors  = [CATEGORY_COLORS.get(l, "#9CA3AF") for l in labels]
    total   = sum(amounts)
    return jsonify({"labels": labels, "amounts": amounts, "colors": colors, "total": total,
                    "start": f"{month_str}-01", "end": f"{month_str}-{num_days:02d}"})


@app.route("/api/analytics/daily")
@login_required
def api_daily():
    data  = load_data()
    month = request.args.get("month", "") or date.today().strftime("%Y-%m")
    expenses = [e for e in data["expenses"] if e["date"].startswith(month)]

    daily = defaultdict(float)
    for e in expenses:
        daily[e["date"]] += e["amount"]

    year, mon = map(int, month.split("-"))
    num_days   = monthrange(year, mon)[1]
    day_labels = [f"{month}-{d:02d}" for d in range(1, num_days + 1)]
    amounts    = [round(daily.get(l, 0), 2) for l in day_labels]
    display    = [str(d) for d in range(1, num_days + 1)]
    return jsonify({"labels": display, "amounts": amounts})


@app.route("/api/analytics/payment-methods")
@login_required
def api_payment_methods():
    data      = load_data()
    month_str = request.args.get("month", "") or date.today().strftime("%Y-%m")
    year, mon = map(int, month_str.split("-"))
    num_days  = monthrange(year, mon)[1]
    expenses  = [e for e in data["expenses"] if e["date"].startswith(month_str)]

    pm = defaultdict(float)
    for e in expenses:
        pm[e["payment_method"]] += e["amount"]
    for emi in data.get("fixed_expenses", []):
        if is_emi_active_in_month(emi, year, mon):
            pm["EMI"] += float(emi["amount"])

    sorted_pm = sorted(pm.items(), key=lambda x: x[1], reverse=True)
    return jsonify({"labels":  [r[0] for r in sorted_pm],
                    "amounts": [round(r[1], 2) for r in sorted_pm],
                    "start":   f"{month_str}-01", "end": f"{month_str}-{num_days:02d}"})


@app.route("/api/analytics/comparison")
@login_required
def api_comparison():
    data       = load_data()
    today      = date.today()
    curr_month = today.strftime("%Y-%m")
    prev_month = (today.replace(day=1) - timedelta(days=1)).strftime("%Y-%m")

    def cat_totals(month):
        t = defaultdict(float)
        for e in data["expenses"]:
            if e["date"].startswith(month):
                t[e["category"]] += e["amount"]
        return t

    curr     = cat_totals(curr_month)
    prev     = cat_totals(prev_month)
    all_cats = sorted(set(list(curr.keys()) + list(prev.keys())))
    return jsonify({
        "categories":     all_cats,
        "current":        [round(curr.get(c, 0), 2) for c in all_cats],
        "previous":       [round(prev.get(c, 0), 2) for c in all_cats],
        "current_label":  today.strftime("%B"),
        "previous_label": (today.replace(day=1) - timedelta(days=1)).strftime("%B"),
        "colors":         [CATEGORY_COLORS.get(c, "#9CA3AF") for c in all_cats],
    })


@app.route("/api/analytics/subcategory")
@login_required
def api_subcategory():
    data     = load_data()
    month    = request.args.get("month", date.today().strftime("%Y-%m"))
    expenses = [e for e in data["expenses"] if e["date"].startswith(month)]

    freq   = defaultdict(int)
    totals = defaultdict(float)
    for e in expenses:
        key = f"{e['category']} › {e['subcategory']}"
        freq[key]   += 1
        totals[key] += e["amount"]

    rows = sorted(freq.items(), key=lambda x: x[1], reverse=True)[:15]
    return jsonify({"labels":  [r[0] for r in rows],
                    "counts":  [freq[r[0]] for r in rows],
                    "amounts": [round(totals[r[0]], 2) for r in rows]})


@app.route("/api/budget-summary")
@login_required
def api_budget_summary():
    data  = load_data()
    today = date.today()
    bsd   = int(data.get("billing_start_day", 1))

    salary = get_salary_for_date(data["income"], today)
    bill_start, bill_end = get_billing_period(bsd)

    # Extra income this billing cycle
    extra_this_cycle = sum(
        float(ei["amount"])
        for ei in data["extra_income"]
        if bill_start.isoformat() <= ei.get("date", "") <= bill_end.isoformat()
    )
    total_income = salary + extra_this_cycle

    # Active EMIs this billing month
    active_emis = [
        e for e in data["fixed_expenses"]
        if is_emi_active_in_month(e, today.year, today.month)
    ]
    emi_total = sum(float(e["amount"]) for e in active_emis)

    # Non-cash spending in current billing period
    bs_s, be_s = bill_start.isoformat(), bill_end.isoformat()
    cycle_expenses = [e for e in data["expenses"] if bs_s <= e["date"] <= be_s]
    billing_spent = sum(
        e["amount"] for e in cycle_expenses
        if (e.get("payment_method") or "").strip().lower() != "cash"
    )
    cash_spent = sum(
        e["amount"] for e in cycle_expenses
        if (e.get("payment_method") or "").strip().lower() == "cash"
    )

    available   = max(total_income - emi_total, 0)
    remaining   = available - billing_spent
    net_savings = total_income - emi_total - billing_spent

    # 12-month savings trend
    trend_labels, trend_income, trend_fixed, trend_spent, trend_net = [], [], [], [], []
    for n in range(11, -1, -1):
        ps, pe = get_billing_period_for_n_ago(n, bsd, today)
        ps_s, pe_s = ps.isoformat(), pe.isoformat()
        period_salary = get_salary_for_date(data["income"], ps)
        period_extra  = sum(
            float(ei["amount"])
            for ei in data["extra_income"]
            if ps_s <= ei.get("date", "") <= pe_s
        )
        period_income = period_salary + period_extra
        period_emi    = sum(
            float(e["amount"])
            for e in data["fixed_expenses"]
            if is_emi_active_in_month(e, ps.year, ps.month)
        )
        period_spent  = sum(
            e["amount"] for e in data["expenses"]
            if ps_s <= e["date"] <= pe_s and (e.get("payment_method") or "").strip().lower() != "cash"
        )
        period_net = period_income - period_emi - period_spent
        lbl = f"{ps.strftime('%b')} '{str(ps.year)[2:]}"
        trend_labels.append(lbl)
        trend_income.append(round(period_income, 2))
        trend_fixed.append(round(period_emi, 2))
        trend_spent.append(round(period_spent, 2))
        trend_net.append(round(period_net, 2))

    # Future EMI closure projection (next 12 months)
    future_labels, future_savings = [], []
    for i in range(12):
        fm = today.month + i
        fy = today.year + (fm - 1) // 12
        fm = (fm - 1) % 12 + 1
        fut_emi = sum(
            float(e["amount"])
            for e in data["fixed_expenses"]
            if is_emi_active_in_month(e, fy, fm)
        )
        fut_net = (get_salary_for_date(data["income"], date(fy, fm, 1)) or salary) - fut_emi
        lbl = date(fy, fm, 1).strftime("%b '%y")
        future_labels.append(lbl)
        future_savings.append(round(fut_net, 2))

    return jsonify({
        "salary":         round(salary, 2),
        "extra_income":   round(extra_this_cycle, 2),
        "total_income":   round(total_income, 2),
        "emi_total":      round(emi_total, 2),
        "billing_spent":  round(billing_spent, 2),
        "cash_spent":     round(cash_spent, 2),
        "available":      round(available, 2),
        "remaining":      round(remaining, 2),
        "net_savings":    round(net_savings, 2),
        "billing_label":  billing_period_label(bsd),
        "has_salary":     salary > 0,
        "trend_labels":   trend_labels,
        "trend_income":   trend_income,
        "trend_fixed":    trend_fixed,
        "trend_spent":    trend_spent,
        "trend_net":      trend_net,
        "future_labels":  future_labels,
        "future_savings": future_savings,
    })

@app.route("/api/spending-forecast")
@login_required
def api_spending_forecast():
    data     = load_data()
    today_dt = date.today()
    bsd      = int(data.get("billing_start_day", 1))
    day      = min(bsd, 28)

    salary = get_salary_for_date(data["income"], today_dt)
    if salary <= 0:
        return jsonify({"has_salary": False})

    # Average monthly non-cash spending over last 3 billing cycles
    past_totals = []
    cur_start, cur_end = get_billing_period(bsd)
    ref = cur_start
    for _ in range(3):
        m = ref.month - 1 or 12
        y = ref.year if ref.month > 1 else ref.year - 1
        max_d   = monthrange(y, m)[1]
        p_start = date(y, m, min(day, max_d))
        p_end   = ref - timedelta(days=1)
        total   = sum(
            e["amount"] for e in data["expenses"]
            if p_start.isoformat() <= e["date"] <= p_end.isoformat()
            and (e.get("payment_method") or "").strip().lower() != "cash"
        )
        past_totals.append(total)
        ref = p_start

    avg_spend = round(sum(past_totals) / len(past_totals), 2) if past_totals else 0

    # Build next 3 billing periods
    months = []
    ref_start = cur_start
    for i in range(3):
        m = ref_start.month % 12 + 1
        y = ref_start.year + (1 if ref_start.month == 12 else 0)
        max_d   = monthrange(y, m)[1]
        f_start = date(y, m, min(day, max_d))
        m2 = f_start.month % 12 + 1
        y2 = f_start.year + (1 if f_start.month == 12 else 0)
        max_d2  = monthrange(y2, m2)[1]
        f_end   = date(y2, m2, min(day, max_d2)) - timedelta(days=1)

        active_emis = []
        for emi in data.get("fixed_expenses", []):
            if is_emi_active_in_month(emi, f_start.year, f_start.month):
                active_emis.append({
                    "name": emi.get("name", "Unnamed"),
                    "amount": float(emi["amount"])
                })

        emi_total = sum(e["amount"] for e in active_emis)
        extra = get_month_extra_income(data.get("extra_income", []), f_start.year, f_start.month)
        total_income  = round(salary + extra, 2)
        available     = round(total_income - emi_total, 2)
        projected_bal = round(available - avg_spend, 2)

        months.append({
            "label":         f_start.strftime("%B %Y"),
            "period":        f"{f_start.strftime('%b %d')} – {f_end.strftime('%b %d')}",
            "salary":        round(salary, 2),
            "extra_income":  round(extra, 2),
            "total_income":  total_income,
            "emi_total":     round(emi_total, 2),
            "emis":          active_emis,
            "available":     available,
            "avg_spend":     avg_spend,
            "projected_bal": projected_bal,
            "status":        "safe" if projected_bal >= 0 else "risk",
        })
        ref_start = f_start

    return jsonify({
        "has_salary": True,
        "avg_spend":  avg_spend,
        "months":     months,
    })


@app.route("/export/csv")
@login_required
def export_csv():
    data = load_data()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Date", "Category", "Amount"])
    for e in data["expenses"]: writer.writerow([e["date"], e["category"], e["amount"]])
    return Response(output.getvalue(), mimetype="text/csv", headers={"Content-Disposition": "attachment; filename=expenses.csv"})


@app.route("/export/pdf")
@login_required
def export_pdf():
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib.units import cm
    except ImportError:
        flash("PDF export requires reportlab. Run: pip install reportlab", "danger")
        return redirect(url_for("view_expenses"))

    data     = load_data()
    symbol   = getattr(g, "currency_symbol", "₹")
    expenses = sorted(data["expenses"], key=lambda x: x["date"], reverse=True)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles   = getSampleStyleSheet()
    elements = []

    title = Paragraph("<b>SpendSight — Expense Report</b>", styles["Title"])
    elements.append(title)
    elements.append(Spacer(1, 0.4*cm))
    sub = Paragraph(
        f"Generated: {date.today().strftime('%B %d, %Y')}  ·  Total records: {len(expenses)}",
        styles["Normal"])
    elements.append(sub)
    elements.append(Spacer(1, 0.6*cm))

    header = ["Date", "Category", "Subcategory", f"Amount ({symbol})", "Payment", "Note"]
    rows   = [header]
    for e in expenses:
        rows.append([
            e.get("date", ""),
            e.get("category", ""),
            e.get("subcategory", ""),
            f"{symbol}{e.get('amount', 0):.0f}",
            e.get("payment_method", ""),
            (e.get("notes", e.get("note", "")) or "")[:40],
        ])

    col_widths = [2.2*cm, 3*cm, 3*cm, 2.5*cm, 2.8*cm, None]
    t = Table(rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, 0), colors.HexColor("#1A56DB")),
        ("TEXTCOLOR",    (0, 0), (-1, 0), colors.white),
        ("FONTNAME",     (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",     (0, 0), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f1f5f9")]),
        ("GRID",         (0, 0), (-1, -1), 0.4, colors.HexColor("#e2e8f0")),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING",   (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
    ]))
    elements.append(t)

    doc.build(elements)
    buf.seek(0)
    return Response(
        buf.getvalue(),
        mimetype="application/pdf",
        headers={"Content-Disposition": "attachment; filename=spendsight_expenses.pdf"}
    )


# ── Cloud Backup Routes ────────────────────────────────────────────────────────

@app.route("/cloud/connect/<service>")
@login_required
def cloud_connect(service):
    tokens = load_cloud_tokens()
    state  = secrets.token_urlsafe(16)
    tokens[f"{service}_state"] = state
    save_cloud_tokens(tokens)

    if service == "gdrive":
        params = {
            "client_id":     CLOUD_CREDENTIALS["gdrive"]["client_id"],
            "redirect_uri":  f"{REDIRECT_BASE}/cloud/callback/gdrive",
            "response_type": "code",
            "scope":         "https://www.googleapis.com/auth/drive.file",
            "access_type":   "offline",
            "prompt":        "consent",
            "state":         state,
        }
        return redirect("https://accounts.google.com/o/oauth2/v2/auth?" + urllib.parse.urlencode(params))

    elif service == "onedrive":
        params = {
            "client_id":     CLOUD_CREDENTIALS["onedrive"]["client_id"],
            "redirect_uri":  f"{REDIRECT_BASE}/cloud/callback/onedrive",
            "response_type": "code",
            "scope":         "Files.ReadWrite offline_access",
            "state":         state,
        }
        return redirect("https://login.microsoftonline.com/common/oauth2/v2.0/authorize?" + urllib.parse.urlencode(params))

    elif service == "dropbox":
        params = {
            "client_id":         CLOUD_CREDENTIALS["dropbox"]["app_key"],
            "redirect_uri":      f"{REDIRECT_BASE}/cloud/callback/dropbox",
            "response_type":     "code",
            "token_access_type": "offline",
            "state":             state,
        }
        return redirect("https://www.dropbox.com/oauth2/authorize?" + urllib.parse.urlencode(params))

    flash("Unknown service.", "danger")
    return redirect(url_for("settings"))


@app.route("/cloud/callback/<service>")
@login_required
def cloud_callback(service):
    code  = request.args.get("code")
    state = request.args.get("state")
    error = request.args.get("error")

    if error:
        flash(f"Authorization denied: {error}", "danger")
        return redirect(url_for("settings") + "#cloud-backup")

    tokens      = load_cloud_tokens()
    saved_state = tokens.get(f"{service}_state", "")
    if state != saved_state:
        flash("Invalid state — possible CSRF. Please try again.", "danger")
        return redirect(url_for("settings") + "#cloud-backup")

    try:
        if service == "gdrive":
            cid     = CLOUD_CREDENTIALS["gdrive"]["client_id"]
            csecret = CLOUD_CREDENTIALS["gdrive"]["client_secret"]
            post_data = urllib.parse.urlencode({
                "code":          code,
                "client_id":     cid,
                "client_secret": csecret,
                "redirect_uri":  f"{REDIRECT_BASE}/cloud/callback/gdrive",
                "grant_type":    "authorization_code",
            }).encode()
            req = urllib.request.Request("https://oauth2.googleapis.com/token",
                                         data=post_data, method="POST")
            with urllib.request.urlopen(req) as r:
                tk = json.loads(r.read())
            tokens["gdrive"] = tk
            tokens["gdrive_file_id"] = tokens.get("gdrive_file_id", "")

        elif service == "onedrive":
            post_data = urllib.parse.urlencode({
                "code":          code,
                "client_id":     CLOUD_CREDENTIALS["onedrive"]["client_id"],
                "redirect_uri":  f"{REDIRECT_BASE}/cloud/callback/onedrive",
                "grant_type":    "authorization_code",
                "scope":         "Files.ReadWrite offline_access",
            }).encode()
            req = urllib.request.Request(
                "https://login.microsoftonline.com/common/oauth2/v2.0/token",
                data=post_data, method="POST")
            with urllib.request.urlopen(req) as r:
                tk = json.loads(r.read())
            tokens["onedrive"] = tk

        elif service == "dropbox":
            key    = CLOUD_CREDENTIALS["dropbox"]["app_key"]
            secret = CLOUD_CREDENTIALS["dropbox"]["app_secret"]
            post_data = urllib.parse.urlencode({
                "code":         code,
                "grant_type":   "authorization_code",
                "redirect_uri": f"{REDIRECT_BASE}/cloud/callback/dropbox",
            }).encode()
            import base64
            auth = base64.b64encode(f"{key}:{secret}".encode()).decode()
            req = urllib.request.Request(
                "https://api.dropboxapi.com/oauth2/token",
                data=post_data, method="POST",
                headers={"Authorization": f"Basic {auth}"})
            with urllib.request.urlopen(req) as r:
                tk = json.loads(r.read())
            tokens["dropbox"] = tk

        save_cloud_tokens(tokens)
        svc_names = {"gdrive": "Google Drive", "onedrive": "OneDrive", "dropbox": "Dropbox"}
        flash(f"✓ Connected to {svc_names.get(service, service)} successfully!", "success")

    except Exception as e:
        flash(f"Connection failed: {e}", "danger")

    return redirect(url_for("settings") + "#cloud-backup")


@app.route("/cloud/backup/<service>", methods=["POST"])
@login_required
def cloud_backup(service):
    tokens    = load_cloud_tokens()
    svc_token = tokens.get(service)
    if not svc_token:
        flash(f"Not connected to {service}. Connect first.", "danger")
        return redirect(url_for("settings") + "#cloud-backup")

    try:
        data_file = get_current_data_file()
        with open(data_file, "rb") as f:
            file_bytes = f.read()

        if service == "gdrive":
            access_token = svc_token.get("access_token", "")
            file_id      = tokens.get("gdrive_file_id", "")
            if file_id:
                req = urllib.request.Request(
                    f"https://www.googleapis.com/upload/drive/v3/files/{file_id}?uploadType=media",
                    data=file_bytes, method="PATCH",
                    headers={"Authorization": f"Bearer {access_token}",
                             "Content-Type": "application/json"})
            else:
                import json as _json
                meta     = _json.dumps({"name": "spendsight_backup.json",
                                        "mimeType": "application/json"}).encode()
                boundary = b"boundary_spendsight"
                body = (b"--" + boundary + b"\r\n"
                        b"Content-Type: application/json; charset=UTF-8\r\n\r\n" +
                        meta + b"\r\n"
                        b"--" + boundary + b"\r\n"
                        b"Content-Type: application/json\r\n\r\n" +
                        file_bytes + b"\r\n"
                        b"--" + boundary + b"--")
                req = urllib.request.Request(
                    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
                    data=body, method="POST",
                    headers={"Authorization": f"Bearer {access_token}",
                             "Content-Type": f"multipart/related; boundary={boundary.decode()}"})
            with urllib.request.urlopen(req) as r:
                resp = json.loads(r.read())
            if not file_id:
                tokens["gdrive_file_id"] = resp.get("id", "")
                save_cloud_tokens(tokens)

        elif service == "onedrive":
            access_token = svc_token.get("access_token", "")
            req = urllib.request.Request(
                "https://graph.microsoft.com/v1.0/me/drive/root:/spendsight_backup.json:/content",
                data=file_bytes, method="PUT",
                headers={"Authorization": f"Bearer {access_token}",
                         "Content-Type": "application/json"})
            with urllib.request.urlopen(req) as r:
                r.read()

        elif service == "dropbox":
            access_token = svc_token.get("access_token", "")
            import json as _json
            arg = _json.dumps({"path": "/spendsight_backup.json", "mode": "overwrite"})
            req = urllib.request.Request(
                "https://content.dropboxapi.com/2/files/upload",
                data=file_bytes, method="POST",
                headers={"Authorization": f"Bearer {access_token}",
                         "Content-Type": "application/octet-stream",
                         "Dropbox-API-Arg": arg})
            with urllib.request.urlopen(req) as r:
                r.read()

        svc_names = {"gdrive": "Google Drive", "onedrive": "OneDrive", "dropbox": "Dropbox"}
        flash(f"✓ Backed up to {svc_names.get(service, service)} successfully!", "success")

    except Exception as e:
        flash(f"Backup failed: {e}", "danger")

    return redirect(url_for("settings") + "#cloud-backup")


@app.route("/cloud/restore/<service>", methods=["POST"])
@login_required
def cloud_restore(service):
    tokens    = load_cloud_tokens()
    svc_token = tokens.get(service)
    if not svc_token:
        flash(f"Not connected to {service}.", "danger")
        return redirect(url_for("settings") + "#cloud-backup")

    try:
        access_token = svc_token.get("access_token", "")

        if service == "gdrive":
            file_id = tokens.get("gdrive_file_id", "")
            if not file_id:
                flash("No backup file found on Google Drive.", "warning")
                return redirect(url_for("settings") + "#cloud-backup")
            req = urllib.request.Request(
                f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media",
                headers={"Authorization": f"Bearer {access_token}"})
            with urllib.request.urlopen(req) as r:
                content = r.read()

        elif service == "onedrive":
            req = urllib.request.Request(
                "https://graph.microsoft.com/v1.0/me/drive/root:/spendsight_backup.json:/content",
                headers={"Authorization": f"Bearer {access_token}"})
            with urllib.request.urlopen(req) as r:
                content = r.read()

        elif service == "dropbox":
            import json as _json
            arg = _json.dumps({"path": "/spendsight_backup.json"})
            req = urllib.request.Request(
                "https://content.dropboxapi.com/2/files/download",
                method="POST",
                headers={"Authorization": f"Bearer {access_token}",
                         "Dropbox-API-Arg": arg,
                         "Content-Type": ""})
            with urllib.request.urlopen(req) as r:
                content = r.read()

        # Validate JSON before overwriting
        json.loads(content)
        data_file = get_current_data_file()
        with open(data_file, "wb") as f:
            f.write(content)

        svc_names = {"gdrive": "Google Drive", "onedrive": "OneDrive", "dropbox": "Dropbox"}
        flash(f"✓ Restored from {svc_names.get(service, service)} successfully!", "success")

    except Exception as e:
        flash(f"Restore failed: {e}", "danger")

    return redirect(url_for("settings") + "#cloud-backup")


@app.route("/cloud/disconnect/<service>", methods=["POST"])
@login_required
def cloud_disconnect(service):
    tokens = load_cloud_tokens()
    tokens.pop(service, None)
    if service == "gdrive":
        tokens.pop("gdrive_file_id", None)
    save_cloud_tokens(tokens)
    svc_names = {"gdrive": "Google Drive", "onedrive": "OneDrive", "dropbox": "Dropbox"}
    flash(f"Disconnected from {svc_names.get(service, service)}.", "info")
    return redirect(url_for("settings") + "#cloud-backup")


@app.route("/api/cloud/status")
@login_required
def api_cloud_status():
    tokens = load_cloud_tokens()
    return jsonify({
        "gdrive":   "gdrive" in tokens,
        "onedrive": "onedrive" in tokens,
        "dropbox":  "dropbox" in tokens,
    })


# ── SpendBot chat endpoint ────────────────────────────────────────────────────

@app.route("/api/chat", methods=["POST"])
@login_required
def api_chat():
    """
    Rule-based SpendBot.  No external AI APIs.

    Request  JSON: {"message": "How much on milk this month?"}
    Response JSON: {"reply": "...", "total": 68.0, "item": "milk",
                    "timeframe": "March 2026"}

    Stateful "Did you mean?" flow
    ------------------------------
    When SpendBot returns match_type="suggestion" this route stores the
    pending state in the Flask session.  If the user replies "yes" (or "y",
    "yeah", "yep") on the next turn, we use the stored state to run the
    confirmed item query directly.
    """
    from spendbot import SpendBot
    from datetime import date as _date

    body         = request.get_json(silent=True) or {}
    user_message = (body.get("message") or "").strip()

    if not user_message:
        return jsonify({
            "reply":     "Please type a question, e.g. 'How much on milk this month?'",
            "total":     None, "item": None, "timeframe": None,
        })

    data = load_data()

    # Resolve currency symbol for formatting inside SpendBot
    currency_code   = data.get("settings", {}).get("currency_code", "INR")
    currency_symbol = next(
        (c["symbol"] for c in CURRENCIES if c["code"] == currency_code), "₹"
    )

    bot = SpendBot(
        expenses           = data.get("expenses", []),
        custom_categories  = data.get("custom_categories", {}),
        recurring_payments = data.get("fixed_expenses", []),
        currency_symbol    = currency_symbol,
    )

    # ── "yes" confirmation for a pending "Did you mean?" suggestion ───────
    _YES = {"yes", "y", "yeah", "yep", "sure", "ok", "okay"}
    pending = session.get("sb_pending")
    if user_message.lower() in _YES and pending:
        session.pop("sb_pending", None)
        try:
            start      = _date.fromisoformat(pending["start"])
            end        = _date.fromisoformat(pending["end"])
            tf_label   = pending["tf_label"]
            tf_defaulted = pending["tf_defaulted"]
            item       = pending["item"]
            result     = bot._item_query(item, start, end, tf_label, tf_defaulted)
            result["reply"] = f"Got it! {result['reply']}"
            result["match_type"] = "confirmed"
            return jsonify(result)
        except (KeyError, ValueError):
            pass   # fall through to normal processing if session data is malformed

    # ── Enhancement 1: Timeframe follow-up ────────────────────────────────
    from datetime import datetime as _datetime
    
    # Context management (10-minute timeout)
    session_context = session.get("sb_context", {})
    if session_context:
        try:
            ctx_ts = _datetime.fromisoformat(session_context.get("timestamp", ""))
            if (_datetime.now() - ctx_ts).total_seconds() > 600:
                session_context = {}
        except: session_context = {}

    pending_tf = session.get("sb_pending_tf")
    if pending_tf:
        try:
            ts = _datetime.fromisoformat(pending_tf["timestamp"])
            if (_datetime.now() - ts).total_seconds() < 300:
                tf_res = bot._detect_timeframe(user_message)
                if "all time" in user_message.lower():
                    tf_res = (_date(2000, 1, 1), _date.today(), "all time", False)
                if tf_res and not tf_res[3]:
                    session.pop("sb_pending_tf", None)
                    start, end, label, defaulted = tf_res
                    item = pending_tf["item"]
                    res = bot._item_query(item, start, end, label, defaulted)
                    # Sync context after TF resolution
                    session_context["last_queried_item"] = item
                    session_context["timestamp"] = _datetime.now().isoformat()
                    session["sb_context"] = session_context
                    return jsonify(res)
            else:
                session.pop("sb_pending_tf", None)
        except (KeyError, ValueError):
            session.pop("sb_pending_tf", None)

    # ── Normal SpendBot processing ────────────────────────────────────────
    result = bot.reply(user_message, context=session_context)

    # Update context
    new_ctx = {
        "timestamp": _datetime.now().isoformat(),
        "last_queried_item": result.get("last_queried_item") or session_context.get("last_queried_item"),
        "last_displayed_rank": result.get("last_displayed_rank") or session_context.get("last_displayed_rank"),
        "context_sorted_expenses": result.get("context_sorted_expenses") or session_context.get("context_sorted_expenses"),
        "context_timeframe": result.get("context_timeframe") or session_context.get("context_timeframe")
    }
    session["sb_context"] = new_ctx

    # If bot is asking "Did you mean X?", persist the pending suggestion so
    # the next "yes" can resolve it without re-parsing the original message.
    if result.get("match_type") == "suggestion" and result.get("suggested_item"):
        session["sb_pending"] = {
            "item":        result["suggested_item"],
            "start":       result.get("_start", ""),
            "end":         result.get("_end", ""),
            "tf_label":    result.get("_tf_label", ""),
            "tf_defaulted": result.get("_tf_defaulted", True),
        }
    elif result.get("needs_timeframe"):
        session["sb_pending_tf"] = {
            "item":      result.get("pending_item"),
            "timestamp": _datetime.now().isoformat(),
        }
        session.pop("sb_pending", None)
    else:
        session.pop("sb_pending", None)   # clear stale state on any other reply
        session.pop("sb_pending_tf", None)

    # Strip internal fields before sending to client
    for k in ("_start", "_end", "_tf_label", "_tf_defaulted", "suggested_item", "pending_item", "needs_timeframe", "last_queried_item", "last_displayed_rank", "context_sorted_expenses", "context_timeframe"):
        result.pop(k, None)

    return jsonify(result)


def open_browser():
    webbrowser.open("http://localhost:5000")

if __name__ == "__main__":
    print("\n" + "="*50)
    print("  SpendSight is starting...")
    print("  Open: http://localhost:5000")
    print("  Press Ctrl+C to stop")
    print("="*50 + "\n")
    threading.Timer(1.2, open_browser).start()
    app.run(debug=False, port=5000, use_reloader=False)
