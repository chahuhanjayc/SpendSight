"""
SQLite persistence adapter for SpendSight.

This module is intentionally independent of Flask so it can be used by
one-off migration scripts, tests, and a future storage switch in app.py.
The adapter preserves the current JSON document contract while writing
queryable relational rows for the core finance entities.
"""

from __future__ import annotations

import hashlib
import json
import os
import sqlite3
from contextlib import contextmanager
from datetime import datetime


SCHEMA_VERSION = 1


ROW_TABLE_KEYS = {
    "expenses": "expenses",
    "templates": "templates",
    "extra_income": "extra_income",
    "fixed_expenses": "fixed_expenses",
    "goals": "goals",
    "transaction_rules": "transaction_rules",
    "accounts": "accounts",
    "net_worth_snapshots": "net_worth_snapshots",
    "budget_envelopes": "budget_envelopes",
    "receipts": "receipts",
    "recurring_payments": "recurring_payments",
}


def _json_dumps(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), default=str)


def _json_loads(value, fallback):
    if not value:
        return fallback
    try:
        return json.loads(value)
    except json.JSONDecodeError:
        return fallback


def _connect(db_path):
    directory = os.path.dirname(os.path.abspath(db_path))
    if directory:
        os.makedirs(directory, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=5000")
    return conn


@contextmanager
def _managed_connection(db_path):
    conn = _connect(db_path)
    try:
        yield conn
    finally:
        conn.close()


def _table_columns(conn, table):
    return {row["name"] for row in conn.execute(f"PRAGMA table_info({table})")}


def _as_bool(value):
    return 1 if bool(value) else 0


def _row_id(item, prefix, index):
    value = str(item.get("id") or "").strip()
    return value or f"{prefix}-{index}"


def init_db(db_path):
    conn = _connect(db_path)
    try:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS schema_meta (
              key TEXT PRIMARY KEY,
              value TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS users (
              id TEXT PRIMARY KEY,
              password_hash TEXT NOT NULL,
              role TEXT NOT NULL DEFAULT 'user',
              created_at TEXT,
              updated_at TEXT
            );

            CREATE TABLE IF NOT EXISTS user_settings (
              user_id TEXT PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
              billing_start_day INTEGER NOT NULL DEFAULT 1,
              currency_code TEXT NOT NULL DEFAULT 'INR',
              app_schema_version INTEGER NOT NULL DEFAULT 3,
              monthly_salary REAL NOT NULL DEFAULT 0,
              salary_updated TEXT,
              raw_json TEXT
            );

            CREATE TABLE IF NOT EXISTS salary_history (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              amount REAL NOT NULL,
              effective_from TEXT NOT NULL,
              added_on TEXT,
              raw_json TEXT
            );

            CREATE TABLE IF NOT EXISTS payment_methods (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              name TEXT NOT NULL,
              sort_order INTEGER NOT NULL DEFAULT 0,
              PRIMARY KEY(user_id, name)
            );

            CREATE TABLE IF NOT EXISTS custom_categories (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              category TEXT NOT NULL,
              subcategory TEXT NOT NULL,
              sort_order INTEGER NOT NULL DEFAULT 0,
              PRIMARY KEY(user_id, category, subcategory)
            );

            CREATE TABLE IF NOT EXISTS budget_limits (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              category TEXT NOT NULL,
              amount REAL NOT NULL DEFAULT 0,
              PRIMARY KEY(user_id, category)
            );

            CREATE TABLE IF NOT EXISTS expenses (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              id TEXT NOT NULL,
              amount REAL NOT NULL DEFAULT 0,
              category TEXT,
              subcategory TEXT,
              date TEXT,
              payment_method TEXT,
              account_id TEXT,
              transfer_group_id TEXT,
              notes TEXT,
              quantity REAL,
              unit TEXT,
              source TEXT,
              review_status TEXT,
              receipt_id TEXT,
              created_at TEXT,
              updated_at TEXT,
              reviewed_at TEXT,
              raw_json TEXT,
              PRIMARY KEY(user_id, id)
            );

            CREATE INDEX IF NOT EXISTS idx_expenses_user_date ON expenses(user_id, date);
            CREATE INDEX IF NOT EXISTS idx_expenses_user_category_date ON expenses(user_id, category, date);
            CREATE INDEX IF NOT EXISTS idx_expenses_user_payment_date ON expenses(user_id, payment_method, date);

            CREATE TABLE IF NOT EXISTS templates (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              id TEXT NOT NULL,
              name TEXT,
              amount REAL,
              category TEXT,
              subcategory TEXT,
              payment_method TEXT,
              notes TEXT,
              raw_json TEXT,
              PRIMARY KEY(user_id, id)
            );

            CREATE TABLE IF NOT EXISTS extra_income (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              id TEXT NOT NULL,
              amount REAL,
              description TEXT,
              type TEXT,
              date TEXT,
              start_date TEXT,
              end_date TEXT,
              frequency TEXT,
              raw_json TEXT,
              PRIMARY KEY(user_id, id)
            );

            CREATE TABLE IF NOT EXISTS fixed_expenses (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              id TEXT NOT NULL,
              name TEXT,
              amount REAL,
              type TEXT,
              frequency TEXT,
              day_of_month INTEGER,
              start_year INTEGER,
              start_month INTEGER,
              total_months INTEGER,
              category TEXT,
              payment_method TEXT,
              raw_json TEXT,
              PRIMARY KEY(user_id, id)
            );

            CREATE TABLE IF NOT EXISTS goals (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              id TEXT NOT NULL,
              name TEXT,
              target_amount REAL,
              current_amount REAL,
              monthly_contribution REAL,
              priority TEXT,
              notes TEXT,
              raw_json TEXT,
              PRIMARY KEY(user_id, id)
            );

            CREATE TABLE IF NOT EXISTS transaction_rules (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              id TEXT NOT NULL,
              name TEXT,
              match_text TEXT,
              min_amount REAL,
              max_amount REAL,
              set_category TEXT,
              set_subcategory TEXT,
              set_payment_method TEXT,
              enabled INTEGER,
              created_at TEXT,
              raw_json TEXT,
              PRIMARY KEY(user_id, id)
            );

            CREATE TABLE IF NOT EXISTS accounts (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              id TEXT NOT NULL,
              name TEXT,
              type TEXT,
              institution TEXT,
              balance REAL,
              currency_code TEXT,
              include_in_net_worth INTEGER,
              is_archived INTEGER,
              notes TEXT,
              created_at TEXT,
              updated_at TEXT,
              raw_json TEXT,
              PRIMARY KEY(user_id, id)
            );

            CREATE TABLE IF NOT EXISTS net_worth_snapshots (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              id TEXT NOT NULL,
              date TEXT,
              assets REAL,
              liabilities REAL,
              net_worth REAL,
              created_at TEXT,
              raw_json TEXT,
              PRIMARY KEY(user_id, id)
            );

            CREATE TABLE IF NOT EXISTS budget_envelopes (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              id TEXT NOT NULL,
              month TEXT,
              category TEXT,
              assigned REAL,
              annual_amount REAL,
              rollover_enabled INTEGER,
              notes TEXT,
              created_at TEXT,
              updated_at TEXT,
              raw_json TEXT,
              PRIMARY KEY(user_id, id)
            );

            CREATE INDEX IF NOT EXISTS idx_budget_envelopes_user_month_category
              ON budget_envelopes(user_id, month, category);

            CREATE TABLE IF NOT EXISTS receipts (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              id TEXT NOT NULL,
              original_filename TEXT,
              stored_filename TEXT,
              content_type TEXT,
              uploaded_at TEXT,
              status TEXT,
              expense_id TEXT,
              extracted_json TEXT,
              raw_json TEXT,
              PRIMARY KEY(user_id, id)
            );

            CREATE INDEX IF NOT EXISTS idx_receipts_user_expense
              ON receipts(user_id, expense_id);

            CREATE TABLE IF NOT EXISTS recurring_payments (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              id TEXT NOT NULL,
              payload_json TEXT,
              raw_json TEXT,
              PRIMARY KEY(user_id, id)
            );

            CREATE TABLE IF NOT EXISTS cloud_token_items (
              user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
              key TEXT NOT NULL,
              value_json TEXT,
              updated_at TEXT NOT NULL,
              PRIMARY KEY(user_id, key)
            );

            CREATE TABLE IF NOT EXISTS migration_runs (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              source_path TEXT NOT NULL,
              source_hash TEXT NOT NULL,
              migrated_at TEXT NOT NULL,
              status TEXT NOT NULL,
              message TEXT
            );
            """
        )
        conn.execute(
            "INSERT OR REPLACE INTO schema_meta(key, value) VALUES('schema_version', ?)",
            (str(SCHEMA_VERSION),),
        )
        expense_columns = _table_columns(conn, "expenses")
        for column, definition in {
            "account_id": "TEXT",
            "transfer_group_id": "TEXT",
        }.items():
            if column not in expense_columns:
                conn.execute(f"ALTER TABLE expenses ADD COLUMN {column} {definition}")
        conn.commit()
    finally:
        conn.close()


class SQLiteStore:
    def __init__(self, db_path):
        self.db_path = db_path
        init_db(db_path)

    def connection(self):
        return _managed_connection(self.db_path)

    def load_users(self):
        with self.connection() as conn:
            rows = conn.execute("SELECT id, password_hash, role FROM users ORDER BY id").fetchall()
        return [{"id": row["id"], "password": row["password_hash"], "role": row["role"]} for row in rows]

    def save_users(self, users):
        now = datetime.now().isoformat()
        incoming_ids = [user["id"] for user in users]
        with self.connection() as conn:
            for user in users:
                conn.execute(
                    """
                    INSERT INTO users(id, password_hash, role, created_at, updated_at)
                    VALUES(?, ?, ?, ?, ?)
                    ON CONFLICT(id) DO UPDATE SET
                      password_hash=excluded.password_hash,
                      role=excluded.role,
                      updated_at=excluded.updated_at
                    """,
                    (
                        user["id"],
                        user.get("password", ""),
                        user.get("role", "user"),
                        user.get("created_at", now),
                        now,
                    ),
                )
            if incoming_ids:
                placeholders = ",".join("?" for _ in incoming_ids)
                conn.execute(f"DELETE FROM users WHERE id NOT IN ({placeholders})", incoming_ids)
            else:
                conn.execute("DELETE FROM users")
            conn.commit()

    def ensure_user(self, user_id):
        now = datetime.now().isoformat()
        with self.connection() as conn:
            conn.execute(
                """
                INSERT OR IGNORE INTO users(id, password_hash, role, created_at, updated_at)
                VALUES(?, '', 'user', ?, ?)
                """,
                (user_id, now, now),
            )
            conn.commit()

    def load_user_data(self, user_id):
        with self.connection() as conn:
            settings = conn.execute("SELECT * FROM user_settings WHERE user_id=?", (user_id,)).fetchone()
            data = _json_loads(settings["raw_json"], {}) if settings else {}
            data.update(
                {
                    "expenses": self._load_collection(conn, user_id, "expenses"),
                    "templates": self._load_collection(conn, user_id, "templates"),
                    "extra_income": self._load_collection(conn, user_id, "extra_income"),
                    "fixed_expenses": self._load_collection(conn, user_id, "fixed_expenses"),
                    "goals": self._load_collection(conn, user_id, "goals"),
                    "transaction_rules": self._load_collection(conn, user_id, "transaction_rules"),
                    "accounts": self._load_collection(conn, user_id, "accounts"),
                    "net_worth_snapshots": self._load_collection(conn, user_id, "net_worth_snapshots"),
                    "budget_envelopes": self._load_collection(conn, user_id, "budget_envelopes"),
                    "receipts": self._load_collection(conn, user_id, "receipts"),
                    "recurring_payments": self._load_collection(conn, user_id, "recurring_payments"),
                    "payment_methods": [row["name"] for row in conn.execute("SELECT name FROM payment_methods WHERE user_id=? ORDER BY sort_order, name", (user_id,))],
                    "custom_categories": self._load_custom_categories(conn, user_id),
                    "budget_limits": {
                        row["category"]: row["amount"]
                        for row in conn.execute("SELECT category, amount FROM budget_limits WHERE user_id=?", (user_id,))
                    },
                }
            )
            if settings:
                data["billing_start_day"] = settings["billing_start_day"]
                data["currency_code"] = settings["currency_code"]
                data["schema_version"] = settings["app_schema_version"]
                income = data.setdefault("income", {})
                income["monthly_salary"] = settings["monthly_salary"]
                if settings["salary_updated"]:
                    income["salary_updated"] = settings["salary_updated"]
                income["salary_history"] = [
                    _json_loads(row["raw_json"], {"amount": row["amount"], "effective_from": row["effective_from"], "added_on": row["added_on"]})
                    for row in conn.execute("SELECT * FROM salary_history WHERE user_id=? ORDER BY effective_from", (user_id,))
                ]
        return data

    def save_user_data(self, user_id, data):
        self.ensure_user(user_id)
        with self.connection() as conn:
            self._delete_user_rows(conn, user_id)
            income = data.get("income", {})
            conn.execute(
                """
                INSERT OR REPLACE INTO user_settings(
                  user_id, billing_start_day, currency_code, app_schema_version,
                  monthly_salary, salary_updated, raw_json
                ) VALUES(?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    user_id,
                    int(data.get("billing_start_day", 1) or 1),
                    data.get("currency_code", "INR"),
                    int(data.get("schema_version", 3) or 3),
                    float(income.get("monthly_salary", 0) or 0),
                    income.get("salary_updated", ""),
                    _json_dumps({k: v for k, v in data.items() if k not in ROW_TABLE_KEYS}),
                ),
            )
            for index, item in enumerate(income.get("salary_history", []) or []):
                conn.execute(
                    """
                    INSERT INTO salary_history(user_id, amount, effective_from, added_on, raw_json)
                    VALUES(?, ?, ?, ?, ?)
                    """,
                    (
                        user_id,
                        float(item.get("amount", 0) or 0),
                        item.get("effective_from", ""),
                        item.get("added_on", ""),
                        _json_dumps(item),
                    ),
                )
            for index, name in enumerate(data.get("payment_methods", []) or []):
                conn.execute(
                    "INSERT OR REPLACE INTO payment_methods(user_id, name, sort_order) VALUES(?, ?, ?)",
                    (user_id, str(name), index),
                )
            for category, subcategories in (data.get("custom_categories", {}) or {}).items():
                for index, subcategory in enumerate(subcategories or []):
                    conn.execute(
                        """
                        INSERT OR REPLACE INTO custom_categories(user_id, category, subcategory, sort_order)
                        VALUES(?, ?, ?, ?)
                        """,
                        (user_id, category, subcategory, index),
                    )
            for category, amount in (data.get("budget_limits", {}) or {}).items():
                conn.execute(
                    "INSERT OR REPLACE INTO budget_limits(user_id, category, amount) VALUES(?, ?, ?)",
                    (user_id, category, float(amount or 0)),
                )
            for key, table in ROW_TABLE_KEYS.items():
                for index, item in enumerate(data.get(key, []) or []):
                    self._insert_collection_row(conn, user_id, table, item, index)
            conn.commit()

    def export_user_data(self, user_id):
        return self.load_user_data(user_id)

    def restore_user_data(self, user_id, data):
        self.save_user_data(user_id, data)

    def load_cloud_tokens(self, user_id):
        with self.connection() as conn:
            rows = conn.execute("SELECT key, value_json FROM cloud_token_items WHERE user_id=?", (user_id,)).fetchall()
        return {row["key"]: _json_loads(row["value_json"], None) for row in rows}

    def save_cloud_tokens(self, user_id, tokens):
        self.ensure_user(user_id)
        now = datetime.now().isoformat()
        with self.connection() as conn:
            conn.execute("DELETE FROM cloud_token_items WHERE user_id=?", (user_id,))
            for key, value in (tokens or {}).items():
                conn.execute(
                    "INSERT INTO cloud_token_items(user_id, key, value_json, updated_at) VALUES(?, ?, ?, ?)",
                    (user_id, key, _json_dumps(value), now),
                )
            conn.commit()

    def record_migration(self, source_path, source_hash, status, message=""):
        with self.connection() as conn:
            conn.execute(
                """
                INSERT INTO migration_runs(source_path, source_hash, migrated_at, status, message)
                VALUES(?, ?, ?, ?, ?)
                """,
                (source_path, source_hash, datetime.now().isoformat(), status, message),
            )
            conn.commit()

    def _delete_user_rows(self, conn, user_id):
        for table in [
            "user_settings",
            "salary_history",
            "payment_methods",
            "custom_categories",
            "budget_limits",
            *ROW_TABLE_KEYS.values(),
        ]:
            conn.execute(f"DELETE FROM {table} WHERE user_id=?", (user_id,))

    def _load_collection(self, conn, user_id, table):
        rows = conn.execute(f"SELECT raw_json FROM {table} WHERE user_id=? ORDER BY rowid", (user_id,)).fetchall()
        return [_json_loads(row["raw_json"], {}) for row in rows]

    def _load_custom_categories(self, conn, user_id):
        categories = {}
        rows = conn.execute(
            "SELECT category, subcategory FROM custom_categories WHERE user_id=? ORDER BY category, sort_order, subcategory",
            (user_id,),
        ).fetchall()
        for row in rows:
            categories.setdefault(row["category"], []).append(row["subcategory"])
        return categories

    def _insert_collection_row(self, conn, user_id, table, item, index):
        if not isinstance(item, dict):
            item = {"value": item}
        item_id = _row_id(item, table, index)
        raw = _json_dumps({**item, "id": item_id} if "id" not in item else item)
        columns = _table_columns(conn, table)
        values = {"user_id": user_id, "id": item_id, "raw_json": raw}

        for key in columns:
            if key in {"user_id", "id", "raw_json"}:
                continue
            if key == "payload_json":
                values[key] = _json_dumps(item)
            elif key == "extracted_json":
                values[key] = _json_dumps(item.get("extracted", {}))
            elif key in {"enabled", "include_in_net_worth", "is_archived", "rollover_enabled"}:
                values[key] = _as_bool(item.get(key, True))
            elif key in item:
                values[key] = item.get(key)

        insert_columns = [column for column in values.keys() if column in columns]
        placeholders = ",".join("?" for _ in insert_columns)
        conn.execute(
            f"INSERT OR REPLACE INTO {table}({','.join(insert_columns)}) VALUES({placeholders})",
            [values[column] for column in insert_columns],
        )


def file_sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def migrate_legacy_json(workspace_dir, db_path):
    store = SQLiteStore(db_path)
    users_path = os.path.join(workspace_dir, "users.json")
    if os.path.exists(users_path):
        with open(users_path, "r", encoding="utf-8") as handle:
            users = json.load(handle)
        store.save_users(users)
        store.record_migration(users_path, file_sha256(users_path), "imported")

    data_dir = os.path.join(workspace_dir, "user_data")
    if not os.path.isdir(data_dir):
        return store

    for name in sorted(os.listdir(data_dir)):
        path = os.path.join(data_dir, name)
        if not os.path.isfile(path):
            continue
        if name.startswith("data_") and name.endswith(".json"):
            user_id = name[len("data_") : -len(".json")]
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
            store.save_user_data(user_id, data)
            store.record_migration(path, file_sha256(path), "imported")
        elif name.startswith("cloud_tokens_") and name.endswith(".json"):
            user_id = name[len("cloud_tokens_") : -len(".json")]
            with open(path, "r", encoding="utf-8") as handle:
                tokens = json.load(handle)
            store.save_cloud_tokens(user_id, tokens)
            store.record_migration(path, file_sha256(path), "imported")
    return store
