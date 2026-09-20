"""Lehká SQLite persistence pro fázi 3."""

import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

from werkzeug.security import check_password_hash, generate_password_hash


SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'viewer' CHECK(role IN ('admin','manager','viewer')),
    active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS time_entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date TEXT NOT NULL,
    start_time TEXT,
    end_time TEXT,
    lunch_duration REAL NOT NULL DEFAULT 0,
    is_free_day INTEGER NOT NULL DEFAULT 0,
    employee_names TEXT NOT NULL,
    source TEXT NOT NULL DEFAULT 'excel',
    synced_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS audit_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp TEXT NOT NULL,
    action TEXT NOT NULL,
    actor TEXT NOT NULL,
    details TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS sync_state (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
"""


class Database:
    def __init__(self, path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.initialize()

    @contextmanager
    def connection(self):
        connection = sqlite3.connect(self.path)
        connection.row_factory = sqlite3.Row
        try:
            yield connection
            connection.commit()
        finally:
            connection.close()

    def initialize(self):
        with self.connection() as connection:
            connection.executescript(SCHEMA)

    def user_count(self):
        with self.connection() as connection:
            return connection.execute("SELECT COUNT(*) FROM users").fetchone()[0]

    def create_user(self, username, password, role="viewer"):
        username = username.strip().lower()
        if len(username) < 3 or len(password) < 8:
            raise ValueError("Uživatelské jméno musí mít alespoň 3 znaky a heslo alespoň 8 znaků.")
        if role not in {"admin", "manager", "viewer"}:
            raise ValueError("Neplatná role uživatele.")
        try:
            with self.connection() as connection:
                cursor = connection.execute(
                    "INSERT INTO users(username,password_hash,role,created_at) VALUES(?,?,?,?)",
                    (username, generate_password_hash(password), role, datetime.now().astimezone().isoformat()),
                )
                return cursor.lastrowid
        except sqlite3.IntegrityError as error:
            raise ValueError("Uživatel již existuje.") from error

    def authenticate(self, username, password):
        with self.connection() as connection:
            user = connection.execute(
                "SELECT * FROM users WHERE username=? AND active=1", (username.strip().lower(),)
            ).fetchone()
        if user and check_password_hash(user["password_hash"], password):
            return dict(user)
        return None

    def get_user(self, user_id):
        with self.connection() as connection:
            user = connection.execute(
                "SELECT id,username,role,active,created_at FROM users WHERE id=? AND active=1", (user_id,)
            ).fetchone()
        return dict(user) if user else None

    def list_users(self):
        with self.connection() as connection:
            return [
                dict(row)
                for row in connection.execute("SELECT id,username,role,active,created_at FROM users ORDER BY username")
            ]

    def set_user_active(self, user_id, active):
        with self.connection() as connection:
            connection.execute("UPDATE users SET active=? WHERE id=?", (int(active), user_id))

    def record_time_entry(
        self, entry_date, start_time, end_time, lunch_duration, is_free_day, employees, source="excel"
    ):
        with self.connection() as connection:
            connection.execute(
                "INSERT INTO time_entries(entry_date,start_time,end_time,lunch_duration,"
                "is_free_day,employee_names,source,synced_at) VALUES(?,?,?,?,?,?,?,?)",
                (
                    entry_date,
                    start_time,
                    end_time,
                    float(lunch_duration or 0),
                    int(bool(is_free_day)),
                    ", ".join(employees),
                    source,
                    datetime.now().astimezone().isoformat(),
                ),
            )

    def set_sync_state(self, key, value):
        with self.connection() as connection:
            connection.execute(
                "INSERT INTO sync_state(key,value,updated_at) VALUES(?,?,?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at",
                (key, value, datetime.now().astimezone().isoformat()),
            )

    def get_sync_state(self):
        with self.connection() as connection:
            return {
                row["key"]: {"value": row["value"], "updated_at": row["updated_at"]}
                for row in connection.execute("SELECT * FROM sync_state")
            }

    def record_audit(self, action, actor, details):
        with self.connection() as connection:
            connection.execute(
                "INSERT INTO audit_events(timestamp,action,actor,details) VALUES(?,?,?,?)",
                (datetime.now().astimezone().isoformat(), action, actor, details),
            )

    def list_audit(self, limit=100):
        with self.connection() as connection:
            return [
                dict(row) for row in connection.execute("SELECT * FROM audit_events ORDER BY id DESC LIMIT ?", (limit,))
            ]
