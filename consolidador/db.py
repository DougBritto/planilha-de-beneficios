from __future__ import annotations

import sqlite3

from flask import Flask, current_app, g

from .config import AppSettings


UPLOAD_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS uploads (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    plan_type TEXT,
    submission_id TEXT NOT NULL,
    submitted_at TEXT NOT NULL,
    sender_name TEXT,
    sender_email TEXT,
    note TEXT,
    original_name TEXT NOT NULL,
    stored_name TEXT,
    stored_path TEXT NOT NULL,
    size_bytes INTEGER NOT NULL,
    file_hash TEXT,
    content_type TEXT,
    remote_addr TEXT,
    user_agent TEXT
)
"""

CONSOLIDATION_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS consolidations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    plan_type TEXT,
    created_at TEXT NOT NULL,
    output_name TEXT NOT NULL,
    stored_path TEXT NOT NULL,
    scope_submission_id TEXT,
    sheet_name TEXT,
    remove_duplicates INTEGER NOT NULL DEFAULT 0,
    ignore_empty INTEGER NOT NULL DEFAULT 1,
    total_files INTEGER NOT NULL DEFAULT 0,
    valid_files INTEGER NOT NULL DEFAULT 0,
    invalid_files INTEGER NOT NULL DEFAULT 0,
    rows_generated INTEGER NOT NULL DEFAULT 0,
    requested_by TEXT
)
"""

AUDIT_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS audit_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at TEXT NOT NULL,
    event_type TEXT NOT NULL,
    actor TEXT,
    remote_addr TEXT,
    details TEXT,
    submission_id TEXT,
    file_name TEXT
)
"""


def ensure_data_dirs(settings: AppSettings) -> None:
    settings.data_dir.mkdir(parents=True, exist_ok=True)
    settings.upload_dir.mkdir(parents=True, exist_ok=True)
    settings.output_dir.mkdir(parents=True, exist_ok=True)


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        settings: AppSettings = current_app.config["APP_SETTINGS"]
        settings.db_path.parent.mkdir(parents=True, exist_ok=True)
        connection = sqlite3.connect(settings.db_path)
        connection.row_factory = sqlite3.Row
        g.db = connection
    return g.db


def close_db(_: Exception | None = None) -> None:
    connection = g.pop("db", None)
    if connection is not None:
        connection.close()


def _ensure_columns(conn: sqlite3.Connection, table_name: str, column_sql: dict[str, str]) -> None:
    existing = {
        row["name"]
        for row in conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    }
    for name, definition in column_sql.items():
        if name not in existing:
            conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {name} {definition}")


def initialize_database() -> None:
    conn = get_db()
    conn.execute(UPLOAD_TABLE_SQL)
    conn.execute(CONSOLIDATION_TABLE_SQL)
    conn.execute(AUDIT_TABLE_SQL)
    _ensure_columns(
        conn,
        "uploads",
        {
            "plan_type": "TEXT",
            "stored_name": "TEXT",
            "file_hash": "TEXT",
            "content_type": "TEXT",
            "remote_addr": "TEXT",
            "user_agent": "TEXT",
        },
    )
    _ensure_columns(
        conn,
        "consolidations",
        {
            "plan_type": "TEXT",
            "requested_by": "TEXT",
            "scope_submission_id": "TEXT",
            "sheet_name": "TEXT",
            "remove_duplicates": "INTEGER NOT NULL DEFAULT 0",
            "ignore_empty": "INTEGER NOT NULL DEFAULT 1",
            "total_files": "INTEGER NOT NULL DEFAULT 0",
            "valid_files": "INTEGER NOT NULL DEFAULT 0",
            "invalid_files": "INTEGER NOT NULL DEFAULT 0",
            "rows_generated": "INTEGER NOT NULL DEFAULT 0",
        },
    )
    conn.commit()


def init_app(app: Flask) -> None:
    settings: AppSettings = app.config["APP_SETTINGS"]
    ensure_data_dirs(settings)
    app.teardown_appcontext(close_db)
    with app.app_context():
        initialize_database()


def fetch_one(query: str, params: tuple[object, ...] = ()) -> sqlite3.Row | None:
    return get_db().execute(query, params).fetchone()


def fetch_all(query: str, params: tuple[object, ...] = ()) -> list[sqlite3.Row]:
    return get_db().execute(query, params).fetchall()


def execute(query: str, params: tuple[object, ...] = ()) -> int:
    conn = get_db()
    cursor = conn.execute(query, params)
    conn.commit()
    return int(cursor.lastrowid)
