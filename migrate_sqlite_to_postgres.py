#!/usr/bin/env python
"""
DataWise AI Stage 7.12 — SQLite to PostgreSQL metadata migration.

This utility copies application metadata only:
  app_users, dataset_history, report_history, activity_log, app_settings

It does not copy uploaded source files or generated report binaries.
It does not delete or modify the SQLite source database.

Usage:
  pip install "psycopg[binary]>=3.2,<4"
  # Set DATABASE_URL in your environment (never commit it).
  python migrate_sqlite_to_postgres.py --sqlite datawise_app.db
  python migrate_sqlite_to_postgres.py --sqlite datawise_app.db --verify-only
"""
import argparse
import os
import sqlite3
import sys

try:
    import psycopg
    from psycopg.rows import dict_row
except ImportError:
    raise SystemExit("Missing driver. Install with: python -m pip install 'psycopg[binary]>=3.2,<4'")

TABLES = {
    "app_users": """
        CREATE TABLE IF NOT EXISTS app_users (
            id BIGINT PRIMARY KEY, email TEXT NOT NULL UNIQUE,
            display_name TEXT NOT NULL, password_hash TEXT NOT NULL,
            password_salt TEXT NOT NULL, role TEXT NOT NULL,
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL, last_login_at TEXT
        )""",
    "dataset_history": """
        CREATE TABLE IF NOT EXISTS dataset_history (
            id BIGINT PRIMARY KEY, dataset_signature TEXT NOT NULL UNIQUE,
            session_id TEXT NOT NULL, owner_user_id BIGINT,
            source_name TEXT NOT NULL, source_type TEXT NOT NULL,
            row_count BIGINT NOT NULL, column_count BIGINT NOT NULL,
            columns_json TEXT NOT NULL, uploaded_at TEXT NOT NULL
        )""",
    "report_history": """
        CREATE TABLE IF NOT EXISTS report_history (
            id BIGINT PRIMARY KEY, session_id TEXT NOT NULL,
            owner_user_id BIGINT, dataset_signature TEXT,
            report_title TEXT NOT NULL, source_name TEXT NOT NULL,
            formats_json TEXT NOT NULL, row_count BIGINT NOT NULL,
            column_count BIGINT NOT NULL, created_at TEXT NOT NULL
        )""",
    "activity_log": """
        CREATE TABLE IF NOT EXISTS activity_log (
            id BIGINT PRIMARY KEY, session_id TEXT NOT NULL,
            event_type TEXT NOT NULL, description TEXT NOT NULL,
            created_at TEXT NOT NULL
        )""",
    "app_settings": """
        CREATE TABLE IF NOT EXISTS app_settings (
            setting_key TEXT PRIMARY KEY, setting_value TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )""",
}
COLUMNS = {
    "app_users": ["id","email","display_name","password_hash","password_salt","role","is_active","created_at","last_login_at"],
    "dataset_history": ["id","dataset_signature","session_id","owner_user_id","source_name","source_type","row_count","column_count","columns_json","uploaded_at"],
    "report_history": ["id","session_id","owner_user_id","dataset_signature","report_title","source_name","formats_json","row_count","column_count","created_at"],
    "activity_log": ["id","session_id","event_type","description","created_at"],
    "app_settings": ["setting_key","setting_value","updated_at"],
}

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--sqlite", required=True, help="Path to existing DataWise SQLite database")
    parser.add_argument("--verify-only", action="store_true", help="Only compare source/target row counts")
    args = parser.parse_args()
    db_url = os.environ.get("DATABASE_URL")
    if not db_url:
        raise SystemExit("DATABASE_URL is not set. Configure it as an environment variable; do not put credentials in this script.")
    if not os.path.isfile(args.sqlite):
        raise SystemExit(f"SQLite file not found: {args.sqlite}")

    source = sqlite3.connect(f"file:{os.path.abspath(args.sqlite)}?mode=ro", uri=True)
    source.row_factory = sqlite3.Row
    with psycopg.connect(db_url, connect_timeout=10) as target:
        if not args.verify_only:
            for ddl in TABLES.values():
                target.execute(ddl)
            target.commit()

        print("DataWise AI migration verification")
        all_ok = True
        for table, cols in COLUMNS.items():
            source_count = source.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
            target_count = target.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
            if not args.verify_only and source_count:
                existing = {r[0] for r in target.execute(f"SELECT {cols[0]} FROM {table}").fetchall()}
                rows = source.execute(f"SELECT {', '.join(cols)} FROM {table}").fetchall()
                pending = [tuple(row[c] for c in cols) for row in rows if row[0] not in existing]
                if pending:
                    placeholders = ", ".join(["%s"] * len(cols))
                    col_sql = ", ".join(cols)
                    target.cursor().executemany(
                        f"INSERT INTO {table} ({col_sql}) VALUES ({placeholders})",
                        pending
                    )
                    target.commit()
                target_count = target.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
            ok = target_count >= source_count
            all_ok = all_ok and ok
            print(f"{table:18} SQLite={source_count:6} PostgreSQL={target_count:6} {'OK' if ok else 'CHECK'}")
        source.close()
        if not all_ok:
            raise SystemExit("Verification failed: one or more PostgreSQL tables contain fewer rows than SQLite.")
        print("Migration check completed. SQLite source remains unchanged.")

if __name__ == "__main__":
    main()
