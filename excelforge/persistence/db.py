from __future__ import annotations

import sqlite3
from pathlib import Path

from excelforge.config import AppConfig


class Database:
    def __init__(self, config: AppConfig) -> None:
        self._path = config.sqlite_path
        self._path.parent.mkdir(parents=True, exist_ok=True)

    @property
    def path(self) -> Path:
        return self._path

    def connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self._path)
        conn.row_factory = sqlite3.Row
        return conn

    def init_schema(self) -> None:
        with self.connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS audit_operation (
                    operation_id TEXT PRIMARY KEY,
                    tool_name TEXT NOT NULL,
                    workbook_id TEXT,
                    file_path TEXT,
                    actor_id TEXT,
                    os_user TEXT,
                    machine_name TEXT,
                    client_name TEXT,
                    client_request_id TEXT,
                    started_at TEXT NOT NULL,
                    duration_ms INTEGER NOT NULL,
                    success INTEGER NOT NULL,
                    code TEXT NOT NULL,
                    message TEXT NOT NULL,
                    affected_sheet TEXT,
                    affected_range TEXT,
                    snapshot_id TEXT,
                    args_summary TEXT
                );

                CREATE TABLE IF NOT EXISTS snapshot_meta (
                    snapshot_id TEXT PRIMARY KEY,
                    workbook_id TEXT NOT NULL,
                    file_path TEXT NOT NULL,
                    sheet_name TEXT NOT NULL,
                    range_address TEXT NOT NULL,
                    source_tool TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    cell_count INTEGER NOT NULL,
                    file_path_snapshot TEXT NOT NULL,
                    expired INTEGER NOT NULL DEFAULT 0,
                    file_size_bytes INTEGER NOT NULL DEFAULT 0,
                    expired_reason TEXT,
                    cleaned_at TEXT
                );

                CREATE TABLE IF NOT EXISTS preview_token (
                    token TEXT PRIMARY KEY,
                    snapshot_id TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    expires_at TEXT NOT NULL,
                    used INTEGER NOT NULL DEFAULT 0
                );

                CREATE TABLE IF NOT EXISTS backup_meta (
                    backup_id TEXT PRIMARY KEY,
                    workbook_id TEXT NOT NULL,
                    file_path TEXT NOT NULL,
                    backup_file_path TEXT NOT NULL,
                    file_size_bytes INTEGER NOT NULL,
                    source_tool TEXT NOT NULL,
                    source_operation_id TEXT,
                    description TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    expired INTEGER NOT NULL DEFAULT 0,
                    expired_reason TEXT,
                    cleaned_at TEXT
                );

                CREATE INDEX IF NOT EXISTS idx_audit_started_at ON audit_operation(started_at);
                CREATE INDEX IF NOT EXISTS idx_snapshot_workbook ON snapshot_meta(workbook_id);
                CREATE INDEX IF NOT EXISTS idx_snapshot_expired ON snapshot_meta(expired);
                CREATE INDEX IF NOT EXISTS idx_preview_snapshot ON preview_token(snapshot_id);
                CREATE INDEX IF NOT EXISTS idx_backup_workbook ON backup_meta(workbook_id);
                CREATE INDEX IF NOT EXISTS idx_backup_expired ON backup_meta(expired);
                """
            )
            self._ensure_column(conn, "snapshot_meta", "file_size_bytes", "INTEGER NOT NULL DEFAULT 0")
            self._ensure_column(conn, "snapshot_meta", "expired_reason", "TEXT")
            self._ensure_column(conn, "snapshot_meta", "cleaned_at", "TEXT")
            self._ensure_column(conn, "backup_meta", "expired_reason", "TEXT")
            self._ensure_column(conn, "backup_meta", "cleaned_at", "TEXT")

    @staticmethod
    def _ensure_column(conn: sqlite3.Connection, table: str, column: str, ddl: str) -> None:
        rows = conn.execute(f"PRAGMA table_info({table})").fetchall()
        existing = {str(row[1]) for row in rows}
        if column in existing:
            return
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {ddl}")
