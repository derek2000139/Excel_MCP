from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from excelforge.persistence.db import Database


@dataclass
class BackupMetaRecord:
    backup_id: str
    workbook_id: str
    file_path: str
    backup_file_path: str
    file_size_bytes: int
    source_tool: str
    source_operation_id: str | None
    description: str
    created_at: str
    expired: bool = False
    expired_reason: str | None = None
    cleaned_at: str | None = None


class BackupRepository:
    def __init__(self, db: Database) -> None:
        self._db = db

    def insert_meta(self, record: BackupMetaRecord) -> None:
        with self._db.connect() as conn:
            conn.execute(
                """
                INSERT INTO backup_meta (
                    backup_id, workbook_id, file_path, backup_file_path, file_size_bytes,
                    source_tool, source_operation_id, description, created_at, expired,
                    expired_reason, cleaned_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    record.backup_id,
                    record.workbook_id,
                    record.file_path,
                    record.backup_file_path,
                    int(record.file_size_bytes),
                    record.source_tool,
                    record.source_operation_id,
                    record.description,
                    record.created_at,
                    1 if record.expired else 0,
                    record.expired_reason,
                    record.cleaned_at,
                ),
            )

    def list_backups(
        self,
        *,
        workbook_id: str | None,
        file_path: str | None,
        limit: int,
        offset: int,
    ) -> tuple[int, list[dict[str, object]]]:
        conditions: list[str] = []
        params: list[object] = []
        if workbook_id:
            conditions.append("workbook_id = ?")
            params.append(workbook_id)
        if file_path:
            conditions.append("file_path = ?")
            params.append(file_path)
        where = f"WHERE {' AND '.join(conditions)}" if conditions else ""

        with self._db.connect() as conn:
            total_row = conn.execute(
                f"SELECT COUNT(1) AS c FROM backup_meta {where}",
                params,
            ).fetchone()
            rows = conn.execute(
                f"""
                SELECT backup_id, workbook_id, file_path, backup_file_path, file_size_bytes,
                       source_tool, description, created_at, expired, expired_reason, cleaned_at
                FROM backup_meta
                {where}
                ORDER BY created_at DESC
                LIMIT ? OFFSET ?
                """,
                (*params, limit, offset),
            ).fetchall()

        items = [
            {
                "backup_id": str(row["backup_id"]),
                "workbook_id": str(row["workbook_id"]),
                "file_path": str(row["file_path"]),
                "backup_file_path": str(row["backup_file_path"]),
                "file_size_bytes": int(row["file_size_bytes"]),
                "source_tool": str(row["source_tool"]),
                "description": str(row["description"]),
                "created_at": str(row["created_at"]),
                "expired": bool(row["expired"]),
                "expired_reason": row["expired_reason"],
                "cleaned_at": row["cleaned_at"],
            }
            for row in rows
        ]
        return int(total_row["c"]), items

    def list_active_backup_files(self, workbook_id: str | None = None) -> list[dict[str, object]]:
        where = "WHERE expired = 0"
        params: list[object] = []
        if workbook_id:
            where += " AND workbook_id = ?"
            params.append(workbook_id)

        with self._db.connect() as conn:
            rows = conn.execute(
                f"""
                SELECT backup_id, workbook_id, backup_file_path, file_size_bytes, created_at
                FROM backup_meta
                {where}
                ORDER BY created_at ASC
                """,
                params,
            ).fetchall()

        return [
            {
                "backup_id": str(row["backup_id"]),
                "workbook_id": str(row["workbook_id"]),
                "backup_file_path": str(row["backup_file_path"]),
                "file_size_bytes": int(row["file_size_bytes"]),
                "created_at": str(row["created_at"]),
            }
            for row in rows
        ]

    def list_expired_uncleaned(self) -> list[dict[str, object]]:
        with self._db.connect() as conn:
            rows = conn.execute(
                """
                SELECT backup_id, backup_file_path, file_size_bytes
                FROM backup_meta
                WHERE expired = 1 AND cleaned_at IS NULL
                ORDER BY created_at ASC
                """
            ).fetchall()
        return [
            {
                "backup_id": str(row["backup_id"]),
                "backup_file_path": str(row["backup_file_path"]),
                "file_size_bytes": int(row["file_size_bytes"] or 0),
            }
            for row in rows
        ]

    def mark_cleaned(self, backup_ids: list[str], cleaned_at: str) -> int:
        if not backup_ids:
            return 0
        placeholders = ",".join("?" for _ in backup_ids)
        with self._db.connect() as conn:
            cur = conn.execute(
                f"""
                UPDATE backup_meta
                SET cleaned_at = ?
                WHERE backup_id IN ({placeholders}) AND cleaned_at IS NULL
                """,
                (cleaned_at, *backup_ids),
            )
            return int(cur.rowcount)

    def expire_backups(self, backup_ids: list[str], *, reason: str) -> list[dict[str, object]]:
        if not backup_ids:
            return []
        placeholders = ",".join("?" for _ in backup_ids)
        with self._db.connect() as conn:
            rows = conn.execute(
                f"""
                SELECT backup_id, backup_file_path, file_size_bytes
                FROM backup_meta
                WHERE expired = 0 AND backup_id IN ({placeholders})
                """,
                backup_ids,
            ).fetchall()
            conn.execute(
                f"""
                UPDATE backup_meta
                SET expired = 1, expired_reason = ?, cleaned_at = NULL
                WHERE expired = 0 AND backup_id IN ({placeholders})
                """,
                (reason, *backup_ids),
            )

        return [
            {
                "backup_id": str(row["backup_id"]),
                "backup_file_path": str(row["backup_file_path"]),
                "file_size_bytes": int(row["file_size_bytes"]),
            }
            for row in rows
        ]

    def expire_by_age(self, cutoff_ts: str, *, reason: str) -> int:
        with self._db.connect() as conn:
            cur = conn.execute(
                """
                UPDATE backup_meta
                SET expired = 1, expired_reason = ?, cleaned_at = NULL
                WHERE expired = 0 AND created_at < ?
                """,
                (reason, cutoff_ts),
            )
            return int(cur.rowcount)

    def get_stats(self) -> dict[str, object]:
        with self._db.connect() as conn:
            row = conn.execute(
                """
                SELECT
                    SUM(CASE WHEN expired = 0 THEN 1 ELSE 0 END) AS active_count,
                    SUM(CASE WHEN expired = 0 THEN file_size_bytes ELSE 0 END) AS total_size_bytes,
                    MIN(CASE WHEN expired = 0 THEN created_at ELSE NULL END) AS oldest_active_at
                FROM backup_meta
                """
            ).fetchone()
        return {
            "active_count": int(row["active_count"] or 0),
            "total_size_bytes": int(row["total_size_bytes"] or 0),
            "oldest_active_at": row["oldest_active_at"],
        }
