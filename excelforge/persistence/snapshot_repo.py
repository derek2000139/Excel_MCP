from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from excelforge.persistence.db import Database


@dataclass
class SnapshotMetaRecord:
    snapshot_id: str
    workbook_id: str
    file_path: str
    sheet_name: str
    range_address: str
    source_tool: str
    created_at: str
    cell_count: int
    file_path_snapshot: str
    file_size_bytes: int = 0
    expired: bool = False
    expired_reason: str | None = None
    cleaned_at: str | None = None


class SnapshotRepository:
    def __init__(self, db: Database) -> None:
        self._db = db

    def insert_meta(self, record: SnapshotMetaRecord) -> None:
        with self._db.connect() as conn:
            conn.execute(
                """
                INSERT INTO snapshot_meta (
                    snapshot_id, workbook_id, file_path, sheet_name, range_address,
                    source_tool, created_at, cell_count, file_path_snapshot,
                    file_size_bytes, expired, expired_reason, cleaned_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    record.snapshot_id,
                    record.workbook_id,
                    record.file_path,
                    record.sheet_name,
                    record.range_address,
                    record.source_tool,
                    record.created_at,
                    record.cell_count,
                    record.file_path_snapshot,
                    int(record.file_size_bytes),
                    1 if record.expired else 0,
                    record.expired_reason,
                    record.cleaned_at,
                ),
            )

    def get_meta(self, snapshot_id: str) -> dict[str, object] | None:
        with self._db.connect() as conn:
            row = conn.execute(
                """
                SELECT snapshot_id, workbook_id, file_path, sheet_name, range_address,
                       source_tool, created_at, cell_count, file_path_snapshot,
                       file_size_bytes, expired, expired_reason, cleaned_at
                FROM snapshot_meta
                WHERE snapshot_id = ?
                """,
                (snapshot_id,),
            ).fetchone()
        if row is None:
            return None
        return {
            "snapshot_id": row["snapshot_id"],
            "workbook_id": row["workbook_id"],
            "file_path": row["file_path"],
            "sheet_name": row["sheet_name"],
            "range": row["range_address"],
            "source_tool": row["source_tool"],
            "created_at": row["created_at"],
            "cell_count": int(row["cell_count"]),
            "file_path_snapshot": row["file_path_snapshot"],
            "file_size_bytes": int(row["file_size_bytes"] or 0),
            "expired": bool(row["expired"]),
            "expired_reason": row["expired_reason"],
            "cleaned_at": row["cleaned_at"],
        }

    def list_snapshots(
        self,
        *,
        workbook_id: str | None,
        limit: int,
        offset: int,
    ) -> tuple[int, list[dict[str, object]]]:
        where = ""
        params: list[object] = []
        if workbook_id:
            where = "WHERE workbook_id = ?"
            params.append(workbook_id)

        with self._db.connect() as conn:
            total = conn.execute(
                f"SELECT COUNT(1) AS c FROM snapshot_meta {where}",
                params,
            ).fetchone()["c"]
            rows = conn.execute(
                f"""
                SELECT snapshot_id, workbook_id, sheet_name, range_address, source_tool,
                       created_at, cell_count, expired, expired_reason, cleaned_at
                FROM snapshot_meta
                {where}
                ORDER BY created_at DESC
                LIMIT ? OFFSET ?
                """,
                (*params, limit, offset),
            ).fetchall()

        items = [
            {
                "snapshot_id": row["snapshot_id"],
                "workbook_id": row["workbook_id"],
                "sheet_name": row["sheet_name"],
                "range": row["range_address"],
                "source_tool": row["source_tool"],
                "created_at": row["created_at"],
                "cell_count": int(row["cell_count"]),
                "expired_reason": row["expired_reason"],
                "cleaned_at": row["cleaned_at"],
                "restorable": (not bool(row["expired"])) and row["cleaned_at"] is None,
            }
            for row in rows
        ]
        return int(total), items

    def expire_by_workbook(self, workbook_id: str, *, reason: str = "workbook_closed") -> int:
        with self._db.connect() as conn:
            cur = conn.execute(
                """
                UPDATE snapshot_meta
                SET expired = 1, expired_reason = ?, cleaned_at = NULL
                WHERE workbook_id = ? AND expired = 0
                """,
                (reason, workbook_id),
            )
            return int(cur.rowcount)

    def expire_by_workbook_with_rows(
        self,
        workbook_id: str,
        *,
        reason: str = "workbook_closed",
    ) -> list[dict[str, object]]:
        with self._db.connect() as conn:
            rows = conn.execute(
                """
                SELECT snapshot_id, file_path_snapshot, file_size_bytes
                FROM snapshot_meta
                WHERE workbook_id = ? AND expired = 0
                ORDER BY created_at ASC
                """,
                (workbook_id,),
            ).fetchall()
            conn.execute(
                """
                UPDATE snapshot_meta
                SET expired = 1, expired_reason = ?, cleaned_at = NULL
                WHERE workbook_id = ? AND expired = 0
                """,
                (reason, workbook_id),
            )
        return [
            {
                "snapshot_id": str(row["snapshot_id"]),
                "file_path_snapshot": str(row["file_path_snapshot"]),
                "file_size_bytes": int(row["file_size_bytes"] or 0),
            }
            for row in rows
        ]

    def expire_by_sheet_with_rows(
        self,
        workbook_id: str,
        sheet_name: str,
        *,
        reason: str = "sheet_structure_changed",
    ) -> list[dict[str, object]]:
        with self._db.connect() as conn:
            rows = conn.execute(
                """
                SELECT snapshot_id, file_path_snapshot, file_size_bytes
                FROM snapshot_meta
                WHERE workbook_id = ? AND sheet_name = ? AND expired = 0
                ORDER BY created_at ASC
                """,
                (workbook_id, sheet_name),
            ).fetchall()
            conn.execute(
                """
                UPDATE snapshot_meta
                SET expired = 1, expired_reason = ?, cleaned_at = NULL
                WHERE workbook_id = ? AND sheet_name = ? AND expired = 0
                """,
                (reason, workbook_id, sheet_name),
            )
        return [
            {
                "snapshot_id": str(row["snapshot_id"]),
                "file_path_snapshot": str(row["file_path_snapshot"]),
                "file_size_bytes": int(row["file_size_bytes"] or 0),
            }
            for row in rows
        ]

    def mark_expired(self, snapshot_id: str, *, reason: str = "manual") -> None:
        with self._db.connect() as conn:
            conn.execute(
                """
                UPDATE snapshot_meta
                SET expired = 1, expired_reason = ?, cleaned_at = NULL
                WHERE snapshot_id = ?
                """,
                (reason, snapshot_id),
            )

    def list_active_snapshot_files(self, workbook_id: str | None = None) -> list[dict[str, object]]:
        where = "WHERE expired = 0"
        params: list[object] = []
        if workbook_id is not None:
            where += " AND workbook_id = ?"
            params.append(workbook_id)

        with self._db.connect() as conn:
            rows = conn.execute(
                f"""
                SELECT snapshot_id, workbook_id, file_path, created_at, file_path_snapshot, file_size_bytes
                FROM snapshot_meta
                {where}
                ORDER BY created_at ASC
                """,
                params,
            ).fetchall()

        return [
            {
                "snapshot_id": str(row["snapshot_id"]),
                "workbook_id": str(row["workbook_id"]),
                "file_path": str(row["file_path"]),
                "created_at": str(row["created_at"]),
                "file_path_snapshot": str(row["file_path_snapshot"]),
                "file_size_bytes": int(row["file_size_bytes"] or 0),
            }
            for row in rows
        ]

    def expire_snapshots(self, snapshot_ids: list[str], *, reason: str) -> list[dict[str, object]]:
        if not snapshot_ids:
            return []
        placeholders = ",".join("?" for _ in snapshot_ids)
        with self._db.connect() as conn:
            rows = conn.execute(
                f"""
                SELECT snapshot_id, file_path_snapshot, file_size_bytes
                FROM snapshot_meta
                WHERE expired = 0 AND snapshot_id IN ({placeholders})
                """,
                snapshot_ids,
            ).fetchall()
            conn.execute(
                f"""
                UPDATE snapshot_meta
                SET expired = 1, expired_reason = ?, cleaned_at = NULL
                WHERE expired = 0 AND snapshot_id IN ({placeholders})
                """,
                (reason, *snapshot_ids),
            )
        return [
            {
                "snapshot_id": str(row["snapshot_id"]),
                "file_path_snapshot": str(row["file_path_snapshot"]),
                "file_size_bytes": int(row["file_size_bytes"] or 0),
            }
            for row in rows
        ]

    def expire_by_age(
        self,
        *,
        cutoff_ts: str,
        reason: str = "time_expired",
        workbook_id: str | None = None,
    ) -> int:
        where = "created_at < ? AND expired = 0"
        params: list[object] = [cutoff_ts]
        if workbook_id is not None:
            where += " AND workbook_id = ?"
            params.append(workbook_id)
        with self._db.connect() as conn:
            cur = conn.execute(
                f"""
                UPDATE snapshot_meta
                SET expired = 1, expired_reason = ?, cleaned_at = NULL
                WHERE {where}
                """,
                (reason, *params),
            )
            return int(cur.rowcount)

    def list_expired_uncleaned(self, *, workbook_id: str | None = None) -> list[dict[str, object]]:
        where = "WHERE expired = 1 AND cleaned_at IS NULL"
        params: list[object] = []
        if workbook_id is not None:
            where += " AND workbook_id = ?"
            params.append(workbook_id)
        with self._db.connect() as conn:
            rows = conn.execute(
                f"""
                SELECT snapshot_id, file_path_snapshot, file_size_bytes
                FROM snapshot_meta
                {where}
                ORDER BY created_at ASC
                """,
                params,
            ).fetchall()
        return [
            {
                "snapshot_id": str(row["snapshot_id"]),
                "file_path_snapshot": str(row["file_path_snapshot"]),
                "file_size_bytes": int(row["file_size_bytes"] or 0),
            }
            for row in rows
        ]

    def mark_cleaned(self, snapshot_ids: list[str], cleaned_at: str) -> int:
        if not snapshot_ids:
            return 0
        placeholders = ",".join("?" for _ in snapshot_ids)
        with self._db.connect() as conn:
            cur = conn.execute(
                f"""
                UPDATE snapshot_meta
                SET cleaned_at = ?
                WHERE snapshot_id IN ({placeholders}) AND cleaned_at IS NULL
                """,
                (cleaned_at, *snapshot_ids),
            )
            return int(cur.rowcount)

    def get_stats(self, workbook_id: str | None = None) -> dict[str, Any]:
        where = ""
        params: list[object] = []
        if workbook_id:
            where = "WHERE workbook_id = ?"
            params.append(workbook_id)

        with self._db.connect() as conn:
            row = conn.execute(
                f"""
                SELECT
                    COUNT(1) AS total_snapshots,
                    SUM(CASE WHEN expired = 0 THEN 1 ELSE 0 END) AS active_snapshots,
                    SUM(CASE WHEN expired = 1 AND cleaned_at IS NULL THEN 1 ELSE 0 END) AS expired_snapshots,
                    SUM(file_size_bytes) AS total_size_bytes,
                    SUM(CASE WHEN expired = 0 THEN file_size_bytes ELSE 0 END) AS active_size_bytes,
                    MIN(CASE WHEN expired = 0 THEN created_at ELSE NULL END) AS oldest_active_at,
                    MAX(CASE WHEN expired = 0 THEN created_at ELSE NULL END) AS newest_active_at
                FROM snapshot_meta
                {where}
                """,
                params,
            ).fetchone()

            workbook_rows = conn.execute(
                f"""
                SELECT workbook_id, file_path,
                       SUM(CASE WHEN expired = 0 THEN 1 ELSE 0 END) AS active_count,
                       SUM(CASE WHEN expired = 0 THEN file_size_bytes ELSE 0 END) AS size_bytes
                FROM snapshot_meta
                {where}
                GROUP BY workbook_id, file_path
                ORDER BY size_bytes DESC
                """,
                params,
            ).fetchall()

        return {
            "total_snapshots": int(row["total_snapshots"] or 0),
            "active_snapshots": int(row["active_snapshots"] or 0),
            "expired_snapshots": int(row["expired_snapshots"] or 0),
            "total_size_bytes": int(row["total_size_bytes"] or 0),
            "active_size_bytes": int(row["active_size_bytes"] or 0),
            "oldest_active_at": row["oldest_active_at"],
            "newest_active_at": row["newest_active_at"],
            "by_workbook": [
                {
                    "workbook_id": str(w["workbook_id"]),
                    "file_path": str(w["file_path"]),
                    "active_count": int(w["active_count"] or 0),
                    "size_bytes": int(w["size_bytes"] or 0),
                }
                for w in workbook_rows
            ],
        }

    def count_active_for_sheet(self, workbook_id: str, sheet_name: str) -> int:
        with self._db.connect() as conn:
            row = conn.execute(
                """
                SELECT COUNT(1) AS c
                FROM snapshot_meta
                WHERE workbook_id = ? AND sheet_name = ? AND expired = 0
                """,
                (workbook_id, sheet_name),
            ).fetchone()
        return int(row["c"] or 0)

    def rename_sheet_refs(self, workbook_id: str, old_name: str, new_name: str) -> int:
        with self._db.connect() as conn:
            cur = conn.execute(
                """
                UPDATE snapshot_meta
                SET sheet_name = ?
                WHERE workbook_id = ? AND sheet_name = ?
                """,
                (new_name, workbook_id, old_name),
            )
            return int(cur.rowcount)

    def insert_preview_token(
        self,
        *,
        token: str,
        snapshot_id: str,
        created_at: str,
        expires_at: str,
    ) -> None:
        with self._db.connect() as conn:
            conn.execute(
                """
                INSERT INTO preview_token (token, snapshot_id, created_at, expires_at, used)
                VALUES (?, ?, ?, ?, 0)
                """,
                (token, snapshot_id, created_at, expires_at),
            )

    def get_preview_token(self, token: str) -> dict[str, object] | None:
        with self._db.connect() as conn:
            row = conn.execute(
                "SELECT token, snapshot_id, created_at, expires_at, used FROM preview_token WHERE token = ?",
                (token,),
            ).fetchone()
        if row is None:
            return None
        return {
            "token": row["token"],
            "snapshot_id": row["snapshot_id"],
            "created_at": row["created_at"],
            "expires_at": row["expires_at"],
            "used": bool(row["used"]),
        }

    def mark_preview_token_used(self, token: str) -> None:
        with self._db.connect() as conn:
            conn.execute(
                "UPDATE preview_token SET used = 1 WHERE token = ?",
                (token,),
            )

    def cleanup_expired_tokens(self, cutoff_ts: str) -> int:
        with self._db.connect() as conn:
            cur = conn.execute(
                "DELETE FROM preview_token WHERE expires_at < ? OR used = 1",
                (cutoff_ts,),
            )
            return int(cur.rowcount)
