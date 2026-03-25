from __future__ import annotations

import json
from dataclasses import dataclass

from excelforge.persistence.db import Database


@dataclass
class AuditRecord:
    operation_id: str
    tool_name: str
    workbook_id: str | None
    file_path: str | None
    actor_id: str
    os_user: str
    machine_name: str
    client_name: str | None
    client_request_id: str | None
    started_at: str
    duration_ms: int
    success: bool
    code: str
    message: str
    affected_sheet: str | None
    affected_range: str | None
    snapshot_id: str | None
    args_summary: dict[str, object] | None


class AuditRepository:
    def __init__(self, db: Database) -> None:
        self._db = db

    def insert(self, record: AuditRecord) -> None:
        with self._db.connect() as conn:
            conn.execute(
                """
                INSERT INTO audit_operation (
                    operation_id, tool_name, workbook_id, file_path, actor_id,
                    os_user, machine_name, client_name, client_request_id,
                    started_at, duration_ms, success, code, message,
                    affected_sheet, affected_range, snapshot_id, args_summary
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    record.operation_id,
                    record.tool_name,
                    record.workbook_id,
                    record.file_path,
                    record.actor_id,
                    record.os_user,
                    record.machine_name,
                    record.client_name,
                    record.client_request_id,
                    record.started_at,
                    record.duration_ms,
                    1 if record.success else 0,
                    record.code,
                    record.message,
                    record.affected_sheet,
                    record.affected_range,
                    record.snapshot_id,
                    json.dumps(record.args_summary or {}, ensure_ascii=False),
                ),
            )

    def list_operations(
        self,
        *,
        workbook_id: str | None,
        tool_name: str | None,
        success_only: bool,
        limit: int,
        offset: int,
    ) -> tuple[int, list[dict[str, object]]]:
        conditions: list[str] = []
        params: list[object] = []

        if workbook_id:
            conditions.append("workbook_id = ?")
            params.append(workbook_id)
        if tool_name:
            conditions.append("tool_name = ?")
            params.append(tool_name)
        if success_only:
            conditions.append("success = 1")

        where_clause = ""
        if conditions:
            where_clause = "WHERE " + " AND ".join(conditions)

        with self._db.connect() as conn:
            total = conn.execute(
                f"SELECT COUNT(1) AS c FROM audit_operation {where_clause}",
                params,
            ).fetchone()["c"]

            rows = conn.execute(
                f"""
                SELECT operation_id, tool_name, workbook_id, started_at, duration_ms,
                       success, code, message, affected_sheet, affected_range,
                       snapshot_id, client_request_id
                FROM audit_operation
                {where_clause}
                ORDER BY started_at DESC
                LIMIT ? OFFSET ?
                """,
                (*params, limit, offset),
            ).fetchall()

        items = [
            {
                "operation_id": row["operation_id"],
                "tool_name": row["tool_name"],
                "workbook_id": row["workbook_id"],
                "started_at": row["started_at"],
                "duration_ms": int(row["duration_ms"]),
                "success": bool(row["success"]),
                "code": row["code"],
                "message": row["message"],
                "affected_sheet": row["affected_sheet"],
                "affected_range": row["affected_range"],
                "snapshot_id": row["snapshot_id"],
                "client_request_id": row["client_request_id"],
            }
            for row in rows
        ]
        return int(total), items

    def cleanup_older_than(self, cutoff_timestamp: str) -> int:
        with self._db.connect() as conn:
            cur = conn.execute(
                "DELETE FROM audit_operation WHERE started_at < ?",
                (cutoff_timestamp,),
            )
            return int(cur.rowcount)

    def get_operation(self, operation_id: str) -> dict[str, object] | None:
        with self._db.connect() as conn:
            row = conn.execute(
                """
                SELECT operation_id, tool_name, workbook_id, file_path, started_at, duration_ms,
                       success, code, message, affected_sheet, affected_range,
                       snapshot_id, client_request_id
                FROM audit_operation
                WHERE operation_id = ?
                """,
                (operation_id,),
            ).fetchone()

        if row is None:
            return None

        return {
            "operation_id": row["operation_id"],
            "tool_name": row["tool_name"],
            "workbook_id": row["workbook_id"],
            "file_path": row["file_path"],
            "started_at": row["started_at"],
            "duration_ms": int(row["duration_ms"]),
            "success": bool(row["success"]),
            "code": row["code"],
            "message": row["message"],
            "affected_sheet": row["affected_sheet"],
            "affected_range": row["affected_range"],
            "snapshot_id": row["snapshot_id"],
            "client_request_id": row["client_request_id"],
        }
