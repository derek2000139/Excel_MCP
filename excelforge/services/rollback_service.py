from __future__ import annotations

from typing import Any

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.persistence.snapshot_repo import SnapshotRepository
from excelforge.runtime.excel_worker import ExcelWorker
from excelforge.services.snapshot_service import SnapshotService


class RollbackService:
    def __init__(
        self,
        config: AppConfig,
        worker: ExcelWorker,
        snapshot_repo: SnapshotRepository,
        snapshot_service: SnapshotService,
    ) -> None:
        self._config = config
        self._worker = worker
        self._snapshot_repo = snapshot_repo
        self._snapshot_service = snapshot_service

    def list_backup_snapshots(
        self,
        *,
        workbook_id: str,
    ) -> dict[str, Any]:
        with self._lock:
            cache = self._cache.get(workbook_id)
            if cache is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            backups = self._backup_service.list_backups(workbook_id=workbook_id)
            snapshots = self._snapshot_service.list_snapshots(workbook_id=workbook_id)

            return {
                "workbook_id": workbook_id,
                "backups": backups,
                "snapshots": snapshots,
                "count": len(backups) + len(snapshots),
            }

    def undo_last(self, *, workbook_id: str) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            with self._lock:
                cache = self._cache.get(workbook_id)
                if cache is None:
                    raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            snapshots = self._snapshot_service.list_snapshots(workbook_id=workbook_id)
            backups = self._backup_service.list_backups(workbook_id=workbook_id)

            all_ops = []
            for s in snapshots:
                all_ops.append({
                    "type": "snapshot",
                    "id": s["snapshot_id"],
                    "timestamp": s["created_at"],
                    "tool": s.get("source_tool", "unknown"),
                    "sheet": s.get("sheet_name"),
                    "range": s.get("range_address"),
                })
            for b in backups:
                all_ops.append({
                    "type": "backup",
                    "id": b["backup_id"],
                    "timestamp": b["created_at"],
                    "tool": b.get("source_tool", "unknown"),
                    "sheet": None,
                    "range": None,
                })

            all_ops.sort(key=lambda x: x["timestamp"], reverse=True)

            skipped = 0
            for op_info in all_ops:
                tool = op_info["tool"]
                if tool in self._READONLY_TOOLS:
                    skipped += 1
                    continue

                if op_info["type"] == "snapshot":
                    result = self._snapshot_service.restore_snapshot(
                        workbook_id=workbook_id,
                        snapshot_id=op_info["id"],
                    )
                    return {
                        "workbook_id": workbook_id,
                        "undo_type": "snapshot",
                        "undo_id": op_info["id"],
                        "restored_from": tool,
                        "timestamp": op_info["timestamp"],
                        "skipped_readonly": skipped,
                    }
                else:
                    result = self._backup_service.restore_backup(
                        workbook_id=workbook_id,
                        backup_id=op_info["id"],
                    )
                    return {
                        "workbook_id": workbook_id,
                        "undo_type": "backup",
                        "undo_id": op_info["id"],
                        "restored_from": tool,
                        "timestamp": op_info["timestamp"],
                        "skipped_readonly": skipped,
                    }

            raise ExcelForgeError(ErrorCode.E409_NOTHING_TO_UNDO, "No operations available to undo")

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def list_snapshots(self, workbook_id: str | None, limit: int, offset: int) -> dict[str, Any]:
        total, items = self._snapshot_repo.list_snapshots(workbook_id=workbook_id, limit=limit, offset=offset)
        has_more = (offset + len(items)) < total
        next_offset = (offset + len(items)) if has_more else None
        return {
            "total": total,
            "has_more": has_more,
            "next_offset": next_offset,
            "items": items,
        }

    def preview_snapshot(self, snapshot_id: str, sample_limit: int) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            meta, payload = self._snapshot_service.load_snapshot(snapshot_id)
            workbook_id = str(meta["workbook_id"])
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(
                    ErrorCode.E409_SNAPSHOT_EXPIRED,
                    f"Workbook for snapshot is not currently open: {workbook_id}",
                )

            workbook = handle.workbook_obj
            try:
                ws = workbook.Worksheets(str(meta["sheet_name"]))
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, "Snapshot sheet not found") from exc

            changed_count, sample_diffs = self._snapshot_service.preview_diffs(
                worksheet=ws,
                snapshot_payload=payload,
                sample_limit=sample_limit,
            )
            token, expires_at = self._snapshot_service.create_preview_token(snapshot_id)

            return {
                "snapshot_id": snapshot_id,
                "workbook_id": workbook_id,
                "sheet_name": str(meta["sheet_name"]),
                "range": str(meta["range"]),
                "changed_cells_count": changed_count,
                "sample_diffs": sample_diffs,
                "preview_token": token,
                "preview_token_expires_at": expires_at,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def restore_snapshot(self, snapshot_id: str, preview_token: str) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            self._snapshot_service.consume_preview_token(preview_token, snapshot_id)
            meta, payload = self._snapshot_service.load_snapshot(snapshot_id)

            workbook_id = str(meta["workbook_id"])
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(
                    ErrorCode.E409_SNAPSHOT_EXPIRED,
                    f"Workbook for snapshot is not currently open: {workbook_id}",
                )

            workbook = handle.workbook_obj
            try:
                ws = workbook.Worksheets(str(meta["sheet_name"]))
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, "Snapshot sheet not found") from exc

            replacement_snapshot_id = self._snapshot_service.create_snapshot(
                workbook=handle,
                worksheet=ws,
                range_address=str(meta["range"]),
                source_tool="rollback.restore_snapshot",
            )

            try:
                restored = self._snapshot_service.restore_snapshot(
                    workbook=handle,
                    worksheet=ws,
                    snapshot_payload=payload,
                )
            except Exception as exc:
                raise ExcelForgeError(
                    ErrorCode.E500_SNAPSHOT_RESTORE_FAILED,
                    f"Failed to restore snapshot: {exc}",
                ) from exc

            return {
                "snapshot_id": snapshot_id,
                "workbook_id": workbook_id,
                "sheet_name": str(meta["sheet_name"]),
                "restored_range": str(meta["range"]),
                "cells_restored": restored,
                "replacement_snapshot_id": replacement_snapshot_id,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )
