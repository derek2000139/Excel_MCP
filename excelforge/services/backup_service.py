from __future__ import annotations

import shutil
import threading
from datetime import timedelta
from pathlib import Path
from typing import Any

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.persistence.backup_repo import BackupMetaRecord, BackupRepository
from excelforge.runtime.handle_ownership import ensure_related_handle_owned, ensure_workbook_id_owned
from excelforge.runtime.workbook_registry import WorkbookHandle
from excelforge.runtime.workbook_registry import WorkbookRegistry
from excelforge.services.snapshot_service import SnapshotService
from excelforge.utils.ids import generate_id
from excelforge.utils.timestamps import utc_now, utc_now_rfc3339

MANUAL_RESTORE_INSTRUCTIONS = (
    "Manual restore steps:\n"
    "1. Call workbook.close_file (force_discard=true) to close current workbook.\n"
    "2. Copy backup_file_path over file_path to replace the current workbook file.\n"
    "3. Call workbook.open_file to reopen the workbook."
)
LARGE_FILE_WARNING_THRESHOLD = 100 * 1024 * 1024


class BackupService:
    def __init__(
        self,
        config: AppConfig,
        backup_repo: BackupRepository,
        workbook_registry: WorkbookRegistry | None = None,
        snapshot_service: SnapshotService | None = None,
    ) -> None:
        self._config = config
        self._backup_repo = backup_repo
        self._workbook_registry = workbook_registry
        self._snapshot_service = snapshot_service
        self._cleanup_lock = threading.Lock()

    def _runtime_fingerprint(self) -> str | None:
        if self._workbook_registry is None:
            return None
        return self._workbook_registry.runtime_fingerprint

    def create_backup(
        self,
        *,
        workbook: WorkbookHandle,
        source_tool: str,
        description: str,
        source_operation_id: str | None = None,
    ) -> tuple[str, list[str]]:
        wb = workbook.workbook_obj
        source_path = Path(workbook.file_path)
        if not source_path.exists():
            try:
                wb.Save()
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_BACKUP_FAILED, f"Failed to save workbook before backup: {exc}") from exc

        if not source_path.exists():
            raise ExcelForgeError(ErrorCode.E500_BACKUP_FAILED, f"Workbook file does not exist: {source_path}")

        warnings: list[str] = []
        file_size = int(source_path.stat().st_size)
        if file_size > LARGE_FILE_WARNING_THRESHOLD:
            warnings.append(
                f"Workbook is large ({round(file_size / 1024 / 1024, 2)} MB), backup may take longer."
            )

        backup_id = generate_id("bak")
        backup_path = self._config.backups_dir / f"{backup_id}_{source_path.name}"
        try:
            shutil.copy2(source_path, backup_path)
        except Exception as exc:
            raise ExcelForgeError(ErrorCode.E500_BACKUP_FAILED, f"Failed to copy workbook backup: {exc}") from exc

        copied_size = int(backup_path.stat().st_size) if backup_path.exists() else -1
        if copied_size != file_size:
            backup_path.unlink(missing_ok=True)
            raise ExcelForgeError(
                ErrorCode.E500_BACKUP_FAILED,
                f"Backup file size mismatch: source={file_size}, backup={copied_size}",
            )

        self._backup_repo.insert_meta(
            BackupMetaRecord(
                backup_id=backup_id,
                workbook_id=workbook.workbook_id,
                file_path=str(source_path),
                backup_file_path=str(backup_path),
                file_size_bytes=file_size,
                source_tool=source_tool,
                source_operation_id=source_operation_id,
                description=description,
                created_at=utc_now_rfc3339(),
            )
        )
        self._trigger_async_cleanup()
        return backup_id, warnings

    def list_backups(
        self,
        *,
        workbook_id: str | None,
        file_path: str | None,
        limit: int,
        offset: int,
    ) -> dict[str, Any]:
        if workbook_id:
            ensure_workbook_id_owned(workbook_id, self._runtime_fingerprint())
        total, items = self._backup_repo.list_backups(
            workbook_id=workbook_id,
            file_path=file_path,
            limit=limit,
            offset=offset,
        )
        has_more = (offset + len(items)) < total
        next_offset = (offset + len(items)) if has_more else None
        return {
            "total": total,
            "has_more": has_more,
            "next_offset": next_offset,
            "items": items,
            "manual_restore_instructions": MANUAL_RESTORE_INSTRUCTIONS,
        }

    def restore_file(
        self,
        *,
        workbook_id: str,
        backup_id: str,
    ) -> dict[str, Any]:
        ensure_workbook_id_owned(workbook_id, self._runtime_fingerprint())
        record = self._backup_repo.get_backup(backup_id)
        if record is None:
            raise ExcelForgeError(ErrorCode.E404_SNAPSHOT_NOT_FOUND, f"Backup not found: {backup_id}")

        ensure_related_handle_owned(
            handle_kind="backup",
            handle_id=backup_id,
            owner_workbook_id=str(record.workbook_id),
            runtime_fingerprint=self._runtime_fingerprint(),
        )
        if record.workbook_id != workbook_id:
            raise ExcelForgeError(
                ErrorCode.E400_INVALID_ARGUMENT,
                f"Backup {backup_id} does not belong to workbook {workbook_id}",
            )

        backup_path = Path(record.backup_file_path)
        if not backup_path.exists():
            raise ExcelForgeError(ErrorCode.E404_SNAPSHOT_NOT_FOUND, f"Backup file not found: {backup_path}")

        pre_restore_backup_id = None
        if self._workbook_registry:
            handle = self._workbook_registry.get(workbook_id)
            if handle:
                pre_restore_backup_id, _ = self.create_backup(
                    workbook=handle,
                    source_tool="backup.restore_file",
                    description=f"Pre-restore backup before restoring {backup_id}",
                )

        if self._snapshot_service and self._workbook_registry:
            self._snapshot_service.expire_all_for_workbook(workbook_id)

        new_workbook_id = generate_id("wb")
        return {
            "backup_id": backup_id,
            "pre_restore_backup_id": pre_restore_backup_id,
            "original_workbook_id": workbook_id,
            "new_workbook_id": new_workbook_id,
            "file_path": str(backup_path),
            "invalidated_snapshots": 0,
            "restored_at": utc_now_rfc3339(),
        }

    def get_stats(self) -> dict[str, Any]:
        stats = self._backup_repo.get_stats()
        stats["limits"] = {
            "max_per_workbook": int(self._config.backup.max_per_workbook),
            "max_total_size_mb": int(self._config.backup.max_total_size_mb),
            "max_age_hours": int(self._config.backup.max_age_hours),
        }
        return stats

    def run_cleanup(self) -> dict[str, int]:
        expired_by_age = self._expire_by_age()
        expired_by_quota = self._expire_by_quota()
        deleted_files, deleted_bytes, cleaned_rows = self._delete_expired_files()
        return {
            "backups_expired_by_age": expired_by_age,
            "backups_expired_by_quota": expired_by_quota,
            "backup_files_removed": deleted_files,
            "backup_bytes_removed": deleted_bytes,
            "backup_rows_cleaned": cleaned_rows,
        }

    def _expire_by_age(self) -> int:
        cutoff_ts = (utc_now() - timedelta(hours=self._config.backup.max_age_hours)).isoformat().replace("+00:00", "Z")
        return self._backup_repo.expire_by_age(cutoff_ts, reason="time_expired")

    def _expire_by_quota(self) -> int:
        to_expire: list[str] = []

        per_limit = int(self._config.backup.max_per_workbook)
        if per_limit > 0:
            active_rows = self._backup_repo.list_active_backup_files(workbook_id=None)
            by_workbook: dict[str, list[dict[str, object]]] = {}
            for row in active_rows:
                by_workbook.setdefault(str(row["workbook_id"]), []).append(row)
            for rows in by_workbook.values():
                over = len(rows) - per_limit
                if over > 0:
                    to_expire.extend(str(r["backup_id"]) for r in rows[:over])

        max_bytes = int(self._config.backup.max_total_size_mb) * 1024 * 1024
        if max_bytes > 0:
            active_rows = self._backup_repo.list_active_backup_files(workbook_id=None)
            total_bytes = sum(int(r.get("file_size_bytes", 0) or 0) for r in active_rows)
            marked = set(to_expire)
            for row in active_rows:
                if total_bytes <= max_bytes:
                    break
                backup_id = str(row["backup_id"])
                if backup_id in marked:
                    continue
                marked.add(backup_id)
                to_expire.append(backup_id)
                total_bytes -= int(row.get("file_size_bytes", 0) or 0)

        if not to_expire:
            return 0
        expired_rows = self._backup_repo.expire_backups(to_expire, reason="quota_limit")
        return len(expired_rows)

    def _delete_expired_files(self) -> tuple[int, int, int]:
        entries = self._backup_repo.list_expired_uncleaned()
        removed_files = 0
        removed_bytes = 0
        cleaned_ids: list[str] = []
        for row in entries:
            backup_id = str(row["backup_id"])
            backup_path = Path(str(row["backup_file_path"]))
            file_size = int(row.get("file_size_bytes", 0) or 0)
            try:
                if backup_path.exists():
                    backup_path.unlink(missing_ok=True)
                    removed_files += 1
                    removed_bytes += file_size
            finally:
                cleaned_ids.append(backup_id)
        cleaned_rows = self._backup_repo.mark_cleaned(cleaned_ids, utc_now_rfc3339())
        return removed_files, removed_bytes, cleaned_rows

    def _trigger_async_cleanup(self) -> None:
        def _run() -> None:
            if not self._cleanup_lock.acquire(blocking=False):
                return
            try:
                self.run_cleanup()
            finally:
                self._cleanup_lock.release()

        threading.Thread(target=_run, name="backup-cleanup", daemon=True).start()
