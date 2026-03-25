from __future__ import annotations

from datetime import timedelta
from pathlib import Path

from excelforge.config import AppConfig
from excelforge.persistence.audit_repo import AuditRepository
from excelforge.persistence.backup_repo import BackupRepository
from excelforge.persistence.snapshot_repo import SnapshotRepository
from excelforge.utils.timestamps import utc_now, utc_now_rfc3339


class CleanupService:
    def __init__(
        self,
        config: AppConfig,
        audit_repo: AuditRepository,
        snapshot_repo: SnapshotRepository,
        backup_repo: BackupRepository,
    ) -> None:
        self._config = config
        self._audit_repo = audit_repo
        self._snapshot_repo = snapshot_repo
        self._backup_repo = backup_repo

    def run(self) -> dict[str, int]:
        now = utc_now()

        snapshot_cutoff = (now - timedelta(hours=self._config.snapshot.max_age_hours)).isoformat().replace(
            "+00:00", "Z"
        )
        backup_cutoff = (now - timedelta(hours=self._config.backup.max_age_hours)).isoformat().replace("+00:00", "Z")
        audit_cutoff = (now - timedelta(days=self._config.retention.audit_days)).isoformat().replace(
            "+00:00", "Z"
        )

        snapshots_expired = self._snapshot_repo.expire_by_age(
            cutoff_ts=snapshot_cutoff,
            reason="time_expired",
            workbook_id=None,
        )
        snapshot_files_removed, snapshot_space_freed, snapshot_cleaned = self._cleanup_snapshot_files()

        preview_tokens_removed = self._snapshot_repo.cleanup_expired_tokens(utc_now_rfc3339())

        backups_expired = self._backup_repo.expire_by_age(backup_cutoff, reason="time_expired")
        backup_files_removed, backup_space_freed, backup_cleaned = self._cleanup_backup_files()

        audit_rows_removed = self._audit_repo.cleanup_older_than(audit_cutoff)

        return {
            "snapshots_expired": int(snapshots_expired),
            "snapshot_files_removed": int(snapshot_files_removed),
            "snapshot_space_freed_bytes": int(snapshot_space_freed),
            "snapshot_rows_cleaned": int(snapshot_cleaned),
            "backups_expired": int(backups_expired),
            "backup_files_removed": int(backup_files_removed),
            "backup_space_freed_bytes": int(backup_space_freed),
            "backup_rows_cleaned": int(backup_cleaned),
            "preview_tokens_removed": int(preview_tokens_removed),
            "audit_rows_removed": int(audit_rows_removed),
        }

    def _cleanup_snapshot_files(self) -> tuple[int, int, int]:
        stale_files = self._snapshot_repo.list_expired_uncleaned(workbook_id=None)
        files_removed = 0
        space_freed = 0
        cleaned_ids: list[str] = []
        for row in stale_files:
            snapshot_id = str(row["snapshot_id"])
            path = Path(str(row["file_path_snapshot"]))
            file_size = int(row.get("file_size_bytes", 0) or 0)
            if path.exists():
                path.unlink(missing_ok=True)
                files_removed += 1
                space_freed += file_size
            cleaned_ids.append(snapshot_id)

        cleaned_count = self._snapshot_repo.mark_cleaned(cleaned_ids, utc_now_rfc3339())
        return files_removed, space_freed, cleaned_count

    def _cleanup_backup_files(self) -> tuple[int, int, int]:
        stale_files = self._backup_repo.list_expired_uncleaned()
        files_removed = 0
        space_freed = 0
        cleaned_ids: list[str] = []
        for row in stale_files:
            backup_id = str(row["backup_id"])
            path = Path(str(row["backup_file_path"]))
            file_size = int(row.get("file_size_bytes", 0) or 0)
            if path.exists():
                path.unlink(missing_ok=True)
                files_removed += 1
                space_freed += file_size
            cleaned_ids.append(backup_id)

        cleaned_count = self._backup_repo.mark_cleaned(cleaned_ids, utc_now_rfc3339())
        return files_removed, space_freed, cleaned_count
