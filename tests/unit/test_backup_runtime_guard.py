from __future__ import annotations

from pathlib import Path

import pytest

from excelforge.config import (
    AppConfig,
    BackupConfig,
    ExcelConfig,
    LimitsConfig,
    PathsConfig,
    RetentionConfig,
    ServerConfig,
    SnapshotConfig,
)
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.persistence.backup_repo import BackupMetaRecord, BackupRepository
from excelforge.persistence.db import Database
from excelforge.runtime.workbook_registry import WorkbookRegistry
from excelforge.services.backup_service import BackupService
from excelforge.utils.ids import generate_workbook_id


def test_restore_file_rejects_foreign_runtime_backup(tmp_path: Path) -> None:
    data = tmp_path / "runtime"
    cfg = AppConfig(
        server=ServerConfig(),
        excel=ExcelConfig(),
        paths=PathsConfig(
            allowed_roots=[str(tmp_path)],
            snapshots_dir=str(data / "snapshots"),
            backups_dir=str(data / "backups"),
            sqlite_path=str(data / "excelforge.db"),
        ),
        limits=LimitsConfig(),
        snapshot=SnapshotConfig(),
        backup=BackupConfig(),
        retention=RetentionConfig(),
    )
    db = Database(cfg)
    db.init_schema()
    repo = BackupRepository(db)
    registry = WorkbookRegistry(runtime_fingerprint="deadbeef")
    service = BackupService(cfg, repo, workbook_registry=registry)

    source_file = tmp_path / "book.xlsx"
    source_file.write_bytes(b"x")
    backup_file = tmp_path / "backup.xlsx"
    backup_file.write_bytes(b"backup")
    repo.insert_meta(
        BackupMetaRecord(
            backup_id="bak_foreign",
            workbook_id=generate_workbook_id(1, "cafebabe"),
            file_path=str(source_file),
            backup_file_path=str(backup_file),
            file_size_bytes=int(backup_file.stat().st_size),
            source_tool="range.insert_rows",
            source_operation_id="op_test",
            description="foreign runtime backup",
            created_at="2026-03-29T10:00:00Z",
        )
    )

    with pytest.raises(ExcelForgeError) as exc_info:
        service.restore_file(
            workbook_id=generate_workbook_id(1, "deadbeef"),
            backup_id="bak_foreign",
        )

    assert exc_info.value.code == ErrorCode.E424_HANDLE_RUNTIME_MISMATCH
