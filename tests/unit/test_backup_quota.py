from __future__ import annotations

from pathlib import Path

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
from excelforge.persistence.backup_repo import BackupMetaRecord, BackupRepository
from excelforge.persistence.db import Database
from excelforge.services.backup_service import BackupService


def _build_backup_service(
    tmp_path: Path,
    *,
    max_per_workbook: int,
    max_total_size_mb: int,
) -> tuple[BackupService, BackupRepository]:
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
        backup=BackupConfig(
            max_per_workbook=max_per_workbook,
            max_total_size_mb=max_total_size_mb,
            max_age_hours=240,
            confirm_token_ttl_minutes=5,
        ),
        retention=RetentionConfig(),
    )
    db = Database(cfg)
    db.init_schema()
    repo = BackupRepository(db)
    service = BackupService(cfg, repo)
    return service, repo


def _insert_backup(
    repo: BackupRepository,
    *,
    backup_id: str,
    workbook_id: str,
    source_file: Path,
    backup_file: Path,
    created_at: str,
) -> None:
    repo.insert_meta(
        BackupMetaRecord(
            backup_id=backup_id,
            workbook_id=workbook_id,
            file_path=str(source_file),
            backup_file_path=str(backup_file),
            file_size_bytes=int(backup_file.stat().st_size),
            source_tool="range.insert_rows",
            source_operation_id="op_test",
            description="quota test",
            created_at=created_at,
        )
    )


def test_backup_quota_enforces_max_per_workbook(tmp_path: Path) -> None:
    service, repo = _build_backup_service(
        tmp_path,
        max_per_workbook=2,
        max_total_size_mb=500,
    )
    source = tmp_path / "book.xlsx"
    source.write_bytes(b"x")
    b1 = tmp_path / "b1.xlsx"
    b2 = tmp_path / "b2.xlsx"
    b3 = tmp_path / "b3.xlsx"
    b1.write_bytes(b"a" * 10)
    b2.write_bytes(b"b" * 10)
    b3.write_bytes(b"c" * 10)

    wb_id = "wb_g1_aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
    _insert_backup(repo, backup_id="bak_1", workbook_id=wb_id, source_file=source, backup_file=b1, created_at="2026-03-22T10:00:00Z")
    _insert_backup(repo, backup_id="bak_2", workbook_id=wb_id, source_file=source, backup_file=b2, created_at="2026-03-22T10:00:01Z")
    _insert_backup(repo, backup_id="bak_3", workbook_id=wb_id, source_file=source, backup_file=b3, created_at="2026-03-22T10:00:02Z")

    stats = service.run_cleanup()

    assert stats["backups_expired_by_quota"] == 1
    active = repo.list_active_backup_files(workbook_id=wb_id)
    active_ids = [str(item["backup_id"]) for item in active]
    assert active_ids == ["bak_2", "bak_3"]
    assert not b1.exists()
    assert b2.exists()
    assert b3.exists()


def test_backup_quota_enforces_global_total_size(tmp_path: Path) -> None:
    service, repo = _build_backup_service(
        tmp_path,
        max_per_workbook=20,
        max_total_size_mb=1,
    )
    source = tmp_path / "book.xlsx"
    source.write_bytes(b"x")
    old_file = tmp_path / "old.xlsx"
    new_file = tmp_path / "new.xlsx"
    old_file.write_bytes(b"a" * 700_000)
    new_file.write_bytes(b"b" * 700_000)

    _insert_backup(
        repo,
        backup_id="bak_old",
        workbook_id="wb_g1_aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
        source_file=source,
        backup_file=old_file,
        created_at="2026-03-22T10:00:00Z",
    )
    _insert_backup(
        repo,
        backup_id="bak_new",
        workbook_id="wb_g1_bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
        source_file=source,
        backup_file=new_file,
        created_at="2026-03-22T10:00:01Z",
    )

    stats = service.run_cleanup()

    assert stats["backups_expired_by_quota"] == 1
    active = repo.list_active_backup_files(workbook_id=None)
    active_ids = [str(item["backup_id"]) for item in active]
    assert active_ids == ["bak_new"]
    assert not old_file.exists()
    assert new_file.exists()
