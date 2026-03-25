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
from excelforge.persistence.db import Database
from excelforge.persistence.snapshot_repo import SnapshotMetaRecord, SnapshotRepository
from excelforge.services.snapshot_service import SnapshotService


def _build_snapshot_service(
    tmp_path: Path,
    *,
    max_per_workbook: int,
    max_total_size_mb: int,
) -> tuple[SnapshotService, SnapshotRepository]:
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
        snapshot=SnapshotConfig(
            max_per_workbook=max_per_workbook,
            max_total_size_mb=max_total_size_mb,
        ),
        backup=BackupConfig(),
        retention=RetentionConfig(),
    )
    db = Database(cfg)
    db.init_schema()
    repo = SnapshotRepository(db)
    service = SnapshotService(cfg, repo)
    return service, repo


def _insert_snapshot(
    repo: SnapshotRepository,
    *,
    snap_id: str,
    workbook_id: str,
    file_path_snapshot: Path,
    created_at: str,
) -> None:
    file_size = int(file_path_snapshot.stat().st_size) if file_path_snapshot.exists() else 0
    repo.insert_meta(
        SnapshotMetaRecord(
            snapshot_id=snap_id,
            workbook_id=workbook_id,
            file_path="D:/ExcelForge/book.xlsx",
            sheet_name="Sheet1",
            range_address="A1",
            source_tool="range.write_values",
            created_at=created_at,
            cell_count=1,
            file_path_snapshot=str(file_path_snapshot),
            file_size_bytes=file_size,
        )
    )


def test_snapshot_quota_enforces_max_per_workbook(tmp_path: Path) -> None:
    service, repo = _build_snapshot_service(
        tmp_path,
        max_per_workbook=2,
        max_total_size_mb=200,
    )
    p1 = tmp_path / "s1.json.gz"
    p2 = tmp_path / "s2.json.gz"
    p3 = tmp_path / "s3.json.gz"
    p1.write_bytes(b"x" * 10)
    p2.write_bytes(b"x" * 10)
    p3.write_bytes(b"x" * 10)

    _insert_snapshot(repo, snap_id="snap_1", workbook_id="wb_g1_aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", file_path_snapshot=p1, created_at="2026-03-22T10:00:00Z")
    _insert_snapshot(repo, snap_id="snap_2", workbook_id="wb_g1_aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", file_path_snapshot=p2, created_at="2026-03-22T10:00:01Z")
    _insert_snapshot(repo, snap_id="snap_3", workbook_id="wb_g1_aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", file_path_snapshot=p3, created_at="2026-03-22T10:00:02Z")

    service._enforce_snapshot_quotas("wb_g1_aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa")

    active = repo.list_active_snapshot_files(workbook_id="wb_g1_aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa")
    assert len(active) == 2
    active_ids = [str(item["snapshot_id"]) for item in active]
    assert "snap_1" not in active_ids
    assert not p1.exists()
    assert p2.exists()
    assert p3.exists()


def test_snapshot_quota_enforces_global_total_size(tmp_path: Path) -> None:
    service, repo = _build_snapshot_service(
        tmp_path,
        max_per_workbook=20,
        max_total_size_mb=1,
    )
    p1 = tmp_path / "g1.json.gz"
    p2 = tmp_path / "g2.json.gz"
    p1.write_bytes(b"x" * 700_000)
    p2.write_bytes(b"x" * 700_000)

    _insert_snapshot(repo, snap_id="snap_old", workbook_id="wb_g1_aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", file_path_snapshot=p1, created_at="2026-03-22T10:00:00Z")
    _insert_snapshot(repo, snap_id="snap_new", workbook_id="wb_g1_bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb", file_path_snapshot=p2, created_at="2026-03-22T10:00:01Z")

    service._enforce_snapshot_quotas("wb_g1_bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb")

    active = repo.list_active_snapshot_files(workbook_id=None)
    active_ids = [str(item["snapshot_id"]) for item in active]
    assert active_ids == ["snap_new"]
    assert not p1.exists()
    assert p2.exists()


def test_expire_workbook_snapshots_deletes_files_immediately(tmp_path: Path) -> None:
    service, repo = _build_snapshot_service(
        tmp_path,
        max_per_workbook=20,
        max_total_size_mb=200,
    )
    workbook_id = "wb_g1_cccccccccccccccccccccccccccccccc"
    p1 = tmp_path / "e1.json.gz"
    p2 = tmp_path / "e2.json.gz"
    p1.write_bytes(b"x")
    p2.write_bytes(b"y")

    _insert_snapshot(repo, snap_id="snap_e1", workbook_id=workbook_id, file_path_snapshot=p1, created_at="2026-03-22T10:00:00Z")
    _insert_snapshot(repo, snap_id="snap_e2", workbook_id=workbook_id, file_path_snapshot=p2, created_at="2026-03-22T10:00:01Z")

    removed = service.expire_workbook_snapshots(workbook_id)
    assert removed == 2
    assert not p1.exists()
    assert not p2.exists()
