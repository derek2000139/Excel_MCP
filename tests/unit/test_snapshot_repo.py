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
from excelforge.utils.timestamps import utc_now_rfc3339


def _build_config(tmp_path: Path) -> AppConfig:
    data = tmp_path / "runtime"
    return AppConfig(
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


def test_snapshot_repo_preview_token_lifecycle(tmp_path: Path) -> None:
    cfg = _build_config(tmp_path)
    db = Database(cfg)
    db.init_schema()
    repo = SnapshotRepository(db)

    snapshot_file = tmp_path / "s.json.gz"
    snapshot_file.write_text("x")

    record = SnapshotMetaRecord(
        snapshot_id="snap_1",
        workbook_id="wb_1",
        file_path=str(tmp_path / "book.xlsx"),
        sheet_name="Sheet1",
        range_address="A1:B2",
        source_tool="range.write_values",
        created_at=utc_now_rfc3339(),
        cell_count=4,
        file_path_snapshot=str(snapshot_file),
    )
    repo.insert_meta(record)

    token = "rtok_1"
    now = utc_now_rfc3339()
    repo.insert_preview_token(token=token, snapshot_id="snap_1", created_at=now, expires_at=now)
    token_row = repo.get_preview_token(token)
    assert token_row is not None
    assert token_row["snapshot_id"] == "snap_1"

    repo.mark_preview_token_used(token)
    token_row2 = repo.get_preview_token(token)
    assert token_row2 is not None and token_row2["used"] is True
