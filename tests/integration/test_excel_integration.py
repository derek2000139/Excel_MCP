from __future__ import annotations

from pathlib import Path

import pytest
import yaml

from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.server import create_app


pytestmark = [pytest.mark.integration]


def _create_sample_workbook(path: Path) -> None:
    import win32com.client  # type: ignore

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Sheet1"
    ws.Cells(1, 1).Value = "Header"
    ws.Cells(2, 1).Value = "A"
    wb.SaveAs(str(path))
    wb.Close(SaveChanges=False)
    excel.Quit()


@pytest.fixture()
def app_with_workbook(tmp_path: Path):
    workbook_path = tmp_path / "sample.xlsx"
    _create_sample_workbook(workbook_path)

    runtime = tmp_path / "runtime"
    config_path = tmp_path / "config.yaml"
    cfg = {
        "server": {"version": "0.3.0", "actor_id": "test-client"},
        "excel": {
            "visible": False,
            "disable_events": True,
            "disable_alerts": True,
            "force_disable_macros": True,
            "health_ping_enabled": True,
            "max_rebuild_attempts": 3,
        },
        "paths": {
            "allowed_roots": [str(tmp_path)],
            "snapshots_dir": str(runtime / "snapshots"),
            "backups_dir": str(runtime / "backups"),
            "sqlite_path": str(runtime / "excelforge.db"),
        },
        "limits": {
            "max_open_workbooks": 8,
            "max_read_cells": 10000,
            "max_write_cells": 10000,
            "max_snapshot_cells": 10000,
            "max_insert_rows": 1000,
            "max_insert_columns": 100,
            "max_vba_code_size_bytes": 1048576,
            "default_read_rows": 200,
            "max_read_rows": 1000,
            "operation_timeout_seconds": 30,
            "max_create_sheets": 20,
        },
        "snapshot": {
            "max_per_workbook": 50,
            "max_total_size_mb": 200,
            "max_age_hours": 24,
            "cleanup_interval_ops": 100,
            "preview_token_ttl_minutes": 5,
        },
        "backup": {
            "max_per_workbook": 10,
            "max_total_size_mb": 500,
            "max_age_hours": 48,
            "confirm_token_ttl_minutes": 5,
        },
        "retention": {"audit_days": 30},
    }
    config_path.write_text(yaml.safe_dump(cfg), encoding="utf-8")

    app = create_app(str(config_path))
    try:
        yield app, workbook_path
    finally:
        app.shutdown()


def test_workbook_open_save_close(app_with_workbook) -> None:
    app, workbook_path = app_with_workbook

    opened = app.tools_ctx.workbook_service.open_file(str(workbook_path), read_only=False)
    wb_id = opened["workbook_id"]

    saved = app.tools_ctx.workbook_service.save_file(wb_id)
    assert saved["dirty"] is False

    closed = app.tools_ctx.workbook_service.close_file(wb_id)
    assert closed["closed"] is True


def test_range_write_and_rollback(app_with_workbook) -> None:
    app, workbook_path = app_with_workbook
    wb = app.tools_ctx.workbook_service.open_file(str(workbook_path), read_only=False)
    wb_id = wb["workbook_id"]

    write = app.tools_ctx.range_service.write_values(
        workbook_id=wb_id,
        sheet_name="Sheet1",
        start_cell="B2",
        values=[[1], [2], [3]],
    )
    snap_id = write["snapshot_id"]

    preview = app.tools_ctx.rollback_service.preview_snapshot(snapshot_id=snap_id, sample_limit=10)
    token = preview["preview_token"]

    restore = app.tools_ctx.rollback_service.restore_snapshot(snapshot_id=snap_id, preview_token=token)
    assert restore["cells_restored"] >= 1


def test_dirty_close_rejected(app_with_workbook) -> None:
    app, workbook_path = app_with_workbook
    wb = app.tools_ctx.workbook_service.open_file(str(workbook_path), read_only=False)
    wb_id = wb["workbook_id"]

    app.tools_ctx.range_service.write_values(
        workbook_id=wb_id,
        sheet_name="Sheet1",
        start_cell="C2",
        values=[["x"]],
    )

    with pytest.raises(ExcelForgeError) as exc:
        app.tools_ctx.workbook_service.close_file(wb_id)
    assert exc.value.code == ErrorCode.E409_WORKBOOK_DIRTY
