from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.config import AppConfig, load_config
from excelforge.persistence.audit_repo import AuditRepository
from excelforge.persistence.backup_repo import BackupRepository
from excelforge.persistence.cleanup import CleanupService
from excelforge.persistence.db import Database
from excelforge.persistence.snapshot_repo import SnapshotRepository
from excelforge.runtime.excel_worker import ExcelWorker
from excelforge.services.audit_service import AuditService
from excelforge.services.backup_service import BackupService
from excelforge.services.format_service import FormatService
from excelforge.services.formula_service import FormulaService
from excelforge.services.operation_service import OperationService
from excelforge.services.range_service import RangeService
from excelforge.services.rollback_service import RollbackService
from excelforge.services.server_service import ServerService
from excelforge.services.sheet_service import SheetService
from excelforge.services.snapshot_service import SnapshotService
from excelforge.services.vba_service import VbaService
from excelforge.services.workbook_service import WorkbookService
from excelforge.services.named_range_service import NamedRangeService
from excelforge.tool_registry import ToolContext, register_all_tools


@dataclass
class ExcelForgeApp:
    config: AppConfig
    mcp: FastMCP
    worker: ExcelWorker
    db: Database
    operation_service: OperationService
    tools_ctx: ToolContext

    def run_stdio(self) -> None:
        self.mcp.run(transport="stdio")

    def shutdown(self) -> None:
        self.worker.stop(wait_seconds=15)


def create_app(config_path: str | None = None) -> ExcelForgeApp:
    config = load_config(config_path)
    db = Database(config)
    db.init_schema()

    worker = ExcelWorker(config)

    audit_repo = AuditRepository(db)
    snapshot_repo = SnapshotRepository(db)
    backup_repo = BackupRepository(db)

    snapshot_service = SnapshotService(config, snapshot_repo)
    backup_service = BackupService(
        config,
        backup_repo,
        workbook_registry=worker.context.registry,
        snapshot_service=snapshot_service,
    )
    workbook_service = WorkbookService(config, worker, snapshot_service)
    sheet_service = SheetService(config, worker, snapshot_service, backup_service)
    range_service = RangeService(config, worker, snapshot_service, backup_service)
    formula_service = FormulaService(config, worker, snapshot_service)
    format_service = FormatService(config, worker)
    vba_service = VbaService(config, worker, backup_service)
    named_range_service = NamedRangeService(config, worker, backup_service)
    rollback_service = RollbackService(config, worker, snapshot_repo, snapshot_service)
    audit_service = AuditService(config, audit_repo)
    cleanup_service = CleanupService(config, audit_repo, snapshot_repo, backup_repo)
    operation_service = OperationService(config, audit_service, cleanup_service)
    operation_service.run_cleanup_on_startup()

    server_service = ServerService(config, worker, snapshot_service, backup_service)

    mcp = FastMCP("ExcelForge")
    ctx = ToolContext(
        server_service=server_service,
        workbook_service=workbook_service,
        sheet_service=sheet_service,
        range_service=range_service,
        formula_service=formula_service,
        format_service=format_service,
        rollback_service=rollback_service,
        snapshot_service=snapshot_service,
        vba_service=vba_service,
        backup_service=backup_service,
        audit_service=audit_service,
        operation_service=operation_service,
        named_range_service=named_range_service,
    )
    tool_registry = register_all_tools(mcp, ctx, config.tools)
    server_service.set_tool_names(tool_registry.get_names())

    return ExcelForgeApp(
        config=config,
        mcp=mcp,
        worker=worker,
        db=db,
        operation_service=operation_service,
        tools_ctx=ctx,
    )


def healthcheck(config_path: str | None = None) -> dict[str, Any]:
    config = load_config(config_path)
    db = Database(config)
    db.init_schema()

    pywin32_ready = False
    try:
        import pythoncom  # type: ignore # noqa: F401
        import win32com.client  # type: ignore # noqa: F401

        pywin32_ready = True
    except Exception:
        pywin32_ready = False

    return {
        "ok": True,
        "server_version": config.server.version,
        "sqlite_path": str(config.sqlite_path),
        "snapshots_dir": str(config.snapshots_dir),
        "backups_dir": str(config.backups_dir),
        "allowed_roots": [str(p) for p in config.allowed_roots],
        "pywin32_ready": pywin32_ready,
    }
