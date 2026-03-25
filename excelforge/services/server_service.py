from __future__ import annotations

from typing import Any

from excelforge.config import AppConfig
from excelforge.runtime.excel_worker import ExcelWorker
from excelforge.services.backup_service import BackupService
from excelforge.services.snapshot_service import SnapshotService


class ServerService:
    def __init__(
        self,
        config: AppConfig,
        worker: ExcelWorker,
        snapshot_service: SnapshotService,
        backup_service: BackupService,
    ) -> None:
        self._config = config
        self._worker = worker
        self._snapshot_service = snapshot_service
        self._backup_service = backup_service
        self._tool_names: list[str] = []

    def set_tool_names(self, names: list[str]) -> None:
        self._tool_names = names

    def get_status(self) -> dict[str, Any]:
        snapshot_stats_raw = self._snapshot_service.get_stats(workbook_id=None)
        backup_stats_raw = self._backup_service.get_stats()
        return {
            "server_version": self._config.server.version,
            "mode": "desktop_agent_hidden_excel",
            "debug_mode": bool(self._config.excel.visible),
            "config_reloadable": False,
            "supported_extensions": self._config.paths.allowed_extensions,
            "excel_worker": {
                "state": self._worker.state,
                "queue_length": self._worker.queue_length,
                "excel_ready": self._worker.context.app_manager.ready,
                "last_health_ping": self._worker.last_health_ping,
                "rebuild_count": self._worker.rebuild_count,
                "last_rebuild_at": self._worker.last_rebuild_at,
            },
            "open_workbooks": self._worker.context.registry.count(),
            "limits": {
                "max_open_workbooks": self._config.limits.max_open_workbooks,
                "max_read_cells": self._config.limits.max_read_cells,
                "max_write_cells": self._config.limits.max_write_cells,
                "max_snapshot_cells": self._config.limits.max_snapshot_cells,
                "max_insert_rows": self._config.limits.max_insert_rows,
                "max_insert_columns": self._config.limits.max_insert_columns,
                "max_vba_code_size_bytes": self._config.limits.max_vba_code_size_bytes,
                "default_read_rows": self._config.limits.default_read_rows,
                "max_read_rows": self._config.limits.max_read_rows,
                "operation_timeout_seconds": self._config.limits.operation_timeout_seconds,
            },
            "snapshot_stats": {
                "active_count": snapshot_stats_raw["active_snapshots"],
                "expired_count": snapshot_stats_raw["expired_snapshots"],
                "total_size_bytes": snapshot_stats_raw["total_size_bytes"],
                "oldest_active_at": snapshot_stats_raw["oldest_active_at"],
            },
            "backup_stats": backup_stats_raw,
            "capabilities": self._tool_names,
        }
