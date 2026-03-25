from __future__ import annotations

from pathlib import Path
from typing import Any

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError


class ExcelAppManager:
    def __init__(self, config: AppConfig) -> None:
        self._config = config
        self._app: Any | None = None

    @property
    def ready(self) -> bool:
        return self._app is not None

    def ensure_app(self) -> Any:
        if self._app is None:
            self._app = self._create_app()
        if self._config.excel.ensure_visibility:
            self._app.Visible = self._config.excel.visible
        return self._app

    def _create_app(self) -> Any:
        try:
            import win32com.client  # type: ignore
        except Exception as exc:  # pragma: no cover - non-Windows environment
            raise ExcelForgeError(
                ErrorCode.E500_EXCEL_UNAVAILABLE,
                f"win32com is unavailable: {exc}",
            ) from exc

        try:
            app = win32com.client.DispatchEx("Excel.Application")
        except Exception as exc:  # pragma: no cover - requires Excel Desktop
            raise ExcelForgeError(
                ErrorCode.E500_EXCEL_UNAVAILABLE,
                f"Failed to create hidden Excel instance: {exc}",
            ) from exc

        visible = bool(self._config.excel.visible)
        app.Visible = visible
        app.DisplayAlerts = not bool(self._config.excel.disable_alerts)
        app.ScreenUpdating = visible
        app.EnableEvents = not bool(self._config.excel.disable_events)
        app.AskToUpdateLinks = False

        try:
            app.AutomationSecurity = 1
        except Exception:
            pass

        return app

    def ping(self) -> bool:
        if self._app is None:
            return False
        try:
            _ = self._app.Workbooks.Count
            return True
        except Exception:
            return False

    def invalidate(self) -> None:
        self._app = None

    def close(self) -> None:
        if self._app is None:
            return
        try:
            self._app.Quit()
        except Exception:
            pass
        finally:
            self._app = None

    def open_workbook(self, file_path: Path, read_only: bool) -> Any:
        app = self.ensure_app()
        wb = app.Workbooks.Open(str(file_path), UpdateLinks=0, ReadOnly=read_only)
        if self._config.excel.ensure_visibility:
            try:
                wb.Windows(1).Visible = True
            except Exception:
                pass
        return wb
