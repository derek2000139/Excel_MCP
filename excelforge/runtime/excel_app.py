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
        import logging
        logger = logging.getLogger(__name__)

        if self._app is None:
            logger.info(f"[ExcelApp] Creating NEW Excel app (was None)")
            self._app = self._create_app()
        else:
            if not self._is_app_valid():
                logger.info(f"[ExcelApp] Existing Excel app is stale, creating new one")
                self._app = None
                self._app = self._create_app()
            else:
                logger.info(f"[ExcelApp] Reusing existing Excel app")

        if self._config.excel.ensure_visibility:
            workbooks_open = self._count_open_workbooks()
            if workbooks_open > 0:
                self._app.Visible = self._config.excel.visible
            elif not self._config.excel.visible:
                self._app.Visible = False
        return self._app

    def _is_app_valid(self) -> bool:
        """检查 Excel COM 对象是否仍然有效。"""
        if self._app is None:
            return False
        try:
            _ = self._app.Workbooks.Count
            return True
        except Exception:
            return False

    def _count_open_workbooks(self) -> int:
        """获取当前打开的工作簿数量。"""
        try:
            return self._app.Workbooks.Count if self._app else 0
        except Exception:
            return 0

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

        app.Visible = False
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        app.EnableEvents = not bool(self._config.excel.disable_events)
        app.AskToUpdateLinks = False

        try:
            app.AutomationSecurity = 1
        except Exception:
            pass

        self._close_blank_workbooks(app)

        try:
            from excelforge.runtime.com_utils import get_excel_pid
            pid = get_excel_pid(app)
            import logging
            logging.getLogger(__name__).info(f"[ExcelApp] Created hidden Excel PID={pid}")
        except Exception:
            pass

        return app

    def _close_blank_workbooks(self, app: Any) -> None:
        """关闭 Excel 启动时自动创建的空白工作簿。"""
        try:
            count = app.Workbooks.Count
            to_close = []
            for i in range(1, count + 1):
                try:
                    wb = app.Workbooks(i)
                    name = wb.Name
                    if name in ("工作簿1", "Book1", "Book", "Sheet1"):
                        to_close.append(wb)
                except Exception:
                    continue

            for wb in to_close:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass

            if to_close:
                import logging
                logging.getLogger(__name__).info(
                    f"[ExcelApp] Closed {len(to_close)} blank workbook(s)"
                )
        except Exception:
            pass

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
