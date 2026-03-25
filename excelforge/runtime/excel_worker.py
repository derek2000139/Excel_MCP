from __future__ import annotations

import queue
import threading
from concurrent.futures import Future, TimeoutError as FutureTimeoutError
from dataclasses import dataclass
from typing import Any, Callable, Generic, TypeVar

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime.excel_app import ExcelAppManager
from excelforge.runtime.retry_policy import run_with_com_retry
from excelforge.runtime.workbook_registry import WorkbookRegistry
from excelforge.utils.timestamps import utc_now_rfc3339

T = TypeVar("T")


@dataclass
class WorkerContext:
    app_manager: ExcelAppManager
    registry: WorkbookRegistry
    worker: ExcelWorker


@dataclass
class WorkerTask(Generic[T]):
    func: Callable[[WorkerContext], T]
    future: Future[T]
    allow_rebuild: bool
    requires_excel: bool


class ExcelWorker:
    def __init__(self, config: AppConfig) -> None:
        self._config = config
        self._queue: queue.Queue[WorkerTask[Any] | None] = queue.Queue()
        self._state = "stopped"
        self._hard_stopped = False
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._last_health_ping: str | None = None
        self._rebuild_count = 0
        self._last_rebuild_at: str | None = None
        self._context = WorkerContext(
            app_manager=ExcelAppManager(config),
            registry=WorkbookRegistry(),
            worker=self,
        )

    @property
    def state(self) -> str:
        with self._lock:
            return self._state

    @property
    def queue_length(self) -> int:
        return self._queue.qsize()

    @property
    def context(self) -> WorkerContext:
        return self._context

    @property
    def generation(self) -> int:
        return self._context.registry.generation

    @property
    def last_health_ping(self) -> str | None:
        return self._last_health_ping

    @property
    def rebuild_count(self) -> int:
        return self._rebuild_count

    @property
    def last_rebuild_at(self) -> str | None:
        return self._last_rebuild_at

    def start(self) -> None:
        with self._lock:
            if self._thread and self._thread.is_alive():
                return
            self._state = "running"
            self._hard_stopped = False
            self._thread = threading.Thread(target=self._run_loop, name="excel-sta-worker", daemon=True)
            self._thread.start()

    def submit(
        self,
        func: Callable[[WorkerContext], T],
        *,
        timeout_seconds: int,
        allow_rebuild: bool = False,
        requires_excel: bool = True,
    ) -> T:
        current_state = self.state
        if requires_excel and current_state == "degraded":
            raise ExcelForgeError(
                ErrorCode.E500_EXCEL_RECOVERING,
                "Excel worker is recovering, please retry shortly",
            )
        if requires_excel and current_state == "stopped" and self._hard_stopped:
            raise ExcelForgeError(
                ErrorCode.E500_EXCEL_UNAVAILABLE,
                "Excel worker is stopped after repeated rebuild failures",
            )

        if current_state == "stopped":
            self.start()

        fut: Future[T] = Future()
        task = WorkerTask(
            func=func,
            future=fut,
            allow_rebuild=allow_rebuild,
            requires_excel=requires_excel,
        )
        self._queue.put(task)

        try:
            return fut.result(timeout=timeout_seconds)
        except FutureTimeoutError as exc:
            raise ExcelForgeError(
                ErrorCode.E500_OPERATION_TIMEOUT,
                f"Operation timed out after {timeout_seconds} seconds",
            ) from exc

    def stop(self, wait_seconds: int = 15) -> None:
        with self._lock:
            thread = self._thread
            if thread is None:
                self._state = "stopped"
                return
            self._state = "stopped"
        self._queue.put(None)
        thread.join(timeout=wait_seconds)

    def _set_state(self, state: str) -> None:
        with self._lock:
            self._state = state

    def _run_loop(self) -> None:
        pythoncom = None
        try:
            import pythoncom as _pythoncom  # type: ignore

            pythoncom = _pythoncom
            pythoncom.CoInitialize()
        except Exception:
            self._set_state("degraded")

        try:
            while True:
                task = self._queue.get()
                if task is None:
                    break

                if task.future.cancelled():
                    continue

                try:
                    result = self._execute_task(task)
                except Exception as exc:  # noqa: BLE001
                    task.future.set_exception(exc)
                else:
                    task.future.set_result(result)
        finally:
            self._context.app_manager.close()
            self._context.registry.invalidate_all()
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _execute_task(self, task: WorkerTask[T]) -> T:
        if task.requires_excel:
            if self._config.excel.health_ping_enabled and not self._health_ping():
                self._set_state("degraded")
                recovered = self._recover_excel_instance()
                if not recovered:
                    raise ExcelForgeError(
                        ErrorCode.E500_EXCEL_UNAVAILABLE,
                        "Excel instance unavailable after rebuild attempts",
                    )
                if not task.allow_rebuild:
                    raise ExcelForgeError(
                        ErrorCode.E410_WORKBOOK_STALE,
                        "Workbook handle is stale after Excel rebuild; reopen with workbook.open_file",
                    )
            else:
                self._context.app_manager.ensure_app()
                if self.state != "running":
                    self._set_state("running")

        try:
            return run_with_com_retry(lambda: task.func(self._context))
        except ExcelForgeError as exc:
            if exc.code != ErrorCode.E500_COM_DISCONNECTED:
                raise

            self._set_state("degraded")
            recovered = self._recover_excel_instance()
            if not recovered:
                raise ExcelForgeError(
                    ErrorCode.E500_EXCEL_UNAVAILABLE,
                    "Excel instance unavailable after rebuild attempts",
                ) from exc

            if not task.allow_rebuild:
                raise ExcelForgeError(
                    ErrorCode.E410_WORKBOOK_STALE,
                    "Workbook handle is stale after Excel rebuild; reopen with workbook.open_file",
                ) from exc

            return run_with_com_retry(lambda: task.func(self._context))

    def _health_ping(self) -> bool:
        self._last_health_ping = utc_now_rfc3339()
        if not self._context.app_manager.ready:
            try:
                self._context.app_manager.ensure_app()
                return True
            except Exception:
                return False
        return self._context.app_manager.ping()

    def _recover_excel_instance(self) -> bool:
        self._invalidate_excel_session()
        attempts = int(self._config.excel.max_rebuild_attempts)
        for _ in range(attempts):
            try:
                self._context.app_manager.ensure_app()
            except Exception:
                continue
            self._rebuild_count += 1
            self._last_rebuild_at = utc_now_rfc3339()
            self._hard_stopped = False
            self._set_state("running")
            return True

        self._hard_stopped = True
        self._set_state("stopped")
        return False

    def _invalidate_excel_session(self) -> None:
        self._context.registry.bump_generation()
        self._context.app_manager.invalidate()

    def mark_unhealthy(self, reason: str) -> None:
        self._set_state("degraded")
        self._recover_excel_instance()
