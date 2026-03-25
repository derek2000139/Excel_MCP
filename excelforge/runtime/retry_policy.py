from __future__ import annotations

import time
from collections.abc import Callable
from typing import TypeVar

from excelforge.models.error_models import ErrorCode, ExcelForgeError

T = TypeVar("T")

BACKOFF_SECONDS = (0.1, 0.3, 0.5, 1.0, 1.5)
_REJECTED_HRESULTS = {-2147418111, -2147417846}
_DISCONNECTED_HRESULTS = {-2147417848}
_SERVER_EXEC_FAILURE_HRESULTS = {-2146777998}


try:
    import pywintypes  # type: ignore
except Exception:  # pragma: no cover - non-Windows environment
    pywintypes = None


def _is_com_rejected(exc: Exception) -> bool:
    if pywintypes is not None and isinstance(exc, pywintypes.com_error):
        hresult = exc.hresult if hasattr(exc, "hresult") else None
        if hresult in _REJECTED_HRESULTS:
            return True
        msg = str(exc).lower()
        return "call was rejected by callee" in msg or "server busy" in msg
    msg = str(exc).lower()
    return "call was rejected by callee" in msg or "server busy" in msg


def _is_server_exec_failure(exc: Exception) -> bool:
    if pywintypes is None or not isinstance(exc, pywintypes.com_error):
        return False
    hresult = exc.hresult if hasattr(exc, "hresult") else None
    return hresult in _SERVER_EXEC_FAILURE_HRESULTS


def _is_unknown_com_error(exc: Exception) -> bool:
    if pywintypes is None:
        return False
    return isinstance(exc, pywintypes.com_error) and not _is_com_rejected(exc) and not _is_com_disconnected(exc)


def _is_com_disconnected(exc: Exception) -> bool:
    if pywintypes is not None and isinstance(exc, pywintypes.com_error):
        hresult = exc.hresult if hasattr(exc, "hresult") else None
        if hresult in _DISCONNECTED_HRESULTS:
            return True
    msg = str(exc).lower()
    return "disconnected from its clients" in msg or "对象已与其客户端断开连接" in msg


def run_with_com_retry(fn: Callable[[], T]) -> T:
    last_exc: Exception | None = None
    for idx, delay in enumerate(BACKOFF_SECONDS, start=1):
        try:
            return fn()
        except Exception as exc:  # noqa: BLE001
            if _is_com_disconnected(exc):
                raise ExcelForgeError(
                    ErrorCode.E500_COM_DISCONNECTED,
                    f"Excel COM object disconnected: {exc}",
                ) from exc
            if _is_server_exec_failure(exc):
                raise ExcelForgeError(
                    ErrorCode.E500_EXCEL_UNAVAILABLE,
                    f"Excel server execution failure: {exc}",
                ) from exc
            if not _is_com_rejected(exc):
                if _is_unknown_com_error(exc) and idx < 2:
                    time.sleep(delay)
                    continue
                raise
            last_exc = exc
            if idx == len(BACKOFF_SECONDS):
                break
            time.sleep(delay)
    raise ExcelForgeError(
        ErrorCode.E500_COM_REJECTED,
        f"Excel COM call rejected after retries: {last_exc}",
    )
