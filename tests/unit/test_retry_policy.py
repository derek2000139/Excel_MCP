from __future__ import annotations

from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime import retry_policy


def test_run_with_com_retry_raises_disconnected_code() -> None:
    def fn() -> None:
        raise RuntimeError("(-2147417848, '被调用的对象已与其客户端断开连接。', None, None)")

    try:
        retry_policy.run_with_com_retry(fn)
    except ExcelForgeError as exc:
        assert exc.code == ErrorCode.E500_COM_DISCONNECTED
    else:
        raise AssertionError("Expected ExcelForgeError")


def test_run_with_com_retry_retries_rejected_then_succeeds(monkeypatch) -> None:
    monkeypatch.setattr(retry_policy, "BACKOFF_SECONDS", (0.0, 0.0, 0.0))
    attempts = {"n": 0}

    def fn() -> str:
        attempts["n"] += 1
        if attempts["n"] < 3:
            raise RuntimeError("Call was rejected by callee")
        return "ok"

    assert retry_policy.run_with_com_retry(fn) == "ok"
    assert attempts["n"] == 3
