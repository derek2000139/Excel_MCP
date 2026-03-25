from __future__ import annotations

from excelforge.models.common import error_envelope, ok_envelope
from excelforge.models.error_models import ErrorCode


def test_ok_envelope_shape() -> None:
    env = ok_envelope(
        tool_name="range.write_values",
        operation_id="op_1",
        duration_ms=10,
        server_version="0.1.0",
        data={"x": 1},
        workbook_id="wb_1",
        snapshot_id="snap_1",
        rollback_supported=True,
    )
    payload = env.model_dump(mode="json")
    assert payload["success"] is True
    assert payload["code"] == "OK"
    assert payload["meta"]["operation_id"] == "op_1"


def test_error_envelope_shape() -> None:
    env = error_envelope(
        tool_name="range.write_values",
        operation_id="op_1",
        duration_ms=10,
        server_version="0.1.0",
        code=ErrorCode.E500_INTERNAL,
        message="boom",
    )
    payload = env.model_dump(mode="json")
    assert payload["success"] is False
    assert payload["data"] is None
    assert payload["code"] == ErrorCode.E500_INTERNAL
