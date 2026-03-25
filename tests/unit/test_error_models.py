from __future__ import annotations

from excelforge.models.error_models import ErrorCode, ExcelForgeError, normalize_exception


def test_normalize_exception_passthrough() -> None:
    exc = ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, "bad")
    out = normalize_exception(exc)
    assert out.code == ErrorCode.E400_INVALID_ARGUMENT


def test_normalize_exception_wraps_unknown() -> None:
    out = normalize_exception(RuntimeError("x"))
    assert out.code == ErrorCode.E500_INTERNAL


def test_v03_error_codes_exist() -> None:
    assert ErrorCode.E400_ROW_OUT_OF_RANGE.value == "E400_ROW_OUT_OF_RANGE"
    assert ErrorCode.E400_COLUMN_OUT_OF_RANGE.value == "E400_COLUMN_OUT_OF_RANGE"
    assert ErrorCode.E403_VBA_ACCESS_DENIED.value == "E403_VBA_ACCESS_DENIED"
    assert ErrorCode.E403_VBA_PROJECT_PROTECTED.value == "E403_VBA_PROJECT_PROTECTED"
    assert ErrorCode.E404_VBA_MODULE_NOT_FOUND.value == "E404_VBA_MODULE_NOT_FOUND"
    assert ErrorCode.E409_CANNOT_DELETE_LAST_SHEET.value == "E409_CANNOT_DELETE_LAST_SHEET"
    assert ErrorCode.E409_CONFIRM_TOKEN_INVALID.value == "E409_CONFIRM_TOKEN_INVALID"
    assert ErrorCode.E500_BACKUP_FAILED.value == "E500_BACKUP_FAILED"
