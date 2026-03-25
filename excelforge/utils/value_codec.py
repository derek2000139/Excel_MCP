from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from typing import Any

from excelforge.models.error_models import ErrorCode, ExcelForgeError

ScalarValue = str | int | float | bool | None


def to_scalar(value: Any) -> ScalarValue:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float, str)):
        return value
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def matrix_to_json(values: Any) -> list[list[ScalarValue]]:
    if values is None:
        return [[None]]
    if isinstance(values, tuple):
        rows = [list(row) if isinstance(row, tuple) else [row] for row in values]
    elif isinstance(values, list):
        rows = [list(row) if isinstance(row, (tuple, list)) else [row] for row in values]
    else:
        rows = [[values]]

    return [[to_scalar(cell) for cell in row] for row in rows]


def ensure_rectangular(values: list[list[ScalarValue]]) -> tuple[int, int]:
    if not values:
        raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, "values cannot be empty")
    width = len(values[0])
    if width == 0:
        raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, "values rows cannot be empty")
    for row in values:
        if len(row) != width:
            raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, "values must be rectangular")
    return len(values), width


def to_excel_matrix(values: list[list[ScalarValue]]) -> tuple[tuple[ScalarValue, ...], ...]:
    ensure_rectangular(values)
    return tuple(tuple(cell for cell in row) for row in values)
