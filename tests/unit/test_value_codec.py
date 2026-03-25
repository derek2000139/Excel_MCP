from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal

import pytest

from excelforge.models.error_models import ExcelForgeError
from excelforge.utils.value_codec import ensure_rectangular, matrix_to_json, to_scalar


def test_to_scalar_basic_types() -> None:
    assert to_scalar(1) == 1
    assert to_scalar(1.5) == 1.5
    assert to_scalar(True) is True
    assert to_scalar(None) is None
    assert to_scalar(Decimal("1.25")) == 1.25
    assert isinstance(to_scalar(datetime(2026, 1, 1)), str)
    assert isinstance(to_scalar(date(2026, 1, 1)), str)


def test_matrix_to_json_tuple_input() -> None:
    result = matrix_to_json(((1, 2), (3, None)))
    assert result == [[1, 2], [3, None]]


def test_ensure_rectangular_rejects_non_rectangular() -> None:
    with pytest.raises(ExcelForgeError):
        ensure_rectangular([[1, 2], [3]])
