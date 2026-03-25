from __future__ import annotations

import pytest

from excelforge.models.error_models import ExcelForgeError
from excelforge.utils.address_parser import (
    CellRef,
    RangeRef,
    cell_to_a1,
    index_to_column,
    parse_cell,
    parse_range,
    range_to_a1,
)


def test_parse_cell_and_back() -> None:
    cell = parse_cell("$B$12")
    assert cell == CellRef(row=12, col=2)
    assert cell_to_a1(cell) == "B12"


def test_parse_range_and_back() -> None:
    rr = parse_range("A1:C3")
    assert rr == RangeRef(start=CellRef(1, 1), end=CellRef(3, 3))
    assert rr.cell_count == 9
    assert range_to_a1(rr) == "A1:C3"


def test_index_to_column() -> None:
    assert index_to_column(1) == "A"
    assert index_to_column(26) == "Z"
    assert index_to_column(27) == "AA"


def test_invalid_range_raises() -> None:
    with pytest.raises(ExcelForgeError):
        parse_range("C3:A1")
