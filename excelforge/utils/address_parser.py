from __future__ import annotations

import re
from dataclasses import dataclass

from excelforge.models.error_models import ErrorCode, ExcelForgeError

CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$")
RANGE_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?$")


@dataclass(frozen=True)
class CellRef:
    row: int
    col: int


@dataclass(frozen=True)
class RangeRef:
    start: CellRef
    end: CellRef

    @property
    def rows(self) -> int:
        return self.end.row - self.start.row + 1

    @property
    def cols(self) -> int:
        return self.end.col - self.start.col + 1

    @property
    def cell_count(self) -> int:
        return self.rows * self.cols


def column_to_index(column: str) -> int:
    idx = 0
    for ch in column.upper():
        if ch < "A" or ch > "Z":
            raise ExcelForgeError(ErrorCode.E400_RANGE_INVALID, f"Invalid column: {column}")
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def index_to_column(index: int) -> str:
    if index <= 0:
        raise ExcelForgeError(ErrorCode.E400_RANGE_INVALID, f"Invalid column index: {index}")
    chars: list[str] = []
    current = index
    while current > 0:
        current -= 1
        chars.append(chr((current % 26) + ord("A")))
        current //= 26
    return "".join(reversed(chars))


def parse_cell(address: str) -> CellRef:
    m = CELL_RE.match(address)
    if not m:
        raise ExcelForgeError(ErrorCode.E400_RANGE_INVALID, f"Invalid A1 cell address: {address}")
    col = column_to_index(m.group(1))
    row = int(m.group(2))
    if row <= 0:
        raise ExcelForgeError(ErrorCode.E400_RANGE_INVALID, f"Invalid row number: {row}")
    return CellRef(row=row, col=col)


def parse_cell_address(address: str) -> CellRef:
    return parse_cell(address)


def parse_range(address: str) -> RangeRef:
    m = RANGE_RE.match(address)
    if not m:
        raise ExcelForgeError(ErrorCode.E400_RANGE_INVALID, f"Invalid A1 range address: {address}")
    start_col = column_to_index(m.group(1))
    start_row = int(m.group(2))
    end_col = column_to_index(m.group(3) or m.group(1))
    end_row = int(m.group(4) or m.group(2))
    if start_row <= 0 or end_row <= 0:
        raise ExcelForgeError(ErrorCode.E400_RANGE_INVALID, "Row index must be positive")
    if end_row < start_row or end_col < start_col:
        raise ExcelForgeError(ErrorCode.E400_RANGE_INVALID, "Range must be top-left to bottom-right")
    return RangeRef(start=CellRef(start_row, start_col), end=CellRef(end_row, end_col))


def cell_to_a1(cell: CellRef) -> str:
    return f"{index_to_column(cell.col)}{cell.row}"


def range_to_a1(rng: RangeRef) -> str:
    start = cell_to_a1(rng.start)
    end = cell_to_a1(rng.end)
    if start == end:
        return start
    return f"{start}:{end}"


def shifted_row_page(source: RangeRef, row_offset: int, row_limit: int) -> RangeRef:
    if row_offset >= source.rows:
        empty_row = source.start.row + source.rows
        return RangeRef(CellRef(empty_row, source.start.col), CellRef(empty_row, source.end.col))
    page_start = source.start.row + row_offset
    page_end = min(source.end.row, page_start + row_limit - 1)
    return RangeRef(CellRef(page_start, source.start.col), CellRef(page_end, source.end.col))
