from __future__ import annotations

import pytest
from pydantic import ValidationError

from excelforge.models.formula_models import FormulaFillRangeRequest
from excelforge.models.range_models import (
    RangeInsertColumnsRequest,
    RangeInsertRowsRequest,
    RangeReadValuesRequest,
    RangeWriteValuesRequest,
)
from excelforge.models.sheet_models import SheetDeleteSheetRequest


def test_range_read_schema_defaults() -> None:
    req = RangeReadValuesRequest(workbook_id="wb_1", sheet_name="Sheet1", range="A1:C10")
    assert req.value_mode == "raw"
    assert req.row_limit == 200


def test_range_write_requires_rectangular_values() -> None:
    with pytest.raises(ValidationError):
        RangeWriteValuesRequest(
            workbook_id="wb_1",
            sheet_name="Sheet1",
            start_cell="A1",
            values=[[1, 2], [3]],
        )


def test_formula_fill_schema_rejects_invalid_range() -> None:
    with pytest.raises(ValidationError):
        FormulaFillRangeRequest(
            workbook_id="wb_1",
            sheet_name="Sheet1",
            range="A:A",
            formula="=SUM(A1:A3)",
        )


def test_range_insert_schema_rejects_invalid_column() -> None:
    with pytest.raises(ValidationError):
        RangeInsertColumnsRequest(
            workbook_id="wb_1",
            sheet_name="Sheet1",
            column="A1",
            count=1,
        )


def test_range_insert_rows_schema_defaults() -> None:
    req = RangeInsertRowsRequest(workbook_id="wb_1", sheet_name="Sheet1", row_number=2)
    assert req.count == 1


def test_sheet_delete_requires_confirm_token() -> None:
    with pytest.raises(ValidationError):
        SheetDeleteSheetRequest(
            workbook_id="wb_1",
            sheet_name="Sheet1",
            confirm_token="",
        )
