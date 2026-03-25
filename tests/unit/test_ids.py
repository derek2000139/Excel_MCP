from __future__ import annotations

from excelforge.utils.ids import generate_workbook_id, parse_workbook_generation


def test_workbook_id_contains_generation() -> None:
    workbook_id = generate_workbook_id(7)
    assert workbook_id.startswith("wb_g7_")
    assert parse_workbook_generation(workbook_id) == 7


def test_parse_workbook_generation_invalid() -> None:
    assert parse_workbook_generation("wb_legacy_123") is None
