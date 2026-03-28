from __future__ import annotations

import pytest

from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime.workbook_registry import WorkbookHandle, WorkbookRegistry
from excelforge.utils.ids import generate_workbook_id


def _make_handle(workbook_id: str) -> WorkbookHandle:
    return WorkbookHandle(
        workbook_id=workbook_id,
        workbook_name="book.xlsx",
        file_path="D:/ExcelForge/book.xlsx",
        read_only=False,
        opened_at="2026-03-22T00:00:00Z",
        workbook_obj=object(),
    )


def test_registry_generation_bump_invalidates_previous_ids() -> None:
    registry = WorkbookRegistry()
    wb1 = generate_workbook_id(registry.generation)
    registry.add(_make_handle(wb1))

    assert registry.get(wb1) is not None
    assert registry.count() == 1

    new_generation = registry.bump_generation()
    assert new_generation == 2
    assert registry.count() == 0
    assert registry.get(wb1) is None


def test_registry_rejects_mismatched_generation_id() -> None:
    registry = WorkbookRegistry()
    stale_id = generate_workbook_id(registry.generation)
    registry.bump_generation()

    assert registry.get(stale_id) is None
    assert registry.remove(stale_id) is None


def test_registry_rejects_foreign_runtime_workbook_id() -> None:
    registry = WorkbookRegistry(runtime_fingerprint="deadbeef")
    foreign_id = generate_workbook_id(registry.generation, "cafebabe")

    with pytest.raises(ExcelForgeError) as exc_info:
        registry.get(foreign_id)

    assert exc_info.value.code == ErrorCode.E424_WORKBOOK_HANDLE_FOREIGN_RUNTIME
    assert registry.is_stale_workbook_id(foreign_id) is False
