from __future__ import annotations

import pytest

from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime.handle_ownership import ensure_related_handle_owned, ensure_workbook_id_owned
from excelforge.utils.ids import generate_workbook_id


def test_ensure_workbook_id_owned_allows_legacy_workbook_id() -> None:
    legacy_workbook_id = generate_workbook_id(1)
    ensure_workbook_id_owned(legacy_workbook_id, "deadbeef")


def test_ensure_workbook_id_owned_rejects_foreign_runtime() -> None:
    foreign_workbook_id = generate_workbook_id(3, "cafebabe")

    with pytest.raises(ExcelForgeError) as exc_info:
        ensure_workbook_id_owned(foreign_workbook_id, "deadbeef")

    assert exc_info.value.code == ErrorCode.E424_WORKBOOK_HANDLE_FOREIGN_RUNTIME


def test_ensure_related_handle_owned_rejects_foreign_runtime_handle() -> None:
    owner_workbook_id = generate_workbook_id(2, "cafebabe")

    with pytest.raises(ExcelForgeError) as exc_info:
        ensure_related_handle_owned(
            handle_kind="snapshot",
            handle_id="snap_123",
            owner_workbook_id=owner_workbook_id,
            runtime_fingerprint="deadbeef",
        )

    assert exc_info.value.code == ErrorCode.E424_HANDLE_RUNTIME_MISMATCH
