from __future__ import annotations

from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.utils.ids import is_same_runtime_fingerprint, parse_workbook_fingerprint


def is_foreign_workbook_id(workbook_id: str, runtime_fingerprint: str | None) -> bool:
    workbook_fingerprint = parse_workbook_fingerprint(workbook_id)
    if runtime_fingerprint is None or workbook_fingerprint is None:
        return False
    return not is_same_runtime_fingerprint(workbook_fingerprint, runtime_fingerprint)


def ensure_workbook_id_owned(workbook_id: str, runtime_fingerprint: str | None) -> None:
    workbook_fingerprint = parse_workbook_fingerprint(workbook_id)
    if runtime_fingerprint is None or workbook_fingerprint is None:
        return
    if is_same_runtime_fingerprint(workbook_fingerprint, runtime_fingerprint):
        return
    raise ExcelForgeError(
        ErrorCode.E424_WORKBOOK_HANDLE_FOREIGN_RUNTIME,
        "Workbook handle belongs to another Runtime; reopen the workbook in the current Host/Runtime",
        details={
            "workbook_id": workbook_id,
            "workbook_runtime_fingerprint": workbook_fingerprint,
            "runtime_fingerprint": runtime_fingerprint,
        },
    )


def ensure_related_handle_owned(
    *,
    handle_kind: str,
    handle_id: str,
    owner_workbook_id: str,
    runtime_fingerprint: str | None,
) -> None:
    owner_fingerprint = parse_workbook_fingerprint(owner_workbook_id)
    if runtime_fingerprint is None or owner_fingerprint is None:
        return
    if is_same_runtime_fingerprint(owner_fingerprint, runtime_fingerprint):
        return
    raise ExcelForgeError(
        ErrorCode.E424_HANDLE_RUNTIME_MISMATCH,
        f"{handle_kind.capitalize()} handle belongs to another Runtime and cannot be used in the current Runtime",
        details={
            "handle_kind": handle_kind,
            "handle_id": handle_id,
            "owner_workbook_id": owner_workbook_id,
            "owner_runtime_fingerprint": owner_fingerprint,
            "runtime_fingerprint": runtime_fingerprint,
        },
    )
