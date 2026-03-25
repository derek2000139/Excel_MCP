from __future__ import annotations

from typing import Any

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime.excel_worker import ExcelWorker
from excelforge.services.backup_service import BackupService


class NamedRangeService:
    def __init__(self, config: AppConfig, worker: ExcelWorker, backup_service: BackupService) -> None:
        self._config = config
        self._worker = worker
        self._backup_service = backup_service

    def list_ranges(
        self,
        *,
        workbook_id: str,
        scope: str = "all",
        sheet_name: str | None = None,
    ) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            wb = handle.workbook_obj
            items: list[dict] = []

            if scope in ("workbook", "all"):
                for nr in wb.Names:
                    try:
                        ref = str(nr.RefersTo) if nr.RefersTo else ""
                        items.append({
                            "name": str(nr.Name),
                            "scope": "workbook",
                            "sheet_name": None,
                            "refers_to": ref,
                            "refers_to_type": self._classify_ref(ref),
                            "visible": bool(nr.Visible),
                            "address": self._resolve_address(wb, nr),
                        })
                    except Exception:
                        continue

            if scope in ("worksheet", "all") and sheet_name:
                try:
                    ws = wb.Worksheets(sheet_name)
                    for nr in ws.Names:
                        try:
                            ref = str(nr.RefersTo) if nr.RefersTo else ""
                            items.append({
                                "name": str(nr.Name),
                                "scope": "worksheet",
                                "sheet_name": sheet_name,
                                "refers_to": ref,
                                "refers_to_type": self._classify_ref(ref),
                                "visible": bool(nr.Visible),
                                "address": self._resolve_address(ws, nr),
                            })
                        except Exception:
                            continue
                except Exception:
                    pass

            return {
                "total": len(items),
                "items": items,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def read_values(
        self,
        *,
        workbook_id: str,
        range_name: str,
        value_mode: str,
        row_offset: int,
        row_limit: int,
    ) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            wb = handle.workbook_obj
            nr = None
            for n in wb.Names:
                if str(n.Name) == range_name:
                    nr = n
                    break

            if nr is None:
                for ws in wb.Worksheets:
                    for n in ws.Names:
                        if str(n.Name) == range_name:
                            nr = n
                            break
                    if nr is not None:
                        break

            if nr is None:
                raise ExcelForgeError(ErrorCode.E404_NAMED_RANGE_NOT_FOUND, f"Named range not found: {range_name}")

            ref = str(nr.RefersTo) if nr.RefersTo else ""
            ref_type = self._classify_ref(ref)
            if ref_type != "reference":
                raise ExcelForgeError(
                    ErrorCode.E423_NAMED_RANGE_NOT_A_REFERENCE,
                    f"Named range '{range_name}' is a {ref_type}, not a cell reference",
                )

            try:
                rng = nr.RefersToRange
            except Exception as exc:
                raise ExcelForgeError(
                    ErrorCode.E404_NAMED_RANGE_NOT_FOUND,
                    f"Cannot resolve named range {range_name}: {exc}",
                ) from exc

            sheet_name = str(rng.Worksheet.Name) if rng.Worksheet else None
            total_rows = int(rng.Rows.Count)
            total_cols = int(rng.Columns.Count)
            actual_rows = min(row_limit, max(0, total_rows - row_offset))
            returned_rows = 0
            values: list[list[Any]] = []

            if actual_rows > 0:
                for r in range(row_offset + 1, row_offset + actual_rows + 1):
                    row_values: list[Any] = []
                    for c in range(1, total_cols + 1):
                        cell = rng.Cells(r, c)
                        val = str(cell.Text) if value_mode == "display" else (cell.Value2 if hasattr(cell, "Value2") else None)
                        row_values.append(val)
                    values.append(row_values)
                    returned_rows += 1

            return {
                "range_name": range_name,
                "scope": "workbook",
                "sheet_name": sheet_name,
                "refers_to": ref,
                "resolved_address": str(rng.Address) if rng else None,
                "total_rows": total_rows,
                "returned_rows": returned_rows,
                "column_count": total_cols,
                "values": values,
                "truncated": (row_offset + returned_rows) < total_rows,
                "next_row_offset": (row_offset + returned_rows) if (row_offset + returned_rows) < total_rows else None,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def create_range(
        self,
        *,
        workbook_id: str,
        name: str,
        refers_to: str,
        scope: str = "workbook",
        sheet_name: str | None = None,
        overwrite: bool = False,
    ) -> dict[str, Any]:
        if not refers_to.startswith("="):
            raise ExcelForgeError(ErrorCode.E423_NAMED_RANGE_NOT_A_REFERENCE, "refers_to must start with '='")

        if scope not in ("workbook", "worksheet"):
            raise ExcelForgeError(ErrorCode.E400_INVALID_SCOPE, "scope must be 'workbook' or 'worksheet'")

        if scope == "worksheet" and not sheet_name:
            raise ExcelForgeError(ErrorCode.E400_MISSING_SHEET_NAME, "sheet_name required when scope is 'worksheet'")

        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            wb = handle.workbook_obj

            existing_name = None
            if scope == "workbook":
                for n in wb.Names:
                    if str(n.Name) == name:
                        existing_name = n
                        break
            else:
                try:
                    ws = wb.Worksheets(sheet_name)
                    for n in ws.Names:
                        if str(n.Name) == name:
                            existing_name = n
                            break
                except Exception:
                    raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, f"Sheet not found: {sheet_name}") from None

            if existing_name is not None and not overwrite:
                raise ExcelForgeError(
                    ErrorCode.E409_NAMED_RANGE_EXISTS,
                    f"Named range '{name}' already exists. Use overwrite=true to replace.",
                )

            backup_id, _ = self._backup_service.create_backup(
                workbook=handle,
                source_tool="named_range.create_range",
                description=f"Create or update named range {name}",
            )

            try:
                if scope == "workbook":
                    if existing_name is not None:
                        existing_name.Delete()
                    wb.Names.Add(Name=name, RefersToR1C1=refers_to)
                else:
                    ws = wb.Worksheets(sheet_name)
                    if existing_name is not None:
                        existing_name.Delete()
                    ws.Names.Add(Name=name, RefersToR1C1=refers_to)

                action = "updated" if existing_name is not None else "created"

                return {
                    "name": name,
                    "scope": scope,
                    "sheet_name": sheet_name,
                    "refers_to": refers_to,
                    "action": action,
                    "backup_id": backup_id,
                }
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to create named range: {exc}") from exc

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def delete_range(
        self,
        *,
        workbook_id: str,
        name: str,
        scope: str = "workbook",
        sheet_name: str | None = None,
    ) -> dict[str, Any]:
        if scope not in ("workbook", "worksheet"):
            raise ExcelForgeError(ErrorCode.E400_INVALID_SCOPE, "scope must be 'workbook' or 'worksheet'")

        if scope == "worksheet" and not sheet_name:
            raise ExcelForgeError(ErrorCode.E400_MISSING_SHEET_NAME, "sheet_name required when scope is 'worksheet'")

        RESERVED_NAMES = {
            "Print_Area", "Print_Titles", "Criteria", "_FilterDatabase",
            "Sheet_Titles", " consolidation_Area", " consolidationAreas_",
        }

        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            wb = handle.workbook_obj
            target_name = None

            if scope == "workbook":
                for n in wb.Names:
                    if str(n.Name) == name:
                        target_name = n
                        break
            else:
                try:
                    ws = wb.Worksheets(sheet_name)
                    for n in ws.Names:
                        if str(n.Name) == name:
                            target_name = n
                            break
                except Exception:
                    raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, f"Sheet not found: {sheet_name}") from None

            if target_name is None:
                raise ExcelForgeError(ErrorCode.E404_NAMED_RANGE_NOT_FOUND, f"Named range not found: {name}")

            if name in RESERVED_NAMES:
                raise ExcelForgeError(
                    ErrorCode.E403_FEATURE_NOT_ALLOWED,
                    f"Cannot delete reserved named range: {name}",
                )

            backup_id, _ = self._backup_service.create_backup(
                workbook=handle,
                source_tool="named_range.delete_range",
                description=f"Delete named range {name}",
            )

            try:
                target_name.Delete()
                deleted = True
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to delete named range: {exc}") from exc

            return {
                "name": name,
                "scope": scope,
                "deleted": deleted,
                "backup_id": backup_id,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def manage(
        self,
        *,
        workbook_id: str,
        action: str,
        name: str | None = None,
        refers_to: str | None = None,
        scope: str = "workbook",
        sheet_name: str | None = None,
        overwrite: bool = False,
    ) -> dict[str, Any]:
        if action == "create":
            if name is None or refers_to is None:
                raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, "name and refers_to required for create action")
            return self.create_range(
                workbook_id=workbook_id,
                name=name,
                refers_to=refers_to,
                scope=scope,
                sheet_name=sheet_name,
                overwrite=overwrite,
            )
        elif action == "delete":
            if name is None:
                raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, "name required for delete action")
            return self.delete_range(
                workbook_id=workbook_id,
                name=name,
                scope=scope,
                sheet_name=sheet_name,
            )
        else:
            raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, f"Invalid action: {action}")

    @staticmethod
    def _classify_ref(ref: str) -> str:
        if not ref:
            return "constant"
        ref_upper = ref.upper()
        if "=" in ref_upper and any(op in ref_upper for op in [":", "!", "$", "["]):
            return "reference"
        if any(c in ref_upper for c in ["+", "-", "*", "/", "^", "("]):
            return "formula"
        return "constant"

    @staticmethod
    def _resolve_address(parent: Any, nr: Any) -> str | None:
        try:
            rng = nr.RefersToRange
            return str(rng.Address) if rng else None
        except Exception:
            return None

    def inspect(
        self,
        *,
        action: str,
        workbook_id: str,
        range_name: str = "",
        scope: str = "all",
        sheet_name: str | None = None,
        value_mode: str = "raw",
        row_offset: int = 0,
        row_limit: int = 200,
    ) -> dict[str, Any]:
        if action == "list":
            return self.list_ranges(
                workbook_id=workbook_id,
                scope=scope,
                sheet_name=sheet_name,
            )
        elif action == "info":
            if not range_name:
                raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, "range_name required for info action")
            return self.read_values(
                workbook_id=workbook_id,
                range_name=range_name,
                value_mode=value_mode,
                row_offset=row_offset,
                row_limit=row_limit,
            )
        else:
            raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, f"Invalid action: {action}")
