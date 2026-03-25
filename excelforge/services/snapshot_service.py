from __future__ import annotations

import gzip
import json
from dataclasses import dataclass
from datetime import timedelta
from pathlib import Path
from typing import Any

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.persistence.snapshot_repo import SnapshotMetaRecord, SnapshotRepository
from excelforge.runtime.workbook_registry import WorkbookHandle
from excelforge.utils.address_parser import CellRef, RangeRef, cell_to_a1, parse_range
from excelforge.utils.ids import generate_id
from excelforge.utils.timestamps import parse_rfc3339, utc_now, utc_now_rfc3339
from excelforge.utils.value_codec import to_scalar


@dataclass
class SnapshotCell:
    state: str
    value: Any
    formula: str | None
    number_format: str


class SnapshotService:
    def __init__(self, config: AppConfig, snapshot_repo: SnapshotRepository) -> None:
        self._config = config
        self._snapshot_repo = snapshot_repo

    def create_snapshot(
        self,
        *,
        workbook: WorkbookHandle,
        worksheet: Any,
        range_address: str,
        source_tool: str,
    ) -> str:
        range_ref = parse_range(range_address)
        if range_ref.cell_count > self._config.limits.max_snapshot_cells:
            raise ExcelForgeError(
                ErrorCode.E413_RANGE_TOO_LARGE,
                f"Snapshot range too large: {range_ref.cell_count}",
            )

        cells = self._read_cells_matrix(worksheet, range_ref)
        snapshot_id = generate_id("snap")
        created_at = utc_now_rfc3339()
        payload = {
            "schema_version": 1,
            "snapshot_id": snapshot_id,
            "workbook_id": workbook.workbook_id,
            "sheet_name": worksheet.Name,
            "range": range_address,
            "created_at": created_at,
            "cells": cells,
        }

        snapshot_path = self._snapshot_file_path(snapshot_id)
        with gzip.open(snapshot_path, "wt", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))
        file_size_bytes = self._snapshot_file_size(str(snapshot_path))

        self._snapshot_repo.insert_meta(
            SnapshotMetaRecord(
                snapshot_id=snapshot_id,
                workbook_id=workbook.workbook_id,
                file_path=workbook.file_path,
                sheet_name=worksheet.Name,
                range_address=range_address,
                source_tool=source_tool,
                created_at=created_at,
                cell_count=range_ref.cell_count,
                file_path_snapshot=str(snapshot_path),
                file_size_bytes=file_size_bytes,
            )
        )
        self._enforce_snapshot_quotas(workbook.workbook_id)
        return snapshot_id

    def _snapshot_file_path(self, snapshot_id: str) -> Path:
        return self._config.snapshots_dir / f"{snapshot_id}.json.gz"

    def _read_cells_matrix(self, worksheet: Any, range_ref: RangeRef) -> list[list[dict[str, Any]]]:
        matrix: list[list[dict[str, Any]]] = []
        for row in range(range_ref.start.row, range_ref.end.row + 1):
            row_items: list[dict[str, Any]] = []
            for col in range(range_ref.start.col, range_ref.end.col + 1):
                cell = worksheet.Cells(row, col)
                value = to_scalar(cell.Value2)
                has_formula = bool(cell.HasFormula)
                formula = str(cell.Formula) if has_formula else None
                number_format = str(cell.NumberFormat)
                state = "blank"
                if has_formula:
                    state = "formula"
                elif value not in (None, ""):
                    state = "value"
                row_items.append(
                    {
                        "state": state,
                        "value": value,
                        "formula": formula,
                        "number_format": number_format,
                    }
                )
            matrix.append(row_items)
        return matrix

    def load_snapshot(self, snapshot_id: str) -> tuple[dict[str, Any], dict[str, Any]]:
        meta = self._snapshot_repo.get_meta(snapshot_id)
        if meta is None:
            raise ExcelForgeError(ErrorCode.E404_SNAPSHOT_NOT_FOUND, f"Snapshot not found: {snapshot_id}")
        if bool(meta["expired"]):
            raise ExcelForgeError(ErrorCode.E409_SNAPSHOT_EXPIRED, f"Snapshot expired: {snapshot_id}")

        path = Path(str(meta["file_path_snapshot"]))
        if not path.exists():
            raise ExcelForgeError(
                ErrorCode.E404_SNAPSHOT_NOT_FOUND,
                f"Snapshot file missing: {snapshot_id}",
            )
        with gzip.open(path, "rt", encoding="utf-8") as f:
            payload = json.load(f)
        return meta, payload

    def restore_snapshot(
        self,
        *,
        workbook: WorkbookHandle,
        worksheet: Any,
        snapshot_payload: dict[str, Any],
    ) -> int:
        range_ref = parse_range(str(snapshot_payload["range"]))
        cells = snapshot_payload["cells"]
        restored = 0
        for row_idx, row in enumerate(range_ref_to_cells(range_ref)):
            for col_idx, cell_ref in enumerate(row):
                data = cells[row_idx][col_idx]
                cell = worksheet.Cells(cell_ref.row, cell_ref.col)
                state = data.get("state")
                if state == "formula":
                    cell.Formula = data.get("formula")
                elif state == "value":
                    cell.Value2 = data.get("value")
                else:
                    cell.ClearContents()
                cell.NumberFormat = data.get("number_format", "General")
                restored += 1
        return restored

    def preview_diffs(
        self,
        *,
        worksheet: Any,
        snapshot_payload: dict[str, Any],
        sample_limit: int,
    ) -> tuple[int, list[dict[str, Any]]]:
        range_ref = parse_range(str(snapshot_payload["range"]))
        cells = snapshot_payload["cells"]
        changed_count = 0
        sample: list[dict[str, Any]] = []

        for row_idx, row in enumerate(range_ref_to_cells(range_ref)):
            for col_idx, cell_ref in enumerate(row):
                snap = cells[row_idx][col_idx]
                current = worksheet.Cells(cell_ref.row, cell_ref.col)
                current_has_formula = bool(current.HasFormula)
                current_formula = str(current.Formula) if current_has_formula else None
                current_value = to_scalar(current.Value2)
                current_state = "formula" if current_has_formula else ("value" if current_value not in (None, "") else "blank")

                changed = (
                    current_state != snap.get("state")
                    or current_formula != snap.get("formula")
                    or current_value != snap.get("value")
                )

                if changed:
                    changed_count += 1
                    if len(sample) < sample_limit:
                        sample.append(
                            {
                                "cell": cell_to_a1(cell_ref),
                                "current_value": current_value,
                                "snapshot_value": snap.get("value"),
                                "current_formula": current_formula,
                                "snapshot_formula": snap.get("formula"),
                            }
                        )

        return changed_count, sample

    def expire_workbook_snapshots(self, workbook_id: str) -> int:
        entries = self._snapshot_repo.expire_by_workbook_with_rows(workbook_id, reason="workbook_closed")
        deleted, _, cleaned_ids = self._delete_snapshot_files(entries)
        if cleaned_ids:
            self._snapshot_repo.mark_cleaned(cleaned_ids, utc_now_rfc3339())
        return deleted

    def expire_sheet_snapshots(self, workbook_id: str, sheet_name: str) -> int:
        entries = self._snapshot_repo.expire_by_sheet_with_rows(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            reason="sheet_structure_changed",
        )
        deleted, _, cleaned_ids = self._delete_snapshot_files(entries)
        if cleaned_ids:
            self._snapshot_repo.mark_cleaned(cleaned_ids, utc_now_rfc3339())
        return deleted

    def expire_all_active_snapshots(self, reason: str = "stale_handle") -> int:
        active = self._snapshot_repo.list_active_snapshot_files(workbook_id=None)
        snapshot_ids = [str(row["snapshot_id"]) for row in active]
        entries = self._snapshot_repo.expire_snapshots(snapshot_ids, reason=reason)
        deleted, _, cleaned_ids = self._delete_snapshot_files(entries)
        if cleaned_ids:
            self._snapshot_repo.mark_cleaned(cleaned_ids, utc_now_rfc3339())
        return deleted

    def expire_all_for_workbook(self, workbook_id: str) -> int:
        entries = self._snapshot_repo.expire_by_workbook_with_rows(workbook_id, reason="restored_from_backup")
        deleted, _, cleaned_ids = self._delete_snapshot_files(entries)
        if cleaned_ids:
            self._snapshot_repo.mark_cleaned(cleaned_ids, utc_now_rfc3339())
        return deleted

    def create_preview_token(self, snapshot_id: str) -> tuple[str, str]:
        token = generate_id("rtok")
        created_at_dt = utc_now()
        ttl_minutes = int(self._config.snapshot.preview_token_ttl_minutes)
        expires_at_dt = created_at_dt + timedelta(minutes=ttl_minutes)
        created_at = created_at_dt.isoformat().replace("+00:00", "Z")
        expires_at = expires_at_dt.isoformat().replace("+00:00", "Z")
        self._snapshot_repo.insert_preview_token(
            token=token,
            snapshot_id=snapshot_id,
            created_at=created_at,
            expires_at=expires_at,
        )
        return token, expires_at

    def consume_preview_token(self, token: str, snapshot_id: str) -> None:
        token_row = self._snapshot_repo.get_preview_token(token)
        if token_row is None:
            raise ExcelForgeError(
                ErrorCode.E409_PREVIEW_TOKEN_INVALID,
                "preview_token is invalid",
            )
        if token_row["snapshot_id"] != snapshot_id:
            raise ExcelForgeError(
                ErrorCode.E409_PREVIEW_TOKEN_INVALID,
                "preview_token does not match snapshot_id",
            )
        if bool(token_row["used"]):
            raise ExcelForgeError(
                ErrorCode.E409_PREVIEW_TOKEN_INVALID,
                "preview_token already used",
            )
        expires_at = parse_rfc3339(str(token_row["expires_at"]))
        if expires_at <= utc_now():
            raise ExcelForgeError(
                ErrorCode.E409_PREVIEW_TOKEN_INVALID,
                "preview_token expired",
            )
        self._snapshot_repo.mark_preview_token_used(token)

    def get_stats(self, workbook_id: str | None = None) -> dict[str, Any]:
        stats = self._snapshot_repo.get_stats(workbook_id=workbook_id)
        stats["limits"] = {
            "max_per_workbook": int(self._config.snapshot.max_per_workbook),
            "max_total_size_mb": int(self._config.snapshot.max_total_size_mb),
            "max_age_hours": int(self._config.snapshot.max_age_hours),
        }
        return stats

    def run_cleanup(
        self,
        *,
        max_age_hours: int | None = None,
        workbook_id: str | None = None,
        dry_run: bool = False,
    ) -> dict[str, Any]:
        age_hours = int(max_age_hours or self._config.snapshot.max_age_hours)
        cutoff_dt = utc_now() - timedelta(hours=age_hours)
        cutoff_ts = cutoff_dt.isoformat().replace("+00:00", "Z")

        if dry_run:
            active_rows = self._snapshot_repo.list_active_snapshot_files(workbook_id=workbook_id)
            would_expire = [r for r in active_rows if parse_rfc3339(str(r["created_at"])) < cutoff_dt]
            expired_rows = self._snapshot_repo.list_expired_uncleaned(workbook_id=workbook_id)
            all_candidates = would_expire + expired_rows
            space_freed = sum(int(row.get("file_size_bytes", 0) or 0) for row in all_candidates)
            remaining_active = len(active_rows) - len(would_expire)
            return {
                "dry_run": True,
                "snapshots_expired": len(would_expire) + len(expired_rows),
                "files_deleted": 0,
                "space_freed_bytes": int(space_freed),
                "remaining_active": max(remaining_active, 0),
            }

        age_expired = self._snapshot_repo.expire_by_age(
            cutoff_ts=cutoff_ts,
            reason="time_expired",
            workbook_id=workbook_id,
        )
        entries = self._snapshot_repo.list_expired_uncleaned(workbook_id=workbook_id)
        deleted_count, space_freed, cleaned_ids = self._delete_snapshot_files(entries)
        if cleaned_ids:
            self._snapshot_repo.mark_cleaned(cleaned_ids, utc_now_rfc3339())
        remaining_active = len(self._snapshot_repo.list_active_snapshot_files(workbook_id=workbook_id))
        return {
            "dry_run": False,
            "snapshots_expired": int(age_expired),
            "files_deleted": int(deleted_count),
            "space_freed_bytes": int(space_freed),
            "remaining_active": int(remaining_active),
        }

    def rename_sheet_snapshot_refs(self, workbook_id: str, old_name: str, new_name: str) -> int:
        return self._snapshot_repo.rename_sheet_refs(workbook_id, old_name, new_name)

    def count_active_for_sheet(self, workbook_id: str, sheet_name: str) -> int:
        return self._snapshot_repo.count_active_for_sheet(workbook_id, sheet_name)

    def _enforce_snapshot_quotas(self, workbook_id: str) -> None:
        per_workbook_limit = int(self._config.snapshot.max_per_workbook)
        count_limit_ids: list[str] = []
        if per_workbook_limit > 0:
            workbook_rows = self._snapshot_repo.list_active_snapshot_files(workbook_id=workbook_id)
            over_count = len(workbook_rows) - per_workbook_limit
            if over_count > 0:
                count_limit_ids.extend(str(row["snapshot_id"]) for row in workbook_rows[:over_count])

        if count_limit_ids:
            count_entries = self._snapshot_repo.expire_snapshots(count_limit_ids, reason="count_limit")
            _, _, cleaned_ids = self._delete_snapshot_files(count_entries)
            if cleaned_ids:
                self._snapshot_repo.mark_cleaned(cleaned_ids, utc_now_rfc3339())

        total_bytes_limit = int(self._config.snapshot.max_total_size_mb) * 1024 * 1024
        size_limit_ids: list[str] = []
        if total_bytes_limit > 0:
            global_rows = self._snapshot_repo.list_active_snapshot_files(workbook_id=None)
            total_bytes = sum(int(row.get("file_size_bytes", 0) or 0) for row in global_rows)
            seen_ids = set(count_limit_ids)
            for row in global_rows:
                if total_bytes <= total_bytes_limit:
                    break
                snapshot_id = str(row["snapshot_id"])
                if snapshot_id in seen_ids:
                    continue
                size_limit_ids.append(snapshot_id)
                seen_ids.add(snapshot_id)
                total_bytes -= int(row.get("file_size_bytes", 0) or 0)

        if not size_limit_ids:
            return

        entries = self._snapshot_repo.expire_snapshots(size_limit_ids, reason="size_limit")
        deleted, _, cleaned_ids = self._delete_snapshot_files(entries)
        if cleaned_ids:
            self._snapshot_repo.mark_cleaned(cleaned_ids, utc_now_rfc3339())
        _ = deleted

    @staticmethod
    def _snapshot_file_size(file_path: str) -> int:
        path = Path(file_path)
        if not path.exists():
            return 0
        try:
            return int(path.stat().st_size)
        except Exception:
            return 0

    @staticmethod
    def _delete_snapshot_files(entries: list[dict[str, object]]) -> tuple[int, int, list[str]]:
        deleted = 0
        space_freed = 0
        cleaned_ids: list[str] = []
        for entry in entries:
            snapshot_id = str(entry["snapshot_id"])
            raw = str(entry["file_path_snapshot"])
            file_size = int(entry.get("file_size_bytes", 0) or 0)
            path = Path(raw)
            if path.exists():
                path.unlink(missing_ok=True)
                deleted += 1
                space_freed += file_size
            cleaned_ids.append(snapshot_id)
        return deleted, space_freed, cleaned_ids


def range_ref_to_cells(range_ref: RangeRef) -> list[list[CellRef]]:
    rows: list[list[CellRef]] = []
    for row in range(range_ref.start.row, range_ref.end.row + 1):
        current_row: list[CellRef] = []
        for col in range(range_ref.start.col, range_ref.end.col + 1):
            current_row.append(CellRef(row=row, col=col))
        rows.append(current_row)
    return rows
