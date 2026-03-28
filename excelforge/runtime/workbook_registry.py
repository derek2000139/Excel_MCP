from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from excelforge.runtime.handle_ownership import ensure_workbook_id_owned, is_foreign_workbook_id
from excelforge.utils.ids import parse_workbook_generation


@dataclass
class WorkbookHandle:
    workbook_id: str
    workbook_name: str
    file_path: str
    read_only: bool
    opened_at: str
    workbook_obj: Any
    file_format: str = "xlsx"
    max_rows: int = 1048576
    max_columns: int = 16384


class WorkbookRegistry:
    def __init__(self, runtime_fingerprint: str | None = None) -> None:
        self._items: dict[str, WorkbookHandle] = {}
        self._generation = 1
        self._runtime_fingerprint = runtime_fingerprint

    @property
    def generation(self) -> int:
        return self._generation

    @property
    def runtime_fingerprint(self) -> str | None:
        return self._runtime_fingerprint

    def count(self) -> int:
        return len(self._items)

    def add(self, handle: WorkbookHandle) -> None:
        self._items[handle.workbook_id] = handle

    def get(self, workbook_id: str) -> WorkbookHandle | None:
        ensure_workbook_id_owned(workbook_id, self._runtime_fingerprint)
        if not self._is_current_generation(workbook_id):
            return None
        return self._items.get(workbook_id)

    def require(self, workbook_id: str) -> WorkbookHandle:
        handle = self.get(workbook_id)
        if handle is None:
            raise KeyError(workbook_id)
        return handle

    def remove(self, workbook_id: str) -> WorkbookHandle | None:
        ensure_workbook_id_owned(workbook_id, self._runtime_fingerprint)
        if not self._is_current_generation(workbook_id):
            return None
        return self._items.pop(workbook_id, None)

    def list_items(self) -> list[WorkbookHandle]:
        return list(self._items.values())

    def find_by_path(self, file_path: str) -> WorkbookHandle | None:
        normalized = str(Path(file_path).resolve()).lower()
        for item in self._items.values():
            if str(Path(item.file_path).resolve()).lower() == normalized:
                return item
        return None

    def invalidate_all(self) -> None:
        self._items.clear()

    def bump_generation(self) -> int:
        self._items.clear()
        self._generation += 1
        return self._generation

    def _is_current_generation(self, workbook_id: str) -> bool:
        generation = parse_workbook_generation(workbook_id)
        if generation is None:
            return True
        return generation == self._generation

    def is_foreign_workbook_id(self, workbook_id: str) -> bool:
        return is_foreign_workbook_id(workbook_id, self._runtime_fingerprint)

    def is_stale_workbook_id(self, workbook_id: str) -> bool:
        if self.is_foreign_workbook_id(workbook_id):
            return False
        generation = parse_workbook_generation(workbook_id)
        return generation is not None and generation != self._generation
