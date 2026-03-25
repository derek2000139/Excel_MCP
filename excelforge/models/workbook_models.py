from __future__ import annotations

from pydantic import Field

from .common import ClientRequestMixin, StrictModel


class WorkbookOpenRequest(ClientRequestMixin):
    file_path: str
    read_only: bool = False


class WorkbookOpenData(StrictModel):
    workbook_id: str
    workbook_name: str
    file_path: str
    read_only: bool
    already_open: bool
    has_macros: bool
    sheet_names: list[str]
    opened_at: str
    file_format: str = "xlsx"
    max_rows: int = 1048576
    max_columns: int = 16384
    vba_enabled: bool = False


class WorkbookListOpenRequest(ClientRequestMixin):
    pass


class WorkbookListItem(StrictModel):
    workbook_id: str
    workbook_name: str
    file_path: str
    read_only: bool
    dirty: bool
    stale: bool
    opened_at: str


class WorkbookListOpenData(StrictModel):
    items: list[WorkbookListItem]


class WorkbookInspectRequest(ClientRequestMixin):
    action: str = Field(pattern="^(list|info)$")
    workbook_id: str = ""


class WorkbookGetInfoRequest(ClientRequestMixin):
    workbook_id: str


class WorkbookSheetInfo(StrictModel):
    index: int = Field(description="工作表索引，从1开始")
    name: str
    visible: bool
    protected: bool
    used_range: str
    used_rows: int
    used_columns: int


class WorkbookGetInfoData(StrictModel):
    workbook_id: str
    workbook_name: str
    file_path: str
    read_only: bool
    dirty: bool
    has_macros: bool
    sheet_count: int
    sheets: list[WorkbookSheetInfo]
    file_format: str = "xlsx"
    max_rows: int = 1048576
    max_columns: int = 16384


class WorkbookSaveFileRequest(ClientRequestMixin):
    workbook_id: str
    save_as_path: str | None = None


class WorkbookSaveFileData(StrictModel):
    workbook_id: str
    saved_path: str
    dirty: bool
    saved_at: str
    save_type: str = "save"
    format_converted: bool = False
    vba_stripped: bool = False
    original_path: str | None = None


class WorkbookCloseFileRequest(ClientRequestMixin):
    workbook_id: str
    force_discard: bool = False


class WorkbookCloseFileData(StrictModel):
    workbook_id: str
    closed: bool
    changes_discarded: bool
    invalidated_snapshot_count: int


class WorkbookCreateFileRequest(ClientRequestMixin):
    file_path: str
    sheet_names: list[str] = Field(default_factory=lambda: ["Sheet1"], min_length=1, max_length=20)
    overwrite: bool = False


class WorkbookCreateFileData(StrictModel):
    workbook_id: str
    workbook_name: str
    file_path: str
    sheet_names: list[str]
    overwritten: bool
    created_at: str
    vba_enabled: bool = False
