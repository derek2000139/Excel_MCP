from __future__ import annotations

from pydantic import Field

from .common import ClientRequestMixin, StrictModel
from .range_models import ScalarValue


class NamedRangeListRangesRequest(ClientRequestMixin):
    workbook_id: str
    scope: str = Field(default="all")
    sheet_name: str | None = Field(default=None, max_length=31)


class NamedRangeInspectRequest(ClientRequestMixin):
    action: str = Field(pattern="^(list|info)$")
    workbook_id: str
    range_name: str = ""
    scope: str = "all"
    sheet_name: str | None = Field(default=None, max_length=31)
    value_mode: str = "raw"
    row_offset: int = Field(default=0, ge=0)
    row_limit: int = Field(default=200, ge=1, le=1000)


class NamedRangeItem(StrictModel):
    name: str
    scope: str
    sheet_name: str | None
    refers_to: str
    refers_to_type: str
    visible: bool
    address: str | None


class NamedRangeListRangesData(StrictModel):
    total: int
    items: list[NamedRangeItem]


class NamedRangeReadValuesRequest(ClientRequestMixin):
    workbook_id: str
    range_name: str
    value_mode: str = Field(default="raw")
    row_offset: int = Field(default=0, ge=0)
    row_limit: int = Field(default=200, ge=1, le=1000)


class NamedRangeReadValuesData(StrictModel):
    range_name: str
    scope: str
    sheet_name: str | None
    refers_to: str
    resolved_address: str
    total_rows: int
    returned_rows: int
    column_count: int
    values: list[list[ScalarValue]]
    truncated: bool
    next_row_offset: int | None


class NamedRangeCreateRangeRequest(ClientRequestMixin):
    workbook_id: str
    name: str = Field(pattern=r"^[A-Za-z_][A-Za-z0-9_.]{0,254}$")
    refers_to: str
    scope: str = Field(default="workbook")
    sheet_name: str | None = Field(default=None, max_length=31)
    overwrite: bool = False


class NamedRangeCreateRangeData(StrictModel):
    name: str
    scope: str
    sheet_name: str | None
    refers_to: str
    action: str
    backup_id: str


class NamedRangeDeleteRangeRequest(ClientRequestMixin):
    workbook_id: str
    name: str
    scope: str = Field(default="workbook")
    sheet_name: str | None = Field(default=None, max_length=31)


class NamedRangeDeleteRangeData(StrictModel):
    name: str
    scope: str
    deleted: bool
    backup_id: str


class NamesManageRequest(ClientRequestMixin):
    workbook_id: str
    action: str = Field(pattern="^(create|delete)$")
    name: str | None = None
    refers_to: str | None = None
    scope: str = Field(default="workbook")
    sheet_name: str | None = Field(default=None, max_length=31)
    overwrite: bool = False


class NamesManageCreateData(StrictModel):
    action: str = "create"
    name: str
    scope: str
    sheet_name: str | None
    refers_to: str
    sync_action: str
    backup_id: str


class NamesManageDeleteData(StrictModel):
    action: str = "delete"
    name: str
    scope: str
    deleted: bool
    backup_id: str
