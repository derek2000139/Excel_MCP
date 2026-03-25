from __future__ import annotations

from pydantic import Field

from .common import ClientRequestMixin, StrictModel
from .range_models import ScalarValue


class RollbackListSnapshotsRequest(ClientRequestMixin):
    workbook_id: str | None = None
    limit: int = Field(default=20, ge=1, le=100)
    offset: int = Field(default=0, ge=0)


class SnapshotListItem(StrictModel):
    snapshot_id: str
    workbook_id: str
    sheet_name: str
    range: str
    source_tool: str
    created_at: str
    cell_count: int
    restorable: bool


class RollbackListSnapshotsData(StrictModel):
    total: int
    has_more: bool
    next_offset: int | None
    items: list[SnapshotListItem]


class RollbackPreviewSnapshotRequest(ClientRequestMixin):
    snapshot_id: str
    sample_limit: int = Field(default=20, ge=1, le=50)


class SnapshotDiffItem(StrictModel):
    cell: str
    current_value: ScalarValue
    snapshot_value: ScalarValue
    current_formula: str | None
    snapshot_formula: str | None


class RollbackPreviewSnapshotData(StrictModel):
    snapshot_id: str
    workbook_id: str
    sheet_name: str
    range: str
    changed_cells_count: int
    sample_diffs: list[SnapshotDiffItem]
    preview_token: str
    preview_token_expires_at: str


class RollbackRestoreSnapshotRequest(ClientRequestMixin):
    snapshot_id: str
    preview_token: str


class RollbackRestoreSnapshotData(StrictModel):
    snapshot_id: str
    workbook_id: str
    sheet_name: str
    restored_range: str
    cells_restored: int
    replacement_snapshot_id: str
