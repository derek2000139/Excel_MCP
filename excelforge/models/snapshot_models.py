from __future__ import annotations

from pydantic import Field

from .common import ClientRequestMixin, StrictModel


class SnapshotGetStatsRequest(ClientRequestMixin):
    workbook_id: str | None = None


class SnapshotStatsWorkbookItem(StrictModel):
    workbook_id: str
    file_path: str
    active_count: int
    size_bytes: int


class SnapshotGetStatsLimits(StrictModel):
    max_per_workbook: int
    max_total_size_mb: int
    max_age_hours: int


class SnapshotGetStatsData(StrictModel):
    total_snapshots: int
    active_snapshots: int
    expired_snapshots: int
    total_size_bytes: int
    active_size_bytes: int
    oldest_active_at: str | None
    newest_active_at: str | None
    by_workbook: list[SnapshotStatsWorkbookItem]
    limits: SnapshotGetStatsLimits


class SnapshotRunCleanupRequest(ClientRequestMixin):
    max_age_hours: int | None = Field(default=None, ge=1)
    workbook_id: str | None = None
    dry_run: bool = False


class SnapshotRunCleanupData(StrictModel):
    dry_run: bool
    snapshots_expired: int
    files_deleted: int
    space_freed_bytes: int
    remaining_active: int
