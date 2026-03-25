from __future__ import annotations

from pydantic import Field

from .common import ClientRequestMixin, StrictModel


class BackupListBackupsRequest(ClientRequestMixin):
    workbook_id: str | None = None
    file_path: str | None = None
    limit: int = Field(default=20, ge=1, le=100)
    offset: int = Field(default=0, ge=0)


class BackupListItem(StrictModel):
    backup_id: str
    workbook_id: str
    file_path: str
    backup_file_path: str
    file_size_bytes: int
    source_tool: str
    description: str
    created_at: str
    expired: bool
    expired_reason: str | None = None
    cleaned_at: str | None = None


class BackupListBackupsData(StrictModel):
    total: int
    has_more: bool
    next_offset: int | None
    items: list[BackupListItem]
    manual_restore_instructions: str


class BackupRestoreFileRequest(ClientRequestMixin):
    workbook_id: str
    backup_id: str


class BackupRestoreFileData(StrictModel):
    backup_id: str
    pre_restore_backup_id: str | None
    original_workbook_id: str
    new_workbook_id: str
    file_path: str
    invalidated_snapshots: int
    restored_at: str
