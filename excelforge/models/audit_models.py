from __future__ import annotations

from pydantic import Field

from .common import ClientRequestMixin, StrictModel


class AuditListOperationsRequest(ClientRequestMixin):
    workbook_id: str | None = None
    tool_name: str | None = None
    success_only: bool = False
    limit: int = Field(default=20, ge=1, le=200)
    offset: int = Field(default=0, ge=0)
    operation_id: str | None = None


class AuditOperationItem(StrictModel):
    operation_id: str
    tool_name: str
    workbook_id: str | None
    started_at: str
    duration_ms: int
    success: bool
    code: str
    message: str
    affected_sheet: str | None
    affected_range: str | None
    snapshot_id: str | None
    client_request_id: str | None
    recovery_strategy: str | None
    recovery_reference_id: str | None
    recovery_tool_name: str | None


class AuditListOperationsData(StrictModel):
    total: int
    has_more: bool
    next_offset: int | None
    items: list[AuditOperationItem]
