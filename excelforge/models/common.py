from __future__ import annotations

from typing import Any

from pydantic import BaseModel, ConfigDict, Field

from .error_models import ErrorCode


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid", populate_by_name=True)


class ToolMeta(StrictModel):
    tool_name: str
    operation_id: str
    workbook_id: str | None = None
    snapshot_id: str | None = None
    rollback_supported: bool = False
    duration_ms: int
    server_version: str
    client_request_id: str | None = None
    warnings: list[str] = Field(default_factory=list)


class ToolEnvelope(StrictModel):
    success: bool
    code: ErrorCode | str
    message: str
    data: Any | None
    meta: ToolMeta


class PaginationResult(StrictModel):
    total: int
    has_more: bool
    next_offset: int | None


class ClientRequestMixin(StrictModel):
    client_request_id: str | None = Field(default=None, max_length=64)


def ok_envelope(
    *,
    tool_name: str,
    operation_id: str,
    duration_ms: int,
    server_version: str,
    data: Any,
    workbook_id: str | None = None,
    snapshot_id: str | None = None,
    rollback_supported: bool = False,
    client_request_id: str | None = None,
    warnings: list[str] | None = None,
) -> ToolEnvelope:
    return ToolEnvelope(
        success=True,
        code=ErrorCode.OK,
        message="operation completed",
        data=data,
        meta=ToolMeta(
            tool_name=tool_name,
            operation_id=operation_id,
            workbook_id=workbook_id,
            snapshot_id=snapshot_id,
            rollback_supported=rollback_supported,
            duration_ms=duration_ms,
            server_version=server_version,
            client_request_id=client_request_id,
            warnings=warnings or [],
        ),
    )


def error_envelope(
    *,
    tool_name: str,
    operation_id: str,
    duration_ms: int,
    server_version: str,
    code: ErrorCode | str,
    message: str,
    workbook_id: str | None = None,
    snapshot_id: str | None = None,
    client_request_id: str | None = None,
    warnings: list[str] | None = None,
) -> ToolEnvelope:
    return ToolEnvelope(
        success=False,
        code=code,
        message=message,
        data=None,
        meta=ToolMeta(
            tool_name=tool_name,
            operation_id=operation_id,
            workbook_id=workbook_id,
            snapshot_id=snapshot_id,
            rollback_supported=False,
            duration_ms=duration_ms,
            server_version=server_version,
            client_request_id=client_request_id,
            warnings=warnings or [],
        ),
    )
