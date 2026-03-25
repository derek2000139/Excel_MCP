from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.models.audit_models import AuditListOperationsRequest
from excelforge.tools.registry import ToolRegistry


def register_audit_tools(mcp: FastMCP, ctx: Any, registry: ToolRegistry) -> None:
    @mcp.tool(name="audit.list_operations")
    def audit_list_operations(
        workbook_id: str = "",
        tool_name: str = "",
        success_only: bool = False,
        limit: int = 20,
        offset: int = 0,
        operation_id: str = "",
        client_request_id: str = "",
    ) -> dict:
        req = AuditListOperationsRequest(
            workbook_id=workbook_id,
            tool_name=tool_name,
            success_only=success_only,
            limit=limit,
            offset=offset,
            operation_id=operation_id,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="audit.list_operations",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.audit_service.list_operations(
                workbook_id=req.workbook_id,
                tool_name=req.tool_name,
                success_only=req.success_only,
                limit=req.limit,
                offset=req.offset,
                operation_id=req.operation_id,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "tool_name": req.tool_name,
                "success_only": req.success_only,
                "limit": req.limit,
                "offset": req.offset,
                "operation_id": req.operation_id,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("audit.list_operations", "audit_tools", "audit")

