from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.models.named_range_models import (
    NamedRangeCreateRangeRequest,
    NamedRangeDeleteRangeRequest,
    NamedRangeInspectRequest,
    NamesManageRequest,
)
from excelforge.tools.registry import ToolRegistry


def register_named_range_tools(mcp: FastMCP, ctx: Any, registry: ToolRegistry) -> None:
    @mcp.tool(name="names.inspect")
    def named_range_inspect(
        action: str,
        workbook_id: str,
        range_name: str = "",
        scope: str = "all",
        sheet_name: str = "",
        value_mode: str = "raw",
        row_offset: int = 0,
        row_limit: int = 200,
        client_request_id: str = "",
    ) -> dict:
        req = NamedRangeInspectRequest(
            action=action,
            workbook_id=workbook_id,
            range_name=range_name,
            scope=scope,
            sheet_name=sheet_name or None,
            value_mode=value_mode,
            row_offset=row_offset,
            row_limit=row_limit,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="names.inspect",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.named_range_service.inspect(
                action=req.action,
                workbook_id=req.workbook_id,
                range_name=req.range_name,
                scope=req.scope,
                sheet_name=req.sheet_name,
                value_mode=req.value_mode,
                row_offset=req.row_offset,
                row_limit=req.row_limit,
            ),
            args_summary={
                "action": req.action,
                "workbook_id": req.workbook_id,
                "range_name": req.range_name,
                "scope": req.scope,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("names.inspect", "named_range_tools", "named_range")

    @mcp.tool(name="names.create")
    def named_range_create_range(
        workbook_id: str,
        name: str,
        refers_to: str,
        scope: str = "workbook",
        sheet_name: str = "",
        overwrite: bool = False,
        client_request_id: str = "",
    ) -> dict:
        req = NamedRangeCreateRangeRequest(
            workbook_id=workbook_id,
            name=name,
            refers_to=refers_to,
            scope=scope,
            sheet_name=sheet_name,
            overwrite=overwrite,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="names.create",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.named_range_service.create_range(
                workbook_id=req.workbook_id,
                name=req.name,
                refers_to=req.refers_to,
                scope=req.scope,
                sheet_name=req.sheet_name,
                overwrite=req.overwrite,
            ),
            args_summary={"workbook_id": req.workbook_id, "name": req.name, "scope": req.scope},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("names.create", "named_range_tools", "named_range")

    @mcp.tool(name="names.delete")
    def named_range_delete_range(
        workbook_id: str,
        name: str,
        scope: str = "workbook",
        sheet_name: str = "",
        client_request_id: str = "",
    ) -> dict:
        req = NamedRangeDeleteRangeRequest(
            workbook_id=workbook_id,
            name=name,
            scope=scope,
            sheet_name=sheet_name,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="names.delete",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.named_range_service.delete_range(
                workbook_id=req.workbook_id,
                name=req.name,
                scope=req.scope,
                sheet_name=req.sheet_name,
            ),
            args_summary={"workbook_id": req.workbook_id, "name": req.name, "scope": req.scope},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("names.delete", "named_range_tools", "named_range")

    @mcp.tool(name="names.manage")
    def named_range_manage(
        workbook_id: str,
        action: str,
        name: str = "",
        refers_to: str = "",
        scope: str = "workbook",
        sheet_name: str = "",
        overwrite: bool = False,
        client_request_id: str = "",
    ) -> dict:
        req = NamesManageRequest(
            workbook_id=workbook_id,
            action=action,
            name=name,
            refers_to=refers_to,
            scope=scope,
            sheet_name=sheet_name,
            overwrite=overwrite,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="names.manage",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.named_range_service.manage(
                workbook_id=req.workbook_id,
                action=req.action,
                name=req.name,
                refers_to=req.refers_to,
                scope=req.scope,
                sheet_name=req.sheet_name,
                overwrite=req.overwrite,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "action": req.action,
                "name": req.name,
                "scope": req.scope,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("names.manage", "named_range_tools", "named_range")
