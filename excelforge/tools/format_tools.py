from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.models.format_models import FormatAutoFitColumnsRequest, FormatManageRequest, FormatSetRangeStyleRequest, StyleModel
from excelforge.tools.registry import ToolRegistry


def register_format_tools(mcp: FastMCP, ctx: Any, registry: ToolRegistry) -> None:
    @mcp.tool(name="format.manage")
    def format_manage(
        action: str,
        workbook_id: str,
        sheet_name: str,
        range: str = "",
        style: dict[str, Any] | None = None,
        columns: str | None = None,
        client_request_id: str = "",
    ) -> dict:
        style_model = StyleModel(**style) if style else None
        req = FormatManageRequest(
            action=action,
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range=range,
            style=style_model,
            columns=columns,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="format.manage",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.format_service.manage(
                action=req.action,
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range=req.range,
                style=req.style,
                columns=req.columns,
            ),
            args_summary={
                "action": req.action,
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range": req.range,
                "has_style": req.style is not None,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("format.manage", "format_tools", "format")
