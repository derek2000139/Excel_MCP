from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.models.common import ClientRequestMixin
from excelforge.tools.registry import ToolRegistry


def register_server_tools(mcp: FastMCP, ctx: Any, registry: ToolRegistry) -> None:
    @mcp.tool(name="server.get_status")
    def server_get_status(client_request_id: str = "") -> dict:
        req = ClientRequestMixin(client_request_id=client_request_id)
        envelope = ctx.operation_service.run(
            tool_name="server.get_status",
            client_request_id=req.client_request_id,
            operation_fn=ctx.server_service.get_status,
            args_summary={},
        )
        return envelope.model_dump(mode="json")

    registry.add("server.get_status", "server_tools", "server")

