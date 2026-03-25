from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.models.workbook_models import (
    WorkbookCloseFileRequest,
    WorkbookCreateFileRequest,
    WorkbookInspectRequest,
    WorkbookOpenRequest,
    WorkbookSaveFileRequest,
)
from excelforge.tools.registry import ToolRegistry


def register_workbook_tools(mcp: FastMCP, ctx: Any, registry: ToolRegistry) -> None:
    @mcp.tool(name="workbook.open_file")
    def workbook_open_file(
        file_path: str,
        read_only: bool = False,
        client_request_id: str = "",
    ) -> dict:
        req = WorkbookOpenRequest(
            file_path=file_path,
            read_only=read_only,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="workbook.open_file",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.workbook_service.open_file(req.file_path, req.read_only),
            args_summary={"file_path": req.file_path, "read_only": req.read_only},
        )
        return envelope.model_dump(mode="json")

    registry.add("workbook.open_file", "workbook_tools", "workbook")

    @mcp.tool(name="workbook.inspect")
    def workbook_inspect(
        action: str,
        workbook_id: str = "",
        client_request_id: str = "",
    ) -> dict:
        req = WorkbookInspectRequest(
            action=action,
            workbook_id=workbook_id,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="workbook.inspect",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.workbook_service.inspect(
                action=req.action,
                workbook_id=req.workbook_id,
            ),
            args_summary={"action": req.action, "workbook_id": req.workbook_id},
            default_workbook_id=req.workbook_id if req.action == "info" else None,
        )
        return envelope.model_dump(mode="json")

    registry.add("workbook.inspect", "workbook_tools", "workbook")

    @mcp.tool(name="workbook.save_file")
    def workbook_save_file(
        workbook_id: str,
        save_as_path: str = "",
        client_request_id: str = "",
    ) -> dict:
        req = WorkbookSaveFileRequest(
            workbook_id=workbook_id,
            save_as_path=save_as_path,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="workbook.save_file",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.workbook_service.save_file(
                workbook_id=req.workbook_id,
                save_as_path=req.save_as_path,
            ),
            args_summary={"workbook_id": req.workbook_id, "save_as_path": req.save_as_path},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("workbook.save_file", "workbook_tools", "workbook")

    @mcp.tool(name="workbook.close_file")
    def workbook_close_file(
        workbook_id: str,
        force_discard: bool = False,
        client_request_id: str = "",
    ) -> dict:
        req = WorkbookCloseFileRequest(
            workbook_id=workbook_id,
            force_discard=force_discard,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="workbook.close_file",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.workbook_service.close_file(req.workbook_id, req.force_discard),
            args_summary={"workbook_id": req.workbook_id, "force_discard": req.force_discard},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("workbook.close_file", "workbook_tools", "workbook")

    @mcp.tool(name="workbook.create_file")
    def workbook_create_file(
        file_path: str,
        sheet_names: list[str] | None = None,
        overwrite: bool = False,
        client_request_id: str = "",
    ) -> dict:
        req = WorkbookCreateFileRequest(
            file_path=file_path,
            sheet_names=sheet_names or ["Sheet1"],
            overwrite=overwrite,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="workbook.create_file",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.workbook_service.create_file(
                file_path=req.file_path,
                sheet_names=req.sheet_names,
                overwrite=req.overwrite,
            ),
            args_summary={
                "file_path": req.file_path,
                "sheet_names": req.sheet_names,
                "overwrite": req.overwrite,
            },
            default_file_path=req.file_path,
        )
        return envelope.model_dump(mode="json")

    registry.add("workbook.create_file", "workbook_tools", "workbook")

