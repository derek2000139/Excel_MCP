from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.models.sheet_models import (
    SheetCreateRequest,
    SheetDeleteSheetRequest,
    SheetGetRulesRequest,
    SheetInspectStructureRequest,
    SheetRenameRequest,
    SheetSetAutoFilterRequest,
)
from excelforge.tools.registry import ToolRegistry


def register_sheet_tools(mcp: FastMCP, ctx: Any, registry: ToolRegistry) -> None:
    @mcp.tool(name="sheet.inspect_structure")
    def sheet_inspect_structure(
        workbook_id: str,
        sheet_name: str,
        sample_rows: int = 5,
        scan_rows: int = 10,
        max_profile_columns: int = 50,
        client_request_id: str = "",
    ) -> dict:
        req = SheetInspectStructureRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            sample_rows=sample_rows,
            scan_rows=scan_rows,
            max_profile_columns=max_profile_columns,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="sheet.inspect_structure",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.sheet_service.inspect_structure(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                sample_rows=req.sample_rows,
                scan_rows=req.scan_rows,
                max_profile_columns=req.max_profile_columns,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "sample_rows": req.sample_rows,
                "scan_rows": req.scan_rows,
                "max_profile_columns": req.max_profile_columns,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("sheet.inspect_structure", "sheet_tools", "sheet")

    @mcp.tool(name="sheet.create_sheet")
    def sheet_create_sheet(
        workbook_id: str,
        sheet_name: str,
        position: str = "last",
        client_request_id: str = "",
    ) -> dict:
        req = SheetCreateRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            position=position,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="sheet.create_sheet",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.sheet_service.create_sheet(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                position=req.position,
            ),
            args_summary={"workbook_id": req.workbook_id, "sheet_name": req.sheet_name, "position": req.position},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("sheet.create_sheet", "sheet_tools", "sheet")

    @mcp.tool(name="sheet.rename_sheet")
    def sheet_rename_sheet(
        workbook_id: str,
        current_name: str,
        new_name: str,
        client_request_id: str = "",
    ) -> dict:
        req = SheetRenameRequest(
            workbook_id=workbook_id,
            current_name=current_name,
            new_name=new_name,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="sheet.rename_sheet",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.sheet_service.rename_sheet(
                workbook_id=req.workbook_id,
                current_name=req.current_name,
                new_name=req.new_name,
            ),
            args_summary={"workbook_id": req.workbook_id, "current_name": req.current_name, "new_name": req.new_name},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("sheet.rename_sheet", "sheet_tools", "sheet")

    @mcp.tool(name="sheet.delete_sheet")
    def sheet_delete_sheet(
        workbook_id: str,
        sheet_name: str,
        preview: bool = False,
        confirm_token: str = "",
        client_request_id: str = "",
    ) -> dict:
        req = SheetDeleteSheetRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            preview=preview,
            confirm_token=confirm_token,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="sheet.delete_sheet",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.sheet_service.delete_sheet(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                preview=req.preview,
                confirm_token=req.confirm_token,
            ),
            args_summary={"workbook_id": req.workbook_id, "sheet_name": req.sheet_name, "preview": req.preview},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("sheet.delete_sheet", "sheet_tools", "sheet")

    @mcp.tool(name="sheet.set_auto_filter")
    def sheet_set_auto_filter(
        workbook_id: str,
        sheet_name: str,
        action: str,
        range: str = "",
        filters: list[dict] | None = None,
        client_request_id: str = "",
    ) -> dict:
        req = SheetSetAutoFilterRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            action=action,
            range=range,
            filters=filters,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="sheet.set_auto_filter",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.sheet_service.set_auto_filter(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                action=req.action,
                range_address=req.range,
                filters=req.filters,
            ),
            args_summary={"workbook_id": req.workbook_id, "sheet_name": req.sheet_name, "action": req.action},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("sheet.set_auto_filter", "sheet_tools", "sheet")

    @mcp.tool(name="sheet.get_rules")
    def sheet_get_rules(
        workbook_id: str,
        sheet_name: str,
        rule_type: str = "conditional_formats",
        range: str = "",
        limit: int = 100,
        client_request_id: str = "",
    ) -> dict:
        req = SheetGetRulesRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            rule_type=rule_type,
            range=range,
            limit=limit,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="sheet.get_rules",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.sheet_service.get_rules(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                rule_type=req.rule_type,
                range_address=req.range,
                limit=req.limit,
            ),
            args_summary={"workbook_id": req.workbook_id, "sheet_name": req.sheet_name, "rule_type": req.rule_type},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("sheet.get_rules", "sheet_tools", "sheet")
