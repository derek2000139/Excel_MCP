from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.models.range_models import (
    RangeAutofitRequest,
    RangeClearContentsRequest,
    RangeCopyRangeRequest,
    RangeFindReplaceRequest,
    RangeManageMergeRequest,
    RangeMergeCellsRequest,
    RangeReadValuesRequest,
    RangeSortDataRequest,
    RangeUnmergeCellsRequest,
    RangeWriteValuesRequest,
)
from excelforge.tools.registry import ToolRegistry


def register_range_tools(mcp: FastMCP, ctx: Any, registry: ToolRegistry) -> None:
    @mcp.tool(name="range.read_values")
    def range_read_values(
        workbook_id: str,
        sheet_name: str,
        range: str,
        value_mode: str = "raw",
        include_formulas: bool = False,
        row_offset: int = 0,
        row_limit: int = 200,
        client_request_id: str = "",
    ) -> dict:
        req = RangeReadValuesRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range=range,
            value_mode=value_mode,
            include_formulas=include_formulas,
            row_offset=row_offset,
            row_limit=row_limit,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="range.read_values",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.range_service.read_values(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
                value_mode=req.value_mode,
                include_formulas=req.include_formulas,
                row_offset=req.row_offset,
                row_limit=req.row_limit,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range": req.range,
                "value_mode": req.value_mode,
                "include_formulas": req.include_formulas,
                "row_offset": req.row_offset,
                "row_limit": req.row_limit,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("range.read_values", "range_tools", "range")

    @mcp.tool(name="range.write_values")
    def range_write_values(
        workbook_id: str,
        sheet_name: str,
        start_cell: str,
        values: list[list[str | int | float | bool | None]],
        client_request_id: str = "",
    ) -> dict:
        req = RangeWriteValuesRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            start_cell=start_cell,
            values=values,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="range.write_values",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.range_service.write_values(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                start_cell=req.start_cell,
                values=req.values,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "start_cell": req.start_cell,
                "rows": len(req.values),
                "cols": len(req.values[0]) if req.values else 0,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("range.write_values", "range_tools", "range")

    @mcp.tool(name="range.clear_contents")
    def range_clear_contents(
        workbook_id: str,
        sheet_name: str,
        range: str,
        scope: str = "contents",
        client_request_id: str = "",
    ) -> dict:
        req = RangeClearContentsRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range=range,
            scope=scope,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="range.clear_contents",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.range_service.clear_contents(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
                scope=req.scope,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range": req.range,
                "scope": req.scope,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("range.clear_contents", "range_tools", "range")

    @mcp.tool(name="range.copy_range")
    def range_copy_range(
        workbook_id: str,
        source_sheet: str,
        source_range: str,
        target_sheet: str,
        target_start_cell: str,
        paste_mode: str = "values",
        target_workbook_id: str = "",
        client_request_id: str = "",
    ) -> dict:
        req = RangeCopyRangeRequest(
            workbook_id=workbook_id,
            source_sheet=source_sheet,
            source_range=source_range,
            target_sheet=target_sheet,
            target_start_cell=target_start_cell,
            paste_mode=paste_mode,
            target_workbook_id=target_workbook_id,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="range.copy_range",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.range_service.copy_range(
                workbook_id=req.workbook_id,
                source_sheet=req.source_sheet,
                source_range=req.source_range,
                target_sheet=req.target_sheet,
                target_start_cell=req.target_start_cell,
                paste_mode=req.paste_mode,
                target_workbook_id=req.target_workbook_id,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "source_sheet": req.source_sheet,
                "source_range": req.source_range,
                "target_sheet": req.target_sheet,
                "target_start_cell": req.target_start_cell,
                "paste_mode": req.paste_mode,
                "target_workbook_id": req.target_workbook_id,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("range.copy_range", "range_tools", "range")

    @mcp.tool(name="range.manage_rows")
    def range_manage_rows(
        workbook_id: str,
        sheet_name: str,
        action: str,
        row: int,
        count: int = 1,
        client_request_id: str = "",
    ) -> dict:
        if action == "insert":
            return ctx.operation_service.run(
                tool_name="range.manage_rows",
                client_request_id=client_request_id,
                operation_fn=lambda: ctx.range_service.insert_rows(
                    workbook_id=workbook_id,
                    sheet_name=sheet_name,
                    row_number=row,
                    count=count,
                ),
                args_summary={
                    "workbook_id": workbook_id,
                    "sheet_name": sheet_name,
                    "action": action,
                    "row": row,
                    "count": count,
                },
                default_workbook_id=workbook_id,
            ).model_dump(mode="json")
        elif action == "delete":
            return ctx.operation_service.run(
                tool_name="range.manage_rows",
                client_request_id=client_request_id,
                operation_fn=lambda: ctx.range_service.delete_rows(
                    workbook_id=workbook_id,
                    sheet_name=sheet_name,
                    start_row=row,
                    count=count,
                ),
                args_summary={
                    "workbook_id": workbook_id,
                    "sheet_name": sheet_name,
                    "action": action,
                    "row": row,
                    "count": count,
                },
                default_workbook_id=workbook_id,
            ).model_dump(mode="json")
        else:
            return {
                "success": False,
                "code": "E400_BAD_REQUEST",
                "message": f"Invalid action: {action}. Must be 'insert' or 'delete'",
            }

    registry.add("range.manage_rows", "range_tools", "range")

    @mcp.tool(name="range.manage_columns")
    def range_manage_columns(
        workbook_id: str,
        sheet_name: str,
        action: str,
        column: str,
        count: int = 1,
        client_request_id: str = "",
    ) -> dict:
        if action == "insert":
            return ctx.operation_service.run(
                tool_name="range.manage_columns",
                client_request_id=client_request_id,
                operation_fn=lambda: ctx.range_service.insert_columns(
                    workbook_id=workbook_id,
                    sheet_name=sheet_name,
                    column=column,
                    count=count,
                ),
                args_summary={
                    "workbook_id": workbook_id,
                    "sheet_name": sheet_name,
                    "action": action,
                    "column": column,
                    "count": count,
                },
                default_workbook_id=workbook_id,
            ).model_dump(mode="json")
        elif action == "delete":
            return ctx.operation_service.run(
                tool_name="range.manage_columns",
                client_request_id=client_request_id,
                operation_fn=lambda: ctx.range_service.delete_columns(
                    workbook_id=workbook_id,
                    sheet_name=sheet_name,
                    start_column=column,
                    count=count,
                ),
                args_summary={
                    "workbook_id": workbook_id,
                    "sheet_name": sheet_name,
                    "action": action,
                    "column": column,
                    "count": count,
                },
                default_workbook_id=workbook_id,
            ).model_dump(mode="json")
        else:
            return {
                "success": False,
                "code": "E400_BAD_REQUEST",
                "message": f"Invalid action: {action}. Must be 'insert' or 'delete'",
            }

    registry.add("range.manage_columns", "range_tools", "range")

    @mcp.tool(name="range.sort_data")
    def range_sort_data(
        workbook_id: str,
        sheet_name: str,
        range: str,
        sort_fields: list[dict],
        has_header: bool = False,
        case_sensitive: bool = False,
        client_request_id: str = "",
    ) -> dict:
        req = RangeSortDataRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range=range,
            sort_fields=sort_fields,
            has_header=has_header,
            case_sensitive=case_sensitive,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="range.sort_data",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.range_service.sort_data(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
                sort_fields=req.sort_fields,
                has_header=req.has_header,
                case_sensitive=req.case_sensitive,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range": req.range,
                "sort_fields_count": len(req.sort_fields),
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("range.sort_data", "range_tools", "range")

    @mcp.tool(name="range.merge_cells")
    def range_merge_cells(
        workbook_id: str,
        sheet_name: str,
        range: str,
        across: bool = False,
        client_request_id: str = "",
    ) -> dict:
        req = RangeMergeCellsRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range=range,
            across=across,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="range.merge_cells",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.range_service.merge_cells(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
                across=req.across,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range": req.range,
                "across": req.across,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("range.merge_cells", "range_tools", "range")

    @mcp.tool(name="range.unmerge_cells")
    def range_unmerge_cells(
        workbook_id: str,
        sheet_name: str,
        range: str,
        client_request_id: str = "",
    ) -> dict:
        req = RangeUnmergeCellsRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range=range,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="range.unmerge_cells",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.range_service.unmerge_cells(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range": req.range,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("range.unmerge_cells", "range_tools", "range")

    @mcp.tool(name="range.manage_merge")
    def range_manage_merge(
        workbook_id: str,
        sheet_name: str,
        range: str,
        action: str,
        across: bool = False,
        client_request_id: str = "",
    ) -> dict:
        req = RangeManageMergeRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range=range,
            action=action,
            across=across,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="range.manage_merge",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.range_service.manage_merge(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
                action=req.action,
                across=req.across,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range": req.range,
                "action": req.action,
                "across": req.across,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("range.manage_merge", "range_tools", "range")

    @mcp.tool(name="range.find_replace")
    def range_find_replace(
        workbook_id: str,
        find_what: str,
        replace_with: str | None = None,
        sheet_name: str | None = None,
        range_address: str | None = None,
        match_case: bool = False,
        match_entire_cell: bool = False,
        client_request_id: str = "",
    ) -> dict:
        req = RangeFindReplaceRequest(
            workbook_id=workbook_id,
            find_what=find_what,
            replace_with=replace_with,
            sheet_name=sheet_name,
            range_address=range_address,
            match_case=match_case,
            match_entire_cell=match_entire_cell,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="range.find_replace",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.range_service.find_replace(
                workbook_id=req.workbook_id,
                find_what=req.find_what,
                replace_with=req.replace_with,
                sheet_name=req.sheet_name,
                range_address=req.range_address,
                match_case=req.match_case,
                match_entire_cell=req.match_entire_cell,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "find_what": req.find_what,
                "replace_with": req.replace_with,
                "sheet_name": req.sheet_name,
                "range_address": req.range_address,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("range.find_replace", "range_tools", "range")

    @mcp.tool(name="range.autofit")
    def range_autofit(
        workbook_id: str,
        sheet_name: str | None = None,
        range_address: str | None = None,
        autofit_type: str = "columns",
        client_request_id: str = "",
    ) -> dict:
        req = RangeAutofitRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range_address=range_address,
            autofit_type=autofit_type,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="range.autofit",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.range_service.autofit(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range_address,
                autofit_type=req.autofit_type,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range_address": req.range_address,
                "autofit_type": req.autofit_type,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("range.autofit", "range_tools", "range")