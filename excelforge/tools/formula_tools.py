from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.models.formula_models import (
    FormulaFillRangeRequest,
    FormulaGetDependenciesRequest,
    FormulaRepairReferencesRequest,
    FormulaSetSingleRequest,
)
from excelforge.tools.registry import ToolRegistry


def register_formula_tools(mcp: FastMCP, ctx: Any, registry: ToolRegistry) -> None:
    @mcp.tool(name="formula.fill_range")
    def formula_fill_range(
        workbook_id: str,
        sheet_name: str,
        range: str,
        formula: str,
        formula_type: str = "standard",
        preview_rows: int = 5,
        client_request_id: str = "",
    ) -> dict:
        req = FormulaFillRangeRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range=range,
            formula=formula,
            formula_type=formula_type,
            preview_rows=preview_rows,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="formula.fill_range",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.formula_service.fill_range(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
                formula=req.formula,
                formula_type=req.formula_type,
                preview_rows=req.preview_rows,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range": req.range,
                "formula_type": req.formula_type,
                "formula_length": len(req.formula),
                "preview_rows": req.preview_rows,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("formula.fill_range", "formula_tools", "formula")

    @mcp.tool(name="formula.set_single")
    def formula_set_single(
        workbook_id: str,
        sheet_name: str,
        cell: str,
        formula: str,
        formula_type: str = "standard",
        client_request_id: str = "",
    ) -> dict:
        req = FormulaSetSingleRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            cell=cell,
            formula=formula,
            formula_type=formula_type,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="formula.set_single",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.formula_service.set_single(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                cell=req.cell,
                formula=req.formula,
                formula_type=req.formula_type,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "cell": req.cell,
                "formula_type": req.formula_type,
                "formula_length": len(req.formula),
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("formula.set_single", "formula_tools", "formula")

    @mcp.tool(name="formula.get_dependencies")
    def formula_get_dependencies(
        workbook_id: str,
        sheet_name: str,
        cell: str,
        client_request_id: str = "",
    ) -> dict:
        req = FormulaGetDependenciesRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            cell=cell,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="formula.get_dependencies",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.formula_service.get_dependencies(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                cell=req.cell,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "cell": req.cell,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("formula.get_dependencies", "formula_tools", "formula")

    @mcp.tool(name="formula.repair_references")
    def formula_repair_references(
        workbook_id: str,
        sheet_name: str,
        range: str,
        action: str,
        replacements: list | None = None,
        client_request_id: str = "",
    ) -> dict:
        req = FormulaRepairReferencesRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range=range,
            action=action,
            replacements=replacements,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="formula.repair_references",
            client_request_id=req.client_request_id,
            operation_fn=lambda: (
                ctx.formula_service.scan_formulas(
                    workbook_id=req.workbook_id,
                    sheet_name=req.sheet_name,
                    range_address=req.range,
                )
                if req.action == "scan"
                else ctx.formula_service.repair_formulas(
                    workbook_id=req.workbook_id,
                    sheet_name=req.sheet_name,
                    range_address=req.range,
                    replacements=req.replacements or [],
                )
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range": req.range,
                "action": req.action,
                "replacements_count": len(req.replacements) if req.replacements else 0,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("formula.repair_references", "formula_tools", "formula")

