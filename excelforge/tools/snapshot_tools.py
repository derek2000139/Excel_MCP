from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.models.snapshot_models import SnapshotRunCleanupRequest
from excelforge.tools.registry import ToolRegistry


def register_snapshot_tools(mcp: FastMCP, ctx: Any, registry: ToolRegistry) -> None:
    @mcp.tool(name="snapshot.list")
    def snapshot_list(
        workbook_id: str = "",
        client_request_id: str = "",
    ) -> dict:
        snapshot_stats = ctx.snapshot_service.get_stats(workbook_id=workbook_id)
        snapshot_list_result = ctx.rollback_service.list_snapshots(
            workbook_id=workbook_id,
            limit=100,
            offset=0,
        )
        return {
            "snapshot_stats": snapshot_stats,
            "snapshots": snapshot_list_result,
        }

    registry.add("snapshot.list", "snapshot_tools", "snapshot")

    @mcp.tool(name="snapshot.preview")
    def snapshot_preview(
        snapshot_id: str,
        sample_limit: int = 20,
        client_request_id: str = "",
    ) -> dict:
        return ctx.operation_service.run(
            tool_name="snapshot.preview",
            client_request_id=client_request_id,
            operation_fn=lambda: ctx.rollback_service.preview_snapshot(
                snapshot_id=snapshot_id,
                sample_limit=sample_limit,
            ),
            args_summary={"snapshot_id": snapshot_id, "sample_limit": sample_limit},
        ).model_dump(mode="json")

    registry.add("snapshot.preview", "snapshot_tools", "snapshot")

    @mcp.tool(name="snapshot.delete")
    def snapshot_delete(
        action: str,
        snapshot_id: str = "",
        max_age_hours: int | None = None,
        workbook_id: str = "",
        dry_run: bool = False,
        client_request_id: str = "",
    ) -> dict:
        if action == "cleanup":
            req = SnapshotRunCleanupRequest(
                max_age_hours=max_age_hours,
                workbook_id=workbook_id,
                dry_run=dry_run,
                client_request_id=client_request_id,
            )
            return ctx.operation_service.run(
                tool_name="snapshot.delete",
                client_request_id=req.client_request_id,
                operation_fn=lambda: ctx.snapshot_service.run_cleanup(
                    max_age_hours=req.max_age_hours,
                    workbook_id=req.workbook_id,
                    dry_run=req.dry_run,
                ),
                args_summary={
                    "max_age_hours": req.max_age_hours,
                    "workbook_id": req.workbook_id,
                    "dry_run": req.dry_run,
                },
                default_workbook_id=req.workbook_id,
            ).model_dump(mode="json")
        else:
            return {
                "success": False,
                "code": "E400_BAD_REQUEST",
                "message": f"Invalid action: {action}. Must be 'cleanup'",
            }

    registry.add("snapshot.delete", "snapshot_tools", "snapshot")

    @mcp.tool(name="rollback.manage")
    def rollback_manage(
        action: str,
        workbook_id: str = "",
        max_age_hours: int | None = None,
        dry_run: bool = False,
        client_request_id: str = "",
    ) -> dict:
        if action == "stats":
            return ctx.operation_service.run(
                tool_name="rollback.manage",
                client_request_id=client_request_id,
                operation_fn=lambda: ctx.snapshot_service.get_stats(),
                args_summary={"action": action},
            ).model_dump(mode="json")
        elif action == "cleanup":
            req = SnapshotRunCleanupRequest(
                max_age_hours=max_age_hours,
                workbook_id=workbook_id,
                dry_run=dry_run,
                client_request_id=client_request_id,
            )
            return ctx.operation_service.run(
                tool_name="rollback.manage",
                client_request_id=req.client_request_id,
                operation_fn=lambda: ctx.snapshot_service.run_cleanup(
                    max_age_hours=req.max_age_hours,
                    workbook_id=req.workbook_id,
                    dry_run=req.dry_run,
                ),
                args_summary={
                    "max_age_hours": req.max_age_hours,
                    "workbook_id": req.workbook_id,
                    "dry_run": req.dry_run,
                    "action": action,
                },
                default_workbook_id=req.workbook_id,
            ).model_dump(mode="json")
        else:
            return {
                "success": False,
                "code": "E400_BAD_REQUEST",
                "message": f"Invalid action: {action}. Must be 'stats' or 'cleanup'",
            }

    registry.add("rollback.manage", "snapshot_tools", "snapshot")