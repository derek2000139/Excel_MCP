from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any

from excelforge.config import ToolsConfig
from excelforge.services.audit_service import AuditService
from excelforge.services.backup_service import BackupService
from excelforge.services.format_service import FormatService
from excelforge.services.formula_service import FormulaService
from excelforge.services.operation_service import OperationService
from excelforge.services.range_service import RangeService
from excelforge.services.rollback_service import RollbackService
from excelforge.services.server_service import ServerService
from excelforge.services.sheet_service import SheetService
from excelforge.services.snapshot_service import SnapshotService
from excelforge.services.vba_service import VbaService
from excelforge.services.workbook_service import WorkbookService
from excelforge.services.named_range_service import NamedRangeService
from excelforge.tools.audit_tools import register_audit_tools
from excelforge.tools.format_tools import register_format_tools
from excelforge.tools.formula_tools import register_formula_tools
from excelforge.tools.range_tools import register_range_tools
from excelforge.tools.server_tools import register_server_tools
from excelforge.tools.sheet_tools import register_sheet_tools
from excelforge.tools.snapshot_tools import register_snapshot_tools
from excelforge.tools.vba_tools import register_vba_tools
from excelforge.tools.workbook_tools import register_workbook_tools
from excelforge.tools.named_range_tools import register_named_range_tools
from excelforge.tools.registry import ToolRegistry

logger = logging.getLogger(__name__)


@dataclass
class ToolContext:
    server_service: ServerService
    workbook_service: WorkbookService
    sheet_service: SheetService
    range_service: RangeService
    formula_service: FormulaService
    format_service: FormatService
    rollback_service: RollbackService
    snapshot_service: SnapshotService
    vba_service: VbaService
    backup_service: BackupService
    audit_service: AuditService
    operation_service: OperationService
    named_range_service: NamedRangeService


def register_all_tools(mcp, ctx: ToolContext, tools_config: ToolsConfig | None = None) -> ToolRegistry:
    registry = ToolRegistry()
    groups = tools_config.groups if tools_config else None

    def should_register(group: str) -> bool:
        if groups is None:
            return True
        return getattr(groups, group, True)

    if should_register("core"):
        register_server_tools(mcp, ctx, registry)
        register_workbook_tools(mcp, ctx, registry)
        register_sheet_tools(mcp, ctx, registry)
        register_range_tools(mcp, ctx, registry)
        register_formula_tools(mcp, ctx, registry)
        register_format_tools(mcp, ctx, registry)

    if should_register("recovery"):
        register_snapshot_tools(mcp, ctx, registry)
        register_audit_tools(mcp, ctx, registry)

    if should_register("names"):
        register_named_range_tools(mcp, ctx, registry)

    if should_register("vba"):
        register_vba_tools(mcp, ctx, registry)

    logger.info("tool_registry.registered count=%d names=%s", registry.count(), registry.get_names())
    return registry