from __future__ import annotations

import argparse
import sys
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.gateway.config import GatewayConfig, load_gateway_config
from excelforge.gateway.profile_resolver import BundleRegistry, ProfileResolutionError, ProfileResolver
from excelforge.gateway.runtime_client_manager import get_global_runtime_client
from excelforge.gateway.runtime_identity import (
    RuntimeIdentity,
    resolve_runtime_identity,
)
from excelforge.gateway.utils import call_runtime


TOOL_MANIFEST_MAP: dict[str, str] = {
    "server.get_status": "server.status",
    "server.health": "server.health",
    "workbook.open_file": "workbook.open",
    "workbook.create_file": "workbook.create",
    "workbook.save_file": "workbook.save",
    "workbook.close_file": "workbook.close",
    "workbook.inspect": "workbook.info",
    "names.inspect": "names.list",
    "names.manage": "names.read",
    "sheet.create_sheet": "sheet.create",
    "sheet.rename_sheet": "sheet.rename",
    "sheet.delete_sheet": "sheet.delete",
    "sheet.set_auto_filter": "sheet.auto_filter",
    "sheet.get_conditional_formats": "sheet.get_conditional_formats",
    "sheet.get_data_validations": "sheet.get_data_validations",
    "range.read_values": "range.read",
    "range.write_values": "range.write",
    "range.clear_contents": "range.clear",
    "range.copy": "range.copy",
    "range.insert_rows": "range.insert_rows",
    "range.delete_rows": "range.delete_rows",
    "range.insert_columns": "range.insert_columns",
    "range.delete_columns": "range.delete_columns",
    "range.sort_data": "range.sort",
    "range.merge": "range.merge",
    "format.set_number_format": "format.set_style",
    "format.set_font": "format.set_style",
    "format.set_fill": "format.set_style",
    "format.set_border": "format.set_style",
    "format.set_alignment": "format.set_style",
    "format.set_column_width": "format.auto_fit",
    "format.set_row_height": "format.auto_fit",
    "vba.inspect_project": "vba.inspect_project",
    "vba.scan_code": "vba.scan_code",
    "vba.sync_module": "vba.sync_module",
    "vba.remove_module": "vba.remove_module",
    "vba.execute": "vba.execute_macro",
    "vba.compile": "vba.compile",
    "rollback.manage": "recovery.undo_last",
    "backups.manage": "recovery.list_backups",
    "snapshot.manage": "recovery.list_snapshots",
    "pq.list_connections": "pq.list_connections",
    "pq.list_queries": "pq.list_queries",
    "pq.get_code": "pq.get_query_code",
    "pq.update_query": "pq.update_query",
    "pq.refresh": "pq.refresh",
    "audit.list_operations": "audit.list_operations",
}


@dataclass(frozen=True)
class HostRuntimeSettings:
    identity: RuntimeIdentity
    auto_start: bool
    connect_timeout: int
    call_timeout: int
    runtime_config_path: str | None
    display_name: str


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-mcp",
        description="ExcelForge Unified MCP Host",
    )
    parser.add_argument(
        "--config",
        help="Path to excel-mcp.yaml (optional, uses runtime-config.yaml by default)",
    )
    parser.add_argument(
        "--profile",
        default="basic_edit",
        help="Profile name (default: basic_edit)",
    )
    parser.add_argument(
        "--enable-bundle",
        action="append",
        default=[],
        dest="enabled_bundles",
        help="Extra bundles to enable (can be repeated)",
    )
    parser.add_argument(
        "--disable-bundle",
        action="append",
        default=[],
        dest="disabled_bundles",
        help="Bundles to disable (can be repeated)",
    )
    parser.add_argument(
        "--strict-profile",
        action="store_true",
        help="Fail immediately if profile not found",
    )
    parser.add_argument(
        "--list-profiles",
        action="store_true",
        help="List available profiles and exit",
    )
    parser.add_argument(
        "--list-bundles",
        action="store_true",
        help="List available bundles and exit",
    )
    parser.add_argument(
        "--runtime-scope",
        default="default",
        help="Runtime scope (default: default)",
    )
    parser.add_argument(
        "--runtime-instance",
        default="default",
        help="Runtime instance name (default: default)",
    )
    parser.add_argument(
        "--print-runtime-endpoint",
        action="store_true",
        help="Print resolved Runtime endpoint on startup",
    )
    return parser


def list_profiles_and_exit(profiles_path: Path | None = None) -> None:
    resolver = ProfileResolver(profiles_path)
    profiles = resolver.list_profiles()
    print("Available profiles:")
    for name in profiles:
        info = resolver.get_profile_info(name)
        print(f"  {name}")
        if info["description"]:
            print(f"    {info['description']}")
        print(f"    bundles: {', '.join(info['bundles'])}")


def list_bundles_and_exit(bundles_path: Path | None = None) -> None:
    registry = BundleRegistry(bundles_path)
    bundles = registry.list_bundles()
    print("Available bundles:")
    for name in bundles:
        info = registry.get_bundle_info(name)
        print(f"  {name}")
        if info["description"]:
            print(f"    {info['description']}")
        print(f"    domains: {', '.join(info['domains'])}")


def check_tool_budget(tool_count: int, profile_info: dict[str, Any]) -> None:
    budget = profile_info.get("tool_budget")
    if budget is None:
        return
    if tool_count > budget:
        warnings.warn(
            f"Tool count ({tool_count}) exceeds budget ({budget}) for profile '{profile_info['name']}'. "
            f"Consider reducing enabled bundles.",
            UserWarning,
        )


def _resolve_path(base_dir: Path, raw_path: str | None) -> str | None:
    if not raw_path:
        return None
    path = Path(raw_path)
    if not path.is_absolute():
        path = (base_dir / path).resolve()
    else:
        path = path.resolve()
    return str(path)


def _resolve_gateway_config_path(raw_path: str | None) -> Path | None:
    if raw_path:
        return Path(raw_path).resolve()

    default_path = Path("excel-mcp.yaml")
    if default_path.exists():
        return default_path.resolve()
    return None


def resolve_host_runtime_settings(args: argparse.Namespace) -> HostRuntimeSettings:
    config_path = _resolve_gateway_config_path(args.config)
    gateway_config: GatewayConfig | None = None
    runtime_data_dir: str | None = None
    runtime_config_path = str(Path("runtime-config.yaml").resolve())
    auto_start = True
    connect_timeout = 10
    call_timeout = 30
    display_name = "ExcelForge"

    if config_path is not None:
        gateway_config = load_gateway_config(config_path)
        base_dir = config_path.parent
        runtime_data_dir = _resolve_path(base_dir, gateway_config.gateway.runtime_data_dir)
        runtime_config_path = _resolve_path(base_dir, gateway_config.gateway.runtime_config_path)
        auto_start = gateway_config.gateway.auto_start_runtime
        connect_timeout = gateway_config.gateway.connect_timeout_seconds
        call_timeout = gateway_config.gateway.call_timeout_seconds
        display_name = gateway_config.gateway.display_name

    identity = resolve_runtime_identity(
        runtime_data_dir=runtime_data_dir,
        scope=args.runtime_scope,
        instance_name=args.runtime_instance,
    )
    return HostRuntimeSettings(
        identity=identity,
        auto_start=auto_start,
        connect_timeout=connect_timeout,
        call_timeout=call_timeout,
        runtime_config_path=runtime_config_path,
        display_name=display_name,
    )


def create_host_runtime_client(settings: HostRuntimeSettings) -> Any:
    client = get_global_runtime_client(
        identity=settings.identity,
        auto_start=settings.auto_start,
        connect_timeout=settings.connect_timeout,
        call_timeout=settings.call_timeout,
        runtime_config_path=settings.runtime_config_path,
    )
    return client


def register_tools_for_profile(
    mcp: FastMCP,
    runtime: Any,
    profile_name: str,
    extra_bundles: list[str],
    disabled_bundles: list[str],
    profiles_path: Path | None = None,
    bundles_path: Path | None = None,
) -> None:
    resolver = ProfileResolver(profiles_path)
    bundle_registry = BundleRegistry(bundles_path)

    profile_info = resolver.resolve(profile_name)
    all_bundles = list(profile_info["bundles"])
    for b in extra_bundles:
        if b not in all_bundles:
            all_bundles.append(b)
    for b in disabled_bundles:
        if b in all_bundles:
            all_bundles.remove(b)

    resolved_bundles = bundle_registry.resolve_bundles(all_bundles)
    enabled_tools = bundle_registry.get_all_tools(resolved_bundles)

    check_tool_budget(len(enabled_tools), profile_info)

    def make_tool_handler(runtime_client, tool, method):
        def handler(**kwargs):
            return call_runtime(runtime_client, tool_name=tool, method=method, params=kwargs)
        return handler

    for tool_name in enabled_tools:
        runtime_method = TOOL_MANIFEST_MAP.get(tool_name, tool_name)
        handler = make_tool_handler(runtime, tool_name, runtime_method)
        mcp.tool(name=tool_name)(handler)


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    profiles_path = Path(__file__).parent / "profiles.yaml"
    bundles_path = Path(__file__).parent / "bundles.yaml"

    if args.list_profiles:
        list_profiles_and_exit(profiles_path)
        return 0

    if args.list_bundles:
        list_bundles_and_exit(bundles_path)
        return 0

    if args.strict_profile:
        resolver = ProfileResolver(profiles_path)
        try:
            resolver.resolve(args.profile)
        except ProfileResolutionError as exc:
            print(f"Error: {exc}", file=sys.stderr)
            return 1

    try:
        settings = resolve_host_runtime_settings(args)
        runtime = create_host_runtime_client(settings)
    except Exception as exc:
        print(f"Error creating Runtime client: {exc}", file=sys.stderr)
        return 1

    if args.print_runtime_endpoint:
        print(f"Runtime endpoint: {settings.identity.pipe_name}")
        print(f"Runtime instance ID: {settings.identity.instance_id}")

    display_name = f"{settings.display_name} ({args.profile})"
    mcp = FastMCP(display_name)

    register_tools_for_profile(
        mcp=mcp,
        runtime=runtime,
        profile_name=args.profile,
        extra_bundles=args.enabled_bundles,
        disabled_bundles=args.disabled_bundles,
        profiles_path=profiles_path,
        bundles_path=bundles_path,
    )

    try:
        mcp.run(transport="stdio")
    finally:
        runtime.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
