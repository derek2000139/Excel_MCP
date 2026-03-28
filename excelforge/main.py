from __future__ import annotations

import argparse
import json
from typing import Any

from excelforge.config import write_default_config
from excelforge.gateway.host import main as host_gateway_main
from excelforge.runtime.main import main as runtime_main
from excelforge.server import healthcheck


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="excelforge", description="ExcelForge launcher")
    parser.add_argument("--config", default=None, help="Path to runtime config.yaml")

    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("runtime", help="Run Runtime JSON-RPC pipe server")
    host = sub.add_parser("gateway-host", help="Run unified MCP host")
    host.add_argument("--gateway-config", default="excel-mcp.yaml", help="Path to excel-mcp.yaml")
    host.add_argument("--profile", default="all", help="Host profile name")
    sub.add_parser("healthcheck", help="Validate runtime prerequisites")
    sub.add_parser("write-default-config", help="Write a default config.yaml")

    return parser


def _print_json(data: dict[str, Any]) -> None:
    print(json.dumps(data, ensure_ascii=False, indent=2))


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "write-default-config":
        path = write_default_config("config.yaml")
        _print_json({"ok": True, "config_path": str(path)})
        return 0

    if args.command == "healthcheck":
        data = healthcheck(config_path=args.config)
        _print_json(data)
        return 0

    if args.command == "runtime":
        return runtime_main(["--config", args.config] if args.config else [])
    if args.command == "gateway-host":
        return host_gateway_main(["--config", args.gateway_config, "--profile", args.profile])

    parser.error(f"Unknown command: {args.command}")
    return 2
