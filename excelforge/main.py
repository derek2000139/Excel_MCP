from __future__ import annotations

import argparse
import json
from typing import Any

from excelforge.config import write_default_config
from excelforge.server import create_app, healthcheck


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="excelforge", description="ExcelForge MCP server")
    parser.add_argument("--config", default=None, help="Path to config.yaml")

    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("serve", help="Run MCP server over stdio")
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

    if args.command == "serve":
        app = create_app(config_path=args.config)
        try:
            app.run_stdio()
        finally:
            app.shutdown()
        return 0

    parser.error(f"Unknown command: {args.command}")
    return 2
