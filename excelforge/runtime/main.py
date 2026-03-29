from __future__ import annotations

import argparse
import logging
import signal
import threading
from datetime import datetime
from pathlib import Path

from excelforge.runtime.bootstrap import create_runtime_services
from excelforge.runtime.handler import RuntimeJsonRpcHandler
from excelforge.runtime.lifecycle import remove_runtime_lock, write_runtime_lock
from excelforge.runtime.pipe_server import JsonRpcPipeServer
from excelforge.runtime_api import RuntimeApiContext, RuntimeApiDispatcher


def setup_runtime_logging():
    """配置 Runtime 日志，写入到 Gateway 的日志文件。"""
    log_dir = Path.home() / ".excelforge" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    log_filename = log_dir / f"excelforge_{datetime.now().strftime('%Y%m%d')}.log"

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)-5s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    file_handler = logging.FileHandler(log_filename, encoding="utf-8", mode="a")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="excelforge.runtime", description="ExcelForge Runtime Service")
    parser.add_argument("--config", default=None, help="Path to runtime-config.yaml")
    return parser


def main(argv: list[str] | None = None) -> int:
    setup_runtime_logging()

    parser = build_parser()
    args = parser.parse_args(argv)

    services = create_runtime_services(args.config)
    ctx = RuntimeApiContext(services)
    dispatcher = RuntimeApiDispatcher(ctx)
    services.server_service.set_tool_names(dispatcher.method_names())
    handler = RuntimeJsonRpcHandler(dispatcher)

    stop_event = threading.Event()

    def _shutdown(*_: object) -> None:
        stop_event.set()

    signal.signal(signal.SIGINT, _shutdown)
    signal.signal(signal.SIGTERM, _shutdown)

    write_runtime_lock(services.config, args.config)
    server = JsonRpcPipeServer(
        pipe_name=services.config.runtime.pipe_name,
        request_handler=handler.handle_request,
        stop_event=stop_event,
    )

    try:
        server.serve_forever()
    finally:
        remove_runtime_lock(services.config)
        services.shutdown()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
