from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Callable

from mcp.server.fastmcp import FastMCP

logger = logging.getLogger(__name__)


class ToolRegistry:
    def __init__(self) -> None:
        self._tools: list[str] = []

    def add(self, name: str, module: str, category: str) -> None:
        self._tools.append(name)
        logger.debug("tool registered tool=%s module=%s category=%s", name, module, category)

    def get_names(self) -> list[str]:
        return list(self._tools)

    def count(self) -> int:
        return len(self._tools)


def register_tool(mcp: FastMCP, name: str, func: Callable) -> None:
    mcp.tool(name=name)(func)
