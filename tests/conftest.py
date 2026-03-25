from __future__ import annotations

import importlib.util

import pytest


def _has_pywin32() -> bool:
    return importlib.util.find_spec("pythoncom") is not None and importlib.util.find_spec("win32com") is not None


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    has_pywin32 = _has_pywin32()
    skip_integration = pytest.mark.skip(reason="requires pywin32 + local Excel Desktop")

    for item in items:
        if "integration" in item.keywords and not has_pywin32:
            item.add_marker(skip_integration)
