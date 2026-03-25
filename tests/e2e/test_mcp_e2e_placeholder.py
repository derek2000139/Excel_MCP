from __future__ import annotations

import pytest

pytestmark = [pytest.mark.e2e]


@pytest.mark.skip(reason="Requires MCP client harness for stdio interaction")
def test_mcp_stdio_flow_placeholder() -> None:
    """Placeholder for open -> inspect -> read -> fill -> rollback -> save -> close flow."""
    assert True
