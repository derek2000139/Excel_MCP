from __future__ import annotations

from argparse import Namespace
from pathlib import Path

import yaml

from excelforge.gateway.host import resolve_host_runtime_settings


def test_resolve_host_runtime_settings_uses_gateway_config_paths(tmp_path: Path) -> None:
    gateway_config_path = tmp_path / "excel-mcp.yaml"
    runtime_config_path = tmp_path / "runtime-config.yaml"
    runtime_config_path.write_text("runtime: {}\n", encoding="utf-8")
    gateway_config_path.write_text(
        yaml.safe_dump(
            {
                "gateway": {
                    "id": "excel-mcp",
                    "display_name": "ExcelForge Unified",
                    "runtime_data_dir": "./custom_runtime_data",
                    "auto_start_runtime": True,
                    "runtime_config_path": "./runtime-config.yaml",
                    "connect_timeout_seconds": 17,
                    "call_timeout_seconds": 45,
                }
            },
            sort_keys=False,
            allow_unicode=True,
        ),
        encoding="utf-8",
    )

    args = Namespace(
        config=str(gateway_config_path),
        runtime_scope="team-alpha",
        runtime_instance="shared",
    )

    settings = resolve_host_runtime_settings(args)

    assert settings.display_name == "ExcelForge Unified"
    assert settings.connect_timeout == 17
    assert settings.call_timeout == 45
    assert settings.runtime_config_path == str(runtime_config_path.resolve())
    assert settings.identity.data_dir == (tmp_path / "custom_runtime_data").resolve()
    assert settings.identity.scope == "team-alpha"
    assert settings.identity.instance_name == "shared"
    assert settings.identity.pipe_name.endswith("excelforge-runtime.team-alpha.shared")
