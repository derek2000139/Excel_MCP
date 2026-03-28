from __future__ import annotations

from pathlib import Path

import yaml

from excelforge.config import load_config


def test_load_config_ignores_runtime_identity_metadata_env(monkeypatch, tmp_path: Path) -> None:
    config_path = tmp_path / "runtime-config.yaml"
    data_dir = tmp_path / ".runtime_data_v2"
    config_path.write_text(
        yaml.safe_dump(
            {
                "server": {"version": "2.0.0", "actor_id": "runtime"},
                "runtime": {
                    "version": "2.0.0",
                    "pipe_name": r"\\.\pipe\excelforge-runtime",
                    "data_dir": str(data_dir),
                },
                "excel": {
                    "visible": False,
                    "disable_events": True,
                    "disable_alerts": True,
                    "force_disable_macros": True,
                    "health_ping_enabled": True,
                    "max_rebuild_attempts": 3,
                    "ensure_visibility": True,
                },
                "paths": {
                    "allowed_roots": [str(tmp_path)],
                    "snapshots_dir": str(data_dir / "snapshots"),
                    "backups_dir": str(data_dir / "backups"),
                    "sqlite_path": str(data_dir / "excelforge.db"),
                },
                "limits": {},
                "snapshot": {},
                "backup": {},
                "retention": {},
            },
            sort_keys=False,
            allow_unicode=True,
        ),
        encoding="utf-8",
    )

    monkeypatch.setenv("EXCELFORGE_RUNTIME_SCOPE", "qa")
    monkeypatch.setenv("EXCELFORGE_RUNTIME_INSTANCE", "shared")
    monkeypatch.setenv("EXCELFORGE_RUNTIME_DATA_DIR", str(tmp_path / "runtime_scope_data"))
    monkeypatch.setenv("EXCELFORGE_RUNTIME__PIPE_NAME", r"\\.\pipe\excelforge-runtime.qa.shared")

    config = load_config(config_path)

    assert config.runtime.pipe_name == r"\\.\pipe\excelforge-runtime.qa.shared"
    assert config.runtime.data_dir == str(data_dir)
