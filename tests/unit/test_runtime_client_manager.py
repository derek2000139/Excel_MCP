from __future__ import annotations

from pathlib import Path

from excelforge.gateway import runtime_client_manager
from excelforge.gateway.runtime_client_manager import RuntimeClientManager
from excelforge.gateway.runtime_identity import resolve_runtime_identity


def test_start_runtime_process_passes_resolved_identity_to_child(monkeypatch, tmp_path: Path) -> None:
    runtime_config_path = tmp_path / "runtime-config.yaml"
    runtime_config_path.write_text("runtime: {}\n", encoding="utf-8")

    identity = resolve_runtime_identity(
        runtime_data_dir=tmp_path / "runtime_data",
        scope="qa",
        instance_name="shared",
    )
    manager = RuntimeClientManager(
        identity=identity,
        auto_start=True,
        connect_timeout=10,
        call_timeout=30,
        runtime_config_path=str(runtime_config_path),
    )

    captured: dict[str, object] = {}

    def fake_popen(cmd, **kwargs):
        captured["cmd"] = cmd
        captured.update(kwargs)

        class _Proc:
            pass

        return _Proc()

    monkeypatch.setattr(runtime_client_manager.subprocess, "Popen", fake_popen)

    manager._start_runtime_process()

    assert captured["cmd"] == [
        runtime_client_manager.sys.executable,
        "-m",
        "excelforge.runtime",
        "--config",
        str(runtime_config_path.resolve()),
    ]
    assert captured["cwd"] == str(tmp_path.resolve())

    env = captured["env"]
    assert isinstance(env, dict)
    assert env["EXCELFORGE_RUNTIME_SCOPE"] == "qa"
    assert env["EXCELFORGE_RUNTIME_INSTANCE"] == "shared"
    assert env["EXCELFORGE_RUNTIME__PIPE_NAME"] == identity.pipe_name
    assert env["EXCELFORGE_RUNTIME__DATA_DIR"] == str(identity.data_dir)
    assert env["EXCELFORGE_PATHS__SNAPSHOTS_DIR"] == str(identity.data_dir / "snapshots")
    assert env["EXCELFORGE_PATHS__BACKUPS_DIR"] == str(identity.data_dir / "backups")
    assert env["EXCELFORGE_PATHS__SQLITE_PATH"] == str(identity.data_dir / "excelforge.db")
