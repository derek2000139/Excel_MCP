from __future__ import annotations

from pathlib import Path

import pytest

from excelforge.models.error_models import ExcelForgeError
from excelforge.utils.path_guard import ensure_same_extension, normalize_allowed_path


def test_normalize_allowed_path_accepts_under_root(tmp_path: Path) -> None:
    root = tmp_path / "allowed"
    root.mkdir()
    target = root / "book.xlsx"
    target.write_text("x")

    resolved = normalize_allowed_path(str(target), [root])
    assert resolved == target.resolve()


def test_normalize_allowed_path_rejects_outside_root(tmp_path: Path) -> None:
    root = tmp_path / "allowed"
    root.mkdir()
    other = tmp_path / "other.xlsx"
    other.write_text("x")

    with pytest.raises(ExcelForgeError):
        normalize_allowed_path(str(other), [root])


def test_ensure_same_extension() -> None:
    ensure_same_extension(Path("a.xlsx"), Path("b.xlsx"))
    with pytest.raises(ExcelForgeError):
        ensure_same_extension(Path("a.xlsx"), Path("b.xlsm"))
