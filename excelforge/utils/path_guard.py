from __future__ import annotations

from pathlib import Path
from typing import Set

from excelforge.models.error_models import ErrorCode, ExcelForgeError

SUPPORTED_EXTENSIONS: Set[str] = {".xlsx", ".xlsm", ".xlsb", ".xls"}


def _is_unc(path: str) -> bool:
    return path.startswith("\\\\")


def ensure_supported_extension(path: Path, allowed_extensions: Set[str] | None = None) -> None:
    ext = path.suffix.lower()
    valid_exts = allowed_extensions if allowed_extensions is not None else SUPPORTED_EXTENSIONS
    if ext not in valid_exts:
        raise ExcelForgeError(
            ErrorCode.E415_EXTENSION_UNSUPPORTED,
            f"Unsupported extension: {path.suffix}",
        )


def normalize_allowed_path(
    raw_path: str,
    allowed_roots: list[Path],
    allowed_extensions: set[str] | None = None,
) -> Path:
    if _is_unc(raw_path):
        raise ExcelForgeError(ErrorCode.E403_PATH_NOT_ALLOWED, "UNC paths are not allowed")

    p = Path(raw_path)
    if not p.is_absolute():
        raise ExcelForgeError(ErrorCode.E403_PATH_NOT_ALLOWED, "Path must be absolute")

    resolved = p.resolve()
    ensure_supported_extension(resolved, allowed_extensions)

    for root in allowed_roots:
        root_str = str(root)
        if root_str == "*":
            return resolved
        root_resolved = root.resolve()
        try:
            resolved.relative_to(root_resolved)
            return resolved
        except ValueError:
            continue

    raise ExcelForgeError(
        ErrorCode.E403_PATH_NOT_ALLOWED,
        f"Path is outside allowed roots: {resolved}",
    )


def ensure_same_extension(src_path: Path, dest_path: Path) -> None:
    if src_path.suffix.lower() != dest_path.suffix.lower():
        raise ExcelForgeError(
            ErrorCode.E423_FEATURE_NOT_SUPPORTED,
            "save_as_path extension must match source workbook extension",
        )
