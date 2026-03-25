from __future__ import annotations

from pathlib import Path

EXTENSION_FORMAT_MAP = {
    ".xlsx": 51,
    ".xlsm": 52,
    ".xlsb": 50,
    ".xls": 56,
}

VBA_SUPPORTED_EXTENSIONS = {".xlsm", ".xlsb", ".xls"}


def get_file_format(extension: str) -> int | None:
    return EXTENSION_FORMAT_MAP.get(extension.lower())


def supports_vba(extension: str) -> bool:
    return extension.lower() in VBA_SUPPORTED_EXTENSIONS


def is_bas_or_cls(extension: str) -> bool:
    return extension.lower() in {".bas", ".cls"}


def validate_extension_for_save(original_ext: str, new_ext: str) -> tuple[bool, bool]:
    original_has_vba = supports_vba(original_ext)
    new_has_vba = supports_vba(new_ext)
    vba_would_be_stripped = original_has_vba and not new_has_vba
    format_would_change = original_ext.lower() != new_ext.lower()
    return vba_would_be_stripped, format_would_change
