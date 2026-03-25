from __future__ import annotations

import re
import secrets


WORKBOOK_ID_RE = re.compile(r"^wb_g(?P<generation>\d+)_[0-9a-f]{32}$")


def generate_id(prefix: str) -> str:
    return f"{prefix}_{secrets.token_hex(16)}"


def generate_workbook_id(generation: int) -> str:
    return f"wb_g{generation}_{secrets.token_hex(16)}"


def parse_workbook_generation(workbook_id: str) -> int | None:
    match = WORKBOOK_ID_RE.match(workbook_id)
    if not match:
        return None
    return int(match.group("generation"))
