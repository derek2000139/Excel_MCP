from __future__ import annotations

from datetime import datetime, timezone


def utc_now() -> datetime:
    return datetime.now(tz=timezone.utc)


def utc_now_rfc3339() -> str:
    return utc_now().isoformat().replace("+00:00", "Z")


def parse_rfc3339(value: str) -> datetime:
    normalized = value.replace("Z", "+00:00")
    return datetime.fromisoformat(normalized)

