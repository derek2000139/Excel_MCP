from __future__ import annotations

import time
from typing import Any


def wait_for_calculation(app: Any, timeout_seconds: int = 10) -> bool:
    try:
        xlDone = 0
        start = time.time()
        deadline = start + timeout_seconds

        app.Calculate()

        while app.CalculationState != xlDone:
            if time.time() > deadline:
                return False
            time.sleep(0.1)

        return True
    except Exception:
        return False


def check_dynamic_array_support(app: Any) -> bool:
    try:
        test_rng = app.Range("XFD1")
        test_rng.Formula2 = "=1"
        test_rng.ClearContents()
        return True
    except Exception:
        return False
