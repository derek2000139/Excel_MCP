# -*- coding: utf-8 -*-
"""
Excel COM 工具函数模块。

提供从 Excel COM 对象获取进程信息等工具函数。
"""

import logging
from typing import Optional

logger = logging.getLogger(__name__)

try:
    import win32process
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
    logger.warning("win32process not available, get_excel_pid will return 0")


def get_excel_pid(excel_app) -> int:
    """
    从 Excel Application COM 对象获取底层进程 PID。

    原理：Excel.Application.Hwnd → GetWindowThreadProcessId

    Args:
        excel_app: Excel Application COM 对象

    Returns:
        Excel 进程的 PID，如果获取失败返回 0
    """
    if not HAS_WIN32:
        return 0

    try:
        hwnd = int(excel_app.Hwnd)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        logger.debug(f"[COM] Excel PID={pid} (from Hwnd={hwnd})")
        return pid
    except Exception as e:
        logger.warning(f"[COM] Failed to get Excel PID: {e}")
        return 0
