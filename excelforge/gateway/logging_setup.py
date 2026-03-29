# -*- coding: utf-8 -*-
"""
ExcelForge 日志配置模块。

提供同时输出到控制台和文件的日志配置。
日志文件位置：~/.excelforge/logs/excelforge_YYYYMMDD.log
支持日志文件自动清理。
"""

import logging
import os
import glob
from datetime import datetime
from pathlib import Path


def setup_logging(
    log_level: str = "DEBUG",
    console_level: str = "INFO",
    max_log_files: int = 30
) -> str:
    """
    配置日志系统：同时输出到文件和控制台。

    日志文件位置：~/.excelforge/logs/excelforge_YYYYMMDD.log
    自动清理超过 max_log_files 天的旧日志。

    Args:
        log_level: 文件日志级别，默认 DEBUG
        console_level: 控制台日志级别，默认 INFO
        max_log_files: 最多保留多少个日志文件，默认 30 天

    Returns:
        当前日志文件的完整路径
    """
    log_dir = Path.home() / ".excelforge" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    log_filename = datetime.now().strftime("excelforge_%Y%m%d.log")
    log_filepath = log_dir / log_filename

    file_formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)-5s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    console_formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)-5s] %(message)s",
        datefmt="%H:%M:%S"
    )

    file_handler = logging.FileHandler(
        log_filepath, encoding="utf-8", mode="a"
    )
    file_handler.setLevel(getattr(logging, log_level.upper()))
    file_handler.setFormatter(file_formatter)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(getattr(logging, console_level.upper()))
    console_handler.setFormatter(console_formatter)

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)

    if root_logger.handlers:
        root_logger.handlers.clear()

    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)

    _cleanup_old_logs(str(log_dir), max_log_files)

    logging.info(f"=== ExcelForge MCP Started ===")
    logging.info(f"Log file: {log_filepath}")

    return str(log_filepath)


def _cleanup_old_logs(log_dir: str, max_files: int):
    """
    按文件修改时间删除最旧的日志文件。

    Args:
        log_dir: 日志目录路径
        max_files: 最多保留的文件数量
    """
    pattern = os.path.join(log_dir, "excelforge_*.log")
    files = sorted(glob.glob(pattern), key=os.path.getmtime)

    while len(files) > max_files:
        oldest = files.pop(0)
        try:
            os.remove(oldest)
            logging.debug(f"Removed old log: {oldest}")
        except OSError:
            pass


def get_log_dir() -> str:
    """
    获取日志目录路径。

    Returns:
        日志目录的完整路径
    """
    return str(Path.home() / ".excelforge" / "logs")


def get_current_log_file() -> str:
    """
    获取当前日志文件路径。

    Returns:
        当前日期日志文件的完整路径
    """
    log_dir = Path.home() / ".excelforge" / "logs"
    log_filename = datetime.now().strftime("excelforge_%Y%m%d.log")
    return str(log_dir / log_filename)
