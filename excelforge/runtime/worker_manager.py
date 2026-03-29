# -*- coding: utf-8 -*-
"""
Excel Worker 进程生命周期管理模块。

职责：
- 追踪当前 Worker 的 PID
- rebuild 时先终止旧进程再创建新进程
- 防止并发 rebuild
- 启动时和定期扫描僵尸进程
"""

import threading
import time
import logging
import os
from typing import Optional, Callable, Any

logger = logging.getLogger(__name__)

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False
    logger.warning("psutil not available, using basic PID management")


class ExcelWorkerManager:
    """
    管理 Excel Worker 进程的完整生命周期。
    """

    def __init__(self):
        self._worker_pid: Optional[int] = None
        self._rebuild_lock = threading.Lock()
        self._rebuild_in_progress = False
        self._kill_grace_seconds = 3
        self._creation_time: Optional[float] = None

    def register_worker_pid(self, pid: int):
        """Worker 创建后注册 PID。"""
        self._worker_pid = pid
        self._creation_time = time.time()
        logger.info(f"[WorkerManager] Registered worker PID={pid}")

    def get_worker_pid(self) -> Optional[int]:
        """获取当前 Worker 的 PID。"""
        return self._worker_pid

    def clear_registration(self):
        """清除 Worker PID 注册。"""
        old_pid = self._worker_pid
        self._worker_pid = None
        self._creation_time = None
        if old_pid:
            logger.info(f"[WorkerManager] Cleared registration PID={old_pid}")

    def kill_current_worker(self) -> bool:
        """
        终止当前 Worker 进程。

        策略：terminate → 等待 → force kill
        """
        if not self._worker_pid:
            return True

        pid = self._worker_pid

        if HAS_PSUTIL:
            return self._kill_with_psutil(pid)
        else:
            return self._kill_with_taskkill(pid)

    def _kill_with_psutil(self, pid: int) -> bool:
        """使用 psutil 终止进程。"""
        try:
            proc = psutil.Process(pid)
        except psutil.NoSuchProcess:
            logger.debug(f"[WorkerManager] PID={pid} already gone")
            self.clear_registration()
            return True

        logger.info(f"[WorkerManager] Terminating PID={pid}...")
        try:
            proc.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            self.clear_registration()
            return True

        try:
            proc.wait(timeout=self._kill_grace_seconds)
            logger.info(f"[WorkerManager] PID={pid} exited gracefully")
            self.clear_registration()
            return True
        except psutil.TimeoutExpired:
            pass

        logger.warning(f"[WorkerManager] PID={pid} force killing")
        try:
            proc.kill()
            proc.wait(timeout=2)
        except (psutil.NoSuchProcess, psutil.TimeoutExpired):
            pass

        self.clear_registration()
        return True

    def _kill_with_taskkill(self, pid: int) -> bool:
        """无 psutil 时的 fallback：用 taskkill 命令。"""
        logger.info(f"[WorkerManager] taskkill PID={pid}")
        try:
            os.system(f"taskkill /PID {pid} /F >nul 2>&1")
            self.clear_registration()
            return True
        except Exception as e:
            logger.error(f"[WorkerManager] taskkill failed: {e}")
            return False

    def is_worker_alive(self) -> bool:
        """检查 Worker 进程是否存活。"""
        if not self._worker_pid:
            return False

        if HAS_PSUTIL:
            try:
                proc = psutil.Process(self._worker_pid)
                return proc.is_running() and proc.status() != psutil.STATUS_ZOMBIE
            except psutil.NoSuchProcess:
                self.clear_registration()
                return False
        else:
            ret = os.system(
                f'tasklist /FI "PID eq {self._worker_pid}" '
                f'/FI "IMAGENAME eq EXCEL.EXE" /NH >nul 2>&1'
            )
            return ret == 0

    def rebuild_worker(
        self,
        create_fn: Callable[[], Any],
        pre_rebuild_hook: Optional[Callable] = None
    ) -> Any:
        """
        安全重建 Worker。

        流程：
        1. 加锁
        2. 调用 pre_rebuild_hook（清理 Registry 等）
        3. 终止旧进程
        4. 等待 OS 回收
        5. 调用 create_fn 创建新进程
        6. 释放锁
        """
        with self._rebuild_lock:
            if self._rebuild_in_progress:
                logger.warning("[WorkerManager] Rebuild already in progress")
                return None
            self._rebuild_in_progress = True

        try:
            old_pid = self._worker_pid
            logger.warning(
                f"[WorkerManager] === REBUILD START === old_pid={old_pid}"
            )

            if pre_rebuild_hook:
                try:
                    pre_rebuild_hook()
                except Exception as e:
                    logger.error(f"[WorkerManager] pre_rebuild_hook failed: {e}")

            if old_pid:
                self.kill_current_worker()

            time.sleep(0.5)

            new_worker = create_fn()

            logger.info(
                f"[WorkerManager] === REBUILD COMPLETE === "
                f"old_pid={old_pid} new_pid={self._worker_pid}"
            )
            return new_worker

        except Exception as e:
            logger.error(f"[WorkerManager] Rebuild failed: {e}")
            raise
        finally:
            self._rebuild_in_progress = False

    def scan_and_cleanup_orphans(self) -> int:
        """
        扫描并清理孤儿 Excel 进程。

        判定标准：
        - 进程名为 EXCEL.EXE
        - 不是当前注册的 Worker PID
        - 没有可见窗口
        - CPU 使用率接近 0
        """
        if not HAS_PSUTIL:
            logger.debug("[WorkerManager] psutil not available, skip scan")
            return 0

        cleaned = 0
        current_pid = self._worker_pid

        for proc in psutil.process_iter(['pid', 'name']):
            try:
                name = (proc.info['name'] or '').upper()
                if 'EXCEL' not in name:
                    continue

                pid = proc.info['pid']
                if pid == current_pid:
                    continue

                cpu = proc.cpu_percent(interval=1)
                if cpu > 1.0:
                    continue

                if self._has_visible_window(pid):
                    continue

                logger.warning(
                    f"[WorkerManager] Orphan detected: PID={pid} CPU={cpu:.1f}%"
                )
                proc.kill()
                cleaned += 1
                logger.info(f"[WorkerManager] Orphan PID={pid} killed")

            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        if cleaned:
            logger.info(
                f"[WorkerManager] Cleaned {cleaned} orphan process(es)"
            )
        return cleaned

    @staticmethod
    def _has_visible_window(pid: int) -> bool:
        """检查进程是否有可见窗口。"""
        try:
            import win32gui
            import win32process

            windows = []

            def callback(hwnd, results):
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                _, w_pid = win32process.GetWindowThreadProcessId(hwnd)
                if w_pid == pid:
                    results.append(hwnd)
                return True

            win32gui.EnumWindows(callback, windows)
            return len(windows) > 0
        except ImportError:
            return False
