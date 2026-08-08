"""文件级并发锁管理器。

对应方案 7.6 文件并发锁。
"""

from __future__ import annotations

import logging
from pathlib import Path

import filelock

logger = logging.getLogger("timeverse_office_doc_mcp.filelock")


class FileLockManager:
    """文件级并发锁管理器 - 防止多客户端同时写同一文件。"""

    def __init__(self, lock_dir: str | None = None) -> None:
        self.lock_dir = lock_dir or str(
            Path(__file__).resolve().parent.parent.parent.parent / ".locks"
        )
        self._locks: dict[str, filelock.FileLock] = {}

    def acquire(self, file_path: str, timeout: int = 30) -> filelock.FileLock:
        """获取文件锁。"""
        Path(self.lock_dir).mkdir(parents=True, exist_ok=True)
        lock_path = str(Path(self.lock_dir) / (Path(file_path).name + ".lock"))
        lock = filelock.FileLock(lock_path, timeout=timeout)
        lock.acquire()
        self._locks[file_path] = lock
        return lock

    def release(self, file_path: str) -> None:
        """释放文件锁。"""
        lock = self._locks.pop(file_path, None)
        if lock is not None:
            lock.release()


# 全局单例
file_lock_mgr = FileLockManager()
