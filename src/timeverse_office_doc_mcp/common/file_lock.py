"""文件级并发锁管理器。

对应方案 7.6 文件并发锁。
"""

from __future__ import annotations

import functools
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

    def is_held(self, file_path: str) -> bool:
        """当前线程/进程是否已持有该文件的锁。"""
        return file_path in self._locks


# 全局单例
file_lock_mgr = FileLockManager()


def _locked_write(func: object) -> object:
    """装饰器：将 handler 调用整体包进文件锁。

    保证「读-改-写」周期（load -> modify -> save）在锁内完成，
    避免多个并发请求读取到同一版本后相互覆盖（丢失更新）。
    session 模式（内存文档）天然互斥，不加锁。
    """

    @functools.wraps(func)  # type: ignore[arg-type]
    def wrapper(*args, **kwargs):  # type: ignore[no-untyped-def]
        # filename 为第一个位置参数，或关键字参数
        filename = args[0] if args else kwargs.get("filename")
        session_id = kwargs.get("session_id")
        if session_id or not filename:
            return func(*args, **kwargs)

        from .path_guard import path_guard

        validated = path_guard.validate_path(filename, "write")
        file_lock_mgr.acquire(validated)
        try:
            return func(*args, **kwargs)
        finally:
            file_lock_mgr.release(validated)

    return wrapper  # type: ignore[return-value]
