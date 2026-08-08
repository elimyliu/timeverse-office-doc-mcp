"""Session 管理 - 内存编辑避免频繁磁盘 IO。

对应方案 5.6 Session 管理工具集。
当 session_id 存在时，编辑工具从内存中获取文档对象直接操作；
不存在时走传统的「打开-修改-保存」流程，对 LLM 透明。
"""

from __future__ import annotations

import logging
import time
import uuid
from dataclasses import dataclass, field
from typing import Any

from .error_handler import ToolError

logger = logging.getLogger("timeverse_office_doc_mcp.session")

# Session 默认过期时间（秒），1 小时无访问自动清理
_DEFAULT_SESSION_TTL = 3600


@dataclass
class Session:
    """一个内存中的文档编辑会话。"""

    session_id: str
    filename: str
    format: str  # word / excel / ppt / pdf
    document: Any  # python-docx Document / openpyxl Workbook / etc.
    modified: bool = False
    created_at: float = field(default_factory=time.time)
    last_accessed: float = field(default_factory=time.time)


class SessionManager:
    """Session 管理器 - 在内存中维护打开的文档。"""

    def __init__(self, ttl: int | None = None) -> None:
        self.ttl = ttl or _DEFAULT_SESSION_TTL
        self._sessions: dict[str, Session] = {}

    def open_session(self, filename: str, format: str, document: Any) -> str:
        """打开文档到内存 Session，返回 session_id。"""
        session_id = uuid.uuid4().hex[:12]
        self._sessions[session_id] = Session(
            session_id=session_id,
            filename=filename,
            format=format,
            document=document,
        )
        logger.info("Opened session %s for %s (%s)", session_id, filename, format)
        return session_id

    def get_session(self, session_id: str) -> Session:
        """获取 Session，不存在则报错。"""
        session = self._sessions.get(session_id)
        if session is None:
            raise ToolError(f"Session 不存在或已过期: {session_id}")
        session.last_accessed = time.time()
        return session

    def get_document(self, session_id: str, expected_format: str | None = None) -> Any:
        """获取 Session 中的文档对象。

        expected_format 不为 None 时校验格式是否匹配，不匹配则抛出 ToolError。
        """
        session = self.get_session(session_id)
        if expected_format and session.format != expected_format:
            raise ToolError(
                f"Session {session_id} 不是 {expected_format} 文档（当前格式: {session.format}）"
            )
        return session.document

    def mark_modified(self, session_id: str) -> None:
        """标记 Session 已修改。"""
        session = self.get_session(session_id)
        session.modified = True

    def save_session(self, session_id: str, output_path: str | None = None) -> str:
        """保存 Session 到磁盘，返回保存路径。"""
        session = self.get_session(session_id)
        save_path = output_path or session.filename
        # 实际保存由各 handler 的 document.save() 完成
        # 这里仅更新状态
        session.modified = False
        session.filename = save_path
        logger.info("Saved session %s to %s", session_id, save_path)
        return save_path

    def close_session(self, session_id: str, save: bool = False) -> None:
        """关闭 Session 并释放内存。"""
        session = self._sessions.get(session_id)
        if session is None:
            raise ToolError(f"Session 不存在或已关闭: {session_id}")
        if save:
            self.save_session(session_id)
        self._sessions.pop(session_id, None)
        logger.info("Closed session %s", session_id)

    def list_sessions(self) -> list[dict[str, Any]]:
        """列出所有活跃 Session。"""
        self._cleanup_expired()
        return [
            {
                "session_id": s.session_id,
                "filename": s.filename,
                "format": s.format,
                "modified": s.modified,
                "created_at": s.created_at,
            }
            for s in self._sessions.values()
        ]

    def _cleanup_expired(self) -> None:
        """清理过期的 Session（TTL 到期）。"""
        now = time.time()
        expired = [sid for sid, s in self._sessions.items() if now - s.last_accessed > self.ttl]
        for sid in expired:
            self._sessions.pop(sid, None)
            logger.info("Expired session %s (TTL %ss)", sid, self.ttl)


# 全局单例
session_manager = SessionManager()
