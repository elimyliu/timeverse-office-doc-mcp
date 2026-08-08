"""统一错误处理 - ToolError 异常。"""

from __future__ import annotations

import logging

logger = logging.getLogger("timeverse_office_doc_mcp")


class ToolError(Exception):
    """工具执行错误，会被序列化为 MCP 错误响应返回给客户端。"""

    def __init__(self, message: str, *, tool: str | None = None) -> None:
        super().__init__(message)
        self.tool = tool
