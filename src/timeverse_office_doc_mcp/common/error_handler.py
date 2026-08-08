"""统一错误处理 - ToolError 异常与工具错误装饰器。"""

from __future__ import annotations

import functools
import logging
from collections.abc import Callable
from typing import Any

logger = logging.getLogger("timeverse_office_doc_mcp")


class ToolError(Exception):
    """工具执行错误，会被序列化为 MCP 错误响应返回给客户端。"""

    def __init__(self, message: str, *, tool: str | None = None) -> None:
        super().__init__(message)
        self.tool = tool


def handle_tool_error(func: Callable[..., Any]) -> Callable[..., Any]:
    """同步工具错误装饰器：捕获异常并记录日志。"""

    @functools.wraps(func)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return func(*args, **kwargs)
        except ToolError:
            raise
        except Exception as e:
            logger.exception("Unhandled error in tool %s", func.__name__)
            raise ToolError(str(e), tool=func.__name__) from e

    return wrapper
