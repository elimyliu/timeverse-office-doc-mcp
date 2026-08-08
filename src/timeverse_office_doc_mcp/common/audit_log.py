"""操作审计日志 - 记录所有文件操作。

对应方案 7.3 操作审计（AuditLog）。
"""

from __future__ import annotations

import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Any

logger = logging.getLogger("timeverse_office_doc_mcp.audit")


class AuditLogger:
    """操作审计日志 - JSON Lines 格式追加写入。"""

    def __init__(self, log_path: str | None = None) -> None:
        self.log_path = log_path or str(
            Path(__file__).resolve().parent.parent.parent.parent / "audit.log"
        )

    def log_operation(
        self,
        tool_name: str,
        args: dict[str, Any],
        result: str,
        duration_ms: int,
        success: bool,
        error: str | None = None,
    ) -> None:
        """记录操作日志。

        - 工具名
        - 参数摘要（脱敏）
        - 执行结果状态
        - 耗时
        - 错误信息（如有）
        """
        entry = {
            "timestamp": datetime.now().isoformat(),
            "tool": tool_name,
            "args_summary": self._summarize_args(args),
            "success": success,
            "duration_ms": duration_ms,
            "error": error,
        }
        try:
            Path(self.log_path).parent.mkdir(parents=True, exist_ok=True)
            with open(self.log_path, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        except OSError:
            logger.warning("Failed to write audit log to %s", self.log_path)

    def _summarize_args(self, args: dict[str, Any]) -> dict[str, Any]:
        """参数摘要：截断过长的值，避免日志膨胀。"""
        summary: dict[str, Any] = {}
        for key, value in args.items():
            if isinstance(value, str) and len(value) > 200:
                summary[key] = value[:200] + "..."
            elif isinstance(value, list) and len(value) > 10:
                summary[key] = f"[{len(value)} items]"
            else:
                summary[key] = value
        return summary


# 全局单例
audit_logger = AuditLogger()
