"""操作审计日志 - 记录所有文件操作。

对应方案 7.3 操作审计（AuditLog）。
"""

from __future__ import annotations

import json
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from ..config import SecurityConfig

logger = logging.getLogger("timeverse_office_doc_mcp.audit")

# 敏感字段对应的正则模式
_SANITIZE_PATTERNS: dict[str, re.Pattern[str]] = {
    "phone": re.compile(r"1[3-9]\d{9}"),
    "email": re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+"),
    "id_card": re.compile(r"\d{17}[\dXx]"),
    "bank_card": re.compile(r"\d{16,19}"),
}


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
        """参数摘要：截断过长的值，避免日志膨胀；按配置脱敏敏感字段。"""
        summary: dict[str, Any] = {}
        for key, value in args.items():
            if isinstance(value, str) and len(value) > 200:
                summary[key] = value[:200] + "..."
            elif isinstance(value, list) and len(value) > 10:
                summary[key] = f"[{len(value)} items]"
            else:
                summary[key] = value
        if SecurityConfig.SANITIZE_ENABLED:
            summary = self._sanitize(summary)
        return summary

    def _sanitize(self, data: dict[str, Any]) -> dict[str, Any]:
        """对参数值中的敏感信息进行掩码处理。"""
        fields = SecurityConfig.SANITIZE_FIELDS
        for key, value in data.items():
            if not isinstance(value, str):
                continue
            for field in fields:
                pattern = _SANITIZE_PATTERNS.get(field)
                if pattern:
                    value = pattern.sub(lambda m: m.group()[:3] + "***", value)
            data[key] = value
        return data


# 全局单例
audit_logger = AuditLogger()
