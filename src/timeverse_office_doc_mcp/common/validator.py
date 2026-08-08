"""统一输入校验器。

对应方案 7.2 输入校验（Validator）。
"""

from __future__ import annotations

import re

from ..config import SecurityConfig
from .error_handler import ToolError


class InputValidator:
    """统一输入校验器 - 所有工具参数的入口校验。"""

    @staticmethod
    def validate_filename(filename: str) -> str:
        """文件名安全校验：禁止特殊字符、控制长度。"""
        if not filename or len(filename) > 255:
            raise ToolError("文件名无效或过长")
        if any(c in filename for c in ["..", "\x00", "\n", "\r"]):
            raise ToolError("文件名包含非法字符")
        return filename

    @staticmethod
    def validate_range(range_str: str) -> str:
        """Excel 区域字符串校验：如 A1:C10。"""
        if not re.match(r"^[A-Z]+\d+:[A-Z]+\d+$", range_str):
            raise ToolError(f"无效的区域格式: {range_str}")
        return range_str

    @staticmethod
    def validate_table_data(
        data: list[list[str]],
        max_rows: int = SecurityConfig.MAX_TABLE_ROWS,
        max_cols: int = SecurityConfig.MAX_TABLE_COLS,
    ) -> list[list[str]]:
        """表格数据校验：限制行列数。"""
        if len(data) > max_rows:
            raise ToolError(f"表格行数超过限制 ({max_rows})")
        for row in data:
            if len(row) > max_cols:
                raise ToolError(f"表格列数超过限制 ({max_cols})")
        return data

    @staticmethod
    def validate_text_length(text: str, max_length: int = SecurityConfig.MAX_TEXT_LENGTH) -> str:
        """文本长度校验。"""
        if len(text) > max_length:
            raise ToolError(f"文本长度超过限制 ({max_length} 字符)")
        return text

    @staticmethod
    def validate_positive_int(value: int, name: str = "value") -> int:
        """校验正整数。"""
        if not isinstance(value, int) or value < 0:
            raise ToolError(f"{name} 必须是非负整数，得到: {value}")
        return value

    @staticmethod
    def validate_choice(value: str, choices: list[str], name: str = "value") -> str:
        """校验枚举值。"""
        if value not in choices:
            raise ToolError(f"{name} 必须是 {choices} 之一，得到: {value}")
        return value
