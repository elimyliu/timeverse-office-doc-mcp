"""路径安全守卫 - 防止路径遍历和越权访问。

对应方案 7.1 路径沙箱（PathGuard）。
"""

from __future__ import annotations

from pathlib import Path

from ..config import SecurityConfig
from .error_handler import ToolError

# 最大文件大小限制（字节），100MB
_MAX_FILE_SIZE = 100 * 1024 * 1024


class PathGuard:
    """路径安全守卫 - 所有文件操作限制在白名单目录内。"""

    def __init__(
        self,
        allowed_dirs: list[str] | None = None,
        blocked_patterns: list[str] | None = None,
        max_file_size: int | None = None,
    ) -> None:
        self.allowed_dirs = [
            Path(d).resolve() for d in (allowed_dirs or SecurityConfig.ALLOWED_DIRS)
        ]
        self.blocked_patterns = blocked_patterns or SecurityConfig.BLOCKED_PATTERNS
        self.max_file_size = max_file_size or _MAX_FILE_SIZE

    def validate_path(self, file_path: str, operation: str = "read") -> str:
        """校验文件路径是否安全。

        检查项：
        1. 路径规范化（resolve .. 和符号链接）
        2. 是否在允许目录内（前缀匹配，覆盖动态子目录）
        3. 是否匹配黑名单模式
        4. 写操作时检查文件扩展名白名单
        5. 文件大小是否超过限制
        """
        resolved = Path(file_path).resolve()

        # 检查是否在允许目录内（前缀匹配）
        if not self._is_allowed(resolved):
            raise ToolError(f"路径 '{file_path}' 不在允许的目录范围内")

        # 检查黑名单
        resolved_str = str(resolved)
        for pattern in self.blocked_patterns:
            if pattern in resolved_str:
                raise ToolError(f"路径 '{file_path}' 被安全策略阻止")

        # 写操作检查扩展名
        if operation == "write":
            allowed_exts = SecurityConfig.ALLOWED_EXTENSIONS
            if resolved.suffix.lower() not in allowed_exts:
                raise ToolError(
                    f"不支持的文件格式: {resolved.suffix}，允许: {', '.join(sorted(allowed_exts))}"
                )

        # 文件大小检查（仅对已存在的文件）
        if (
            resolved.exists()
            and resolved.is_file()
            and resolved.stat().st_size > self.max_file_size
        ):
            raise ToolError(f"文件大小超过限制 ({self.max_file_size // 1024 // 1024}MB)")

        return str(resolved)

    def validate_directory(self, dir_path: str) -> str:
        """校验目录路径是否在白名单内。"""
        resolved = Path(dir_path).resolve()

        if not self._is_allowed(resolved):
            raise ToolError(f"目录 '{dir_path}' 不在允许的目录范围内")

        for pattern in self.blocked_patterns:
            if pattern in str(resolved):
                raise ToolError(f"目录 '{dir_path}' 被安全策略阻止")

        return str(resolved)

    def _is_allowed(self, resolved: Path) -> bool:
        """检查路径是否在任一允许目录下（前缀匹配）。"""
        for allowed in self.allowed_dirs:
            try:
                if resolved.is_relative_to(allowed):
                    return True
            except ValueError:
                continue
        return False


# 全局单例
path_guard = PathGuard()
