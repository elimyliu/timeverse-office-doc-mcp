"""测试 PathGuard 路径沙箱。"""

from __future__ import annotations

from pathlib import Path

import pytest

from timeverse_office_doc_mcp.common.error_handler import ToolError
from timeverse_office_doc_mcp.common.path_guard import PathGuard


@pytest.fixture
def guard(tmp_path: Path) -> PathGuard:
    """创建一个限定在 tmp_path 的 PathGuard。"""
    return PathGuard(allowed_dirs=[str(tmp_path)], blocked_patterns=[".ssh", ".env"])


class TestPathGuardValidatePath:
    """测试路径校验。"""

    def test_allowed_read(self, guard: PathGuard, tmp_path: Path) -> None:
        """白名单内的文件读取应通过。"""
        f = tmp_path / "test.docx"
        f.write_text("test")
        result = guard.validate_path(str(f), "read")
        assert result == str(f.resolve())

    def test_blocked_outside_whitelist(self, guard: PathGuard, tmp_path: Path) -> None:
        """白名单外的路径应被拒绝。"""
        outside = tmp_path.parent / "outside.docx"
        with pytest.raises(ToolError, match="不在允许的目录范围内"):
            guard.validate_path(str(outside), "read")

    def test_blocked_pattern(self, guard: PathGuard, tmp_path: Path) -> None:
        """匹配黑名单的路径应被拒绝。"""
        secret = tmp_path / ".env"
        secret.write_text("secret")
        with pytest.raises(ToolError, match="安全策略阻止"):
            guard.validate_path(str(secret), "read")

    def test_write_extension_check(self, guard: PathGuard, tmp_path: Path) -> None:
        """写操作应检查扩展名白名单。"""
        bad = tmp_path / "test.txt"
        with pytest.raises(ToolError, match="不支持的文件格式"):
            guard.validate_path(str(bad), "write")

    def test_write_docx_allowed(self, guard: PathGuard, tmp_path: Path) -> None:
        """写 .docx 应通过。"""
        f = tmp_path / "test.docx"
        result = guard.validate_path(str(f), "write")
        assert result == str(f.resolve())

    def test_subdirectory_allowed(self, guard: PathGuard, tmp_path: Path) -> None:
        """白名单根目录下的子目录应自动允许。"""
        sub = tmp_path / "subdir" / "deep" / "test.docx"
        sub.parent.mkdir(parents=True)
        sub.write_text("test")
        result = guard.validate_path(str(sub), "read")
        assert result == str(sub.resolve())

    def test_path_traversal_blocked(self, guard: PathGuard, tmp_path: Path) -> None:
        """路径遍历（..）应被规范化后拦截。"""
        traversal = str(tmp_path / ".." / "outside.docx")
        with pytest.raises(ToolError):
            guard.validate_path(traversal, "read")


class TestPathGuardValidateDirectory:
    """测试目录校验。"""

    def test_allowed_directory(self, guard: PathGuard, tmp_path: Path) -> None:
        """白名单内目录应通过。"""
        sub = tmp_path / "subdir"
        sub.mkdir()
        result = guard.validate_directory(str(sub))
        assert result == str(sub.resolve())

    def test_blocked_directory(self, guard: PathGuard, tmp_path: Path) -> None:
        """白名单外目录应被拒绝。"""
        with pytest.raises(ToolError, match="不在允许的目录范围内"):
            guard.validate_directory("/etc")
