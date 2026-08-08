"""配置管理 - 路径白名单、大小限制、Session、传输方式等。

对应方案 7.4 安全配置 与 9.3 配置参数。
"""

from __future__ import annotations

import os
from pathlib import Path


def _get_project_root() -> Path:
    """获取项目根目录（src 的上一级）。"""
    return Path(__file__).resolve().parent.parent.parent


class ServerConfig:
    """服务运行配置（对应方案 9.3）。"""

    PROJECT_ROOT = _get_project_root()

    # 模板目录
    TEMPLATE_DIR = os.environ.get("OFFICE_TEMPLATES", str(PROJECT_ROOT / "templates"))


class SecurityConfig:
    """安全配置（对应方案 7.4）。

    白名单采用根目录前缀匹配（Path.is_relative_to），允许目录下的所有子目录与文件，
    天然覆盖用户动态创建的项目目录。
    """

    # 路径白名单（MCP 启动时配置，支持逗号分隔多目录）
    # 前缀匹配：允许目录下的所有子目录/文件，天然覆盖用户动态创建的项目目录
    ALLOWED_DIRS: list[str] = [
        d
        for d in (
            *os.environ.get("OFFICE_ALLOWED_DIRS", "").split(","),  # 自定义根目录（可多个）
            ServerConfig.TEMPLATE_DIR,
            os.getcwd(),  # 启动时工作目录兜底
        )
        if d
    ]

    # 路径黑名单
    BLOCKED_PATTERNS = [
        ".ssh",
        ".aws",
        ".config",
        ".env",
        "AppData",
        "Library",
        "/etc",
        "/sys",
    ]

    # 表格大小限制
    MAX_TABLE_ROWS = 1000
    MAX_TABLE_COLS = 100

    # 文本长度限制
    MAX_TEXT_LENGTH = 100000  # 10 万字符

    # 允许的文件扩展名
    ALLOWED_EXTENSIONS = {".docx", ".xlsx", ".pptx", ".pdf"}

    # 审计日志脱敏：掩码手机号/邮箱/身份证等敏感字段
    SANITIZE_ENABLED = os.environ.get("SANITIZE_ENABLED", "false").lower() == "true"
    SANITIZE_FIELDS = os.environ.get("SANITIZE_FIELDS", "phone,email,id_card,bank_card").split(",")


def ensure_dirs() -> None:
    """确保运行时目录存在。"""
    Path(ServerConfig.TEMPLATE_DIR).mkdir(parents=True, exist_ok=True)
