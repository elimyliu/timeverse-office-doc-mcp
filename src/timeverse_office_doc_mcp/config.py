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

    # 传输方式: stdio（Phase 1 仅支持 stdio）
    TRANSPORT = os.environ.get("MCP_TRANSPORT", "stdio")

    # 工作目录
    WORKSPACE_DIR = os.environ.get("OFFICE_WORKSPACE", str(PROJECT_ROOT / "workspace"))
    OUTPUT_DIR = os.environ.get("OFFICE_OUTPUT", str(PROJECT_ROOT / "output"))

    # 模板目录
    TEMPLATE_DIR = os.environ.get("OFFICE_TEMPLATES", str(PROJECT_ROOT / "templates"))

    # Session 配置
    SESSION_ENABLED = os.environ.get("SESSION_ENABLED", "true").lower() == "true"
    SESSION_TTL = int(os.environ.get("SESSION_TTL", "3600"))  # 1 小时

    # 日志级别
    LOG_LEVEL = os.environ.get("LOG_LEVEL", "INFO")

    # 按需启用格式
    ENABLED_FORMATS = os.environ.get("ENABLED_FORMATS", "word,excel,ppt,pdf").split(",")

    # 可选: 启用 PyMuPDF 高性能模式
    USE_PYMUPDF = os.environ.get("USE_PYMUPDF", "false").lower() == "true"


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
            os.environ.get("OFFICE_WORKSPACE", str(ServerConfig.PROJECT_ROOT / "workspace")),
            os.environ.get("OFFICE_OUTPUT", str(ServerConfig.PROJECT_ROOT / "output")),
            os.environ.get("OFFICE_TEMPLATES", str(ServerConfig.PROJECT_ROOT / "templates")),
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

    # 文件大小限制
    MAX_FILE_SIZE = int(os.environ.get("MAX_FILE_SIZE", str(100 * 1024 * 1024)))  # 100MB

    # 表格大小限制
    MAX_TABLE_ROWS = 1000
    MAX_TABLE_COLS = 100

    # 文本长度限制
    MAX_TEXT_LENGTH = 100000  # 10 万字符

    # 操作超时
    OPERATION_TIMEOUT = int(os.environ.get("OPERATION_TIMEOUT", "60"))  # 秒

    # 允许的文件扩展名
    ALLOWED_EXTENSIONS = {".docx", ".xlsx", ".pptx", ".pdf"}

    # HTTP/SSE 认证（stdio 模式无需配置）
    AUTH_ENABLED = os.environ.get("AUTH_ENABLED", "false").lower() == "true"
    API_KEYS = os.environ.get("API_KEYS", "").split(",") if AUTH_ENABLED else []

    # 速率限制
    RATE_LIMIT_ENABLED = os.environ.get("RATE_LIMIT_ENABLED", "true").lower() == "true"
    RATE_LIMIT_MAX = int(os.environ.get("RATE_LIMIT_MAX", "60"))  # 每分钟最大请求数

    # 敏感数据脱敏
    SANITIZE_ENABLED = os.environ.get("SANITIZE_ENABLED", "false").lower() == "true"
    SANITIZE_FIELDS = os.environ.get("SANITIZE_FIELDS", "phone,email,id_card,bank_card").split(",")


def ensure_dirs() -> None:
    """确保运行时目录存在。"""
    for d in (ServerConfig.WORKSPACE_DIR, ServerConfig.OUTPUT_DIR, ServerConfig.TEMPLATE_DIR):
        Path(d).mkdir(parents=True, exist_ok=True)
