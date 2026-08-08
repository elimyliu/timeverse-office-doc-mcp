"""测试跨格式 doc handler 工具（模板管理 + Session 管理）。"""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document

from timeverse_office_doc_mcp.common.error_handler import ToolError
from timeverse_office_doc_mcp.handlers import doc_handler, word_handler


@pytest.fixture
def workspace(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """创建临时工作空间并更新 PathGuard 白名单。"""
    ws = tmp_path / "workspace"
    ws.mkdir()
    from timeverse_office_doc_mcp.common.path_guard import path_guard

    monkeypatch.setattr(path_guard, "allowed_dirs", [tmp_path.resolve()])
    # 同时更新模板管理器的目录
    tpl_dir = tmp_path / "templates"
    tpl_dir.mkdir(exist_ok=True)
    from timeverse_office_doc_mcp.common.template_mgr import template_manager

    monkeypatch.setattr(template_manager, "template_dir", tpl_dir)
    monkeypatch.setattr(template_manager, "registry_path", tpl_dir / "registry.json")
    return ws


@pytest.fixture
def template_docx(workspace: Path) -> str:
    """创建一个带占位符的 Word 模板文件。"""
    path = str(workspace / "template.docx")
    doc = Document()
    doc.add_heading("{{title}}", level=1)
    doc.add_paragraph("作者: {{author}}")
    doc.add_paragraph("日期: {{date}}")
    doc.save(path)
    return path


class TestDocRegisterAndDelete:
    def test_register_and_list(self, template_docx: str) -> None:
        # 注册模板
        result = doc_handler.doc_register_template(
            name="test_tpl",
            format="word",
            file_path=template_docx,
            description="测试模板",
        )
        assert result["registered"] is True
        assert result["name"] == "test_tpl"

        # 列出模板
        listing = doc_handler.doc_list_templates(format="word")
        assert listing["count"] >= 1
        names = [t["name"] for t in listing["templates"]]
        assert "test_tpl" in names

    def test_get_template_info(self, template_docx: str) -> None:
        doc_handler.doc_register_template(
            name="info_tpl",
            format="word",
            file_path=template_docx,
            description="信息测试",
        )
        info = doc_handler.doc_get_template_info("info_tpl")
        assert info["name"] == "info_tpl"
        assert info["format"] == "word"
        assert len(info["placeholders"]) >= 2  # title, author, date

    def test_delete_template(self, template_docx: str) -> None:
        doc_handler.doc_register_template(
            name="del_tpl",
            format="word",
            file_path=template_docx,
        )
        doc_handler.doc_delete_template("del_tpl")
        listing = doc_handler.doc_list_templates()
        names = [t["name"] for t in listing["templates"]]
        assert "del_tpl" not in names


class TestDocApplyTemplate:
    def test_apply_word_template(self, template_docx: str, workspace: Path) -> None:
        # 注册模板
        doc_handler.doc_register_template(
            name="apply_tpl",
            format="word",
            file_path=template_docx,
        )

        # 应用模板
        output = str(workspace / "output.docx")
        result = doc_handler.doc_apply_template(
            template_name="apply_tpl",
            output_path=output,
            variables={"title": "测试报告", "author": "张三", "date": "2026-08-06"},
        )
        assert result["format"] == "word"
        assert result["variables_replaced"] >= 2
        assert Path(output).exists()

        # 验证内容
        doc = Document(output)
        text = "\n".join(p.text for p in doc.paragraphs)
        assert "测试报告" in text
        assert "张三" in text
        assert "{{" not in text


class TestDocExtractPlaceholders:
    def test_extract(self, template_docx: str) -> None:
        doc_handler.doc_register_template(
            name="ph_tpl",
            format="word",
            file_path=template_docx,
        )
        result = doc_handler.doc_extract_placeholders("ph_tpl")
        assert result["count"] >= 2
        names = [p["name"] for p in result["placeholders"]]
        assert "title" in names
        assert "author" in names


class TestDocSession:
    def test_open_save_close_session(self, workspace: Path) -> None:
        # 创建文档
        docx_path = str(workspace / "session_test.docx")
        word_handler.word_create_document(docx_path, title="Session 测试")

        # 打开 Session
        result = doc_handler.doc_open_session(docx_path, format="word")
        session_id = result["session_id"]
        assert result["format"] == "word"

        # 在 Session 中编辑
        word_handler.word_add_paragraph(docx_path, text="Session 段落", session_id=session_id)

        # 列出 Session
        listing = doc_handler.doc_list_sessions()
        assert listing["count"] >= 1

        # 保存 Session
        doc_handler.doc_save_session(session_id)
        # 实际写盘需要调用 document.save
        from timeverse_office_doc_mcp.common.session import session_manager

        doc = session_manager.get_session(session_id).document
        doc.save(docx_path)

        # 关闭 Session
        doc_handler.doc_close_session(session_id)

        # 验证 Session 已关闭
        listing = doc_handler.doc_list_sessions()
        assert listing["count"] == 0

        # 验证内容已保存
        doc2 = Document(docx_path)
        text = "\n".join(p.text for p in doc2.paragraphs)
        assert "Session 段落" in text

    def test_close_without_save(self, workspace: Path) -> None:
        docx_path = str(workspace / "nosave.docx")
        word_handler.word_create_document(docx_path, title="不保存测试")
        result = doc_handler.doc_open_session(docx_path, format="word")
        session_id = result["session_id"]
        doc_handler.doc_close_session(session_id, save=False)
        listing = doc_handler.doc_list_sessions()
        assert listing["count"] == 0
