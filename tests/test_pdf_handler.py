"""测试 PDF handler 工具。"""

from __future__ import annotations

from pathlib import Path

import pytest
from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from timeverse_office_doc_mcp.common.error_handler import ToolError
from timeverse_office_doc_mcp.handlers import pdf_handler


@pytest.fixture
def workspace(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """创建临时工作空间并更新 PathGuard 白名单。"""
    ws = tmp_path / "workspace"
    ws.mkdir()
    from timeverse_office_doc_mcp.common.path_guard import path_guard

    monkeypatch.setattr(path_guard, "allowed_dirs", [tmp_path.resolve()])
    return ws


@pytest.fixture
def pdf_path(workspace: Path) -> str:
    """创建一个测试 PDF 文件。"""
    path = str(workspace / "test.pdf")
    c = canvas.Canvas(path, pagesize=A4)
    c.drawString(72, 750, "Hello World 测试文本")
    c.drawString(72, 700, "第二行内容")
    c.showPage()
    c.drawString(72, 750, "第二页内容")
    c.showPage()
    c.save()
    return path


class TestPdfGetInfo:
    def test_get_info(self, pdf_path: str) -> None:
        result = pdf_handler.pdf_get_info(pdf_path)
        assert result["page_count"] == 2
        assert result["encrypted"] is False

    def test_get_info_analyze(self, pdf_path: str) -> None:
        result = pdf_handler.pdf_get_info(pdf_path, analyze=True)
        assert result["page_count"] == 2
        assert result["total_chars"] > 0
        assert len(result["pages"]) == 2


class TestPdfExtractText:
    def test_extract_all(self, pdf_path: str) -> None:
        result = pdf_handler.pdf_extract_text(pdf_path)
        assert result["page_count"] == 2
        assert "Hello" in result["pages"][0]["text"]

    def test_extract_range(self, pdf_path: str) -> None:
        result = pdf_handler.pdf_extract_text(pdf_path, page_range="1")
        assert result["page_count"] == 1


class TestPdfSearchText:
    def test_search(self, pdf_path: str) -> None:
        result = pdf_handler.pdf_search_text(pdf_path, query="Hello")
        assert result["match_count"] >= 1

    def test_search_no_match(self, pdf_path: str) -> None:
        result = pdf_handler.pdf_search_text(pdf_path, query="不存在的文本")
        assert result["match_count"] == 0


class TestPdfRotatePage:
    def test_rotate(self, pdf_path: str) -> None:
        result = pdf_handler.pdf_rotate_page(pdf_path, page_idx=0, angle=90)
        assert result["angle"] == 90

    def test_invalid_angle(self, pdf_path: str) -> None:
        with pytest.raises(ToolError):
            pdf_handler.pdf_rotate_page(pdf_path, page_idx=0, angle=45)


class TestPdfAddText:
    def test_add_text(self, pdf_path: str) -> None:
        result = pdf_handler.pdf_add_text(pdf_path, page_idx=0, text="NewText123", x=100, y=100)
        assert result["text"] == "NewText123"
        # 验证文本已添加
        text_result = pdf_handler.pdf_extract_text(pdf_path)
        assert "NewText123" in text_result["pages"][0]["text"]


class TestPdfAddWatermark:
    def test_add_watermark(self, pdf_path: str) -> None:
        result = pdf_handler.pdf_add_watermark(pdf_path, watermark_text="机密")
        assert result["watermark"] == "机密"
        assert result["pages_watermarked"] == 2


class TestPdfAddAnnotation:
    def test_add_bookmark(self, pdf_path: str) -> None:
        result = pdf_handler.pdf_add_annotation(pdf_path, page_idx=0, annotation_type="bookmark", content="第一章")
        assert result["annotation_type"] == "bookmark"
        assert result["title"] == "第一章"


class TestPdfMerge:
    def test_merge(self, pdf_path: str, workspace: Path) -> None:
        # 创建第二个 PDF
        path2 = str(workspace / "test2.pdf")
        c = canvas.Canvas(path2, pagesize=A4)
        c.drawString(72, 750, "第二个 PDF")
        c.showPage()
        c.save()

        output = str(workspace / "merged.pdf")
        result = pdf_handler.pdf_merge([pdf_path, path2], output)
        assert result["files_merged"] == 2
        assert result["total_pages"] == 3


class TestPdfSplit:
    def test_split(self, pdf_path: str, workspace: Path) -> None:
        prefix = str(workspace / "split")
        result = pdf_handler.pdf_split(pdf_path, page_ranges="1,2", output_prefix=prefix)
        assert result["count"] == 2
        assert Path(f"{prefix}_1.pdf").exists()
        assert Path(f"{prefix}_2.pdf").exists()


class TestPdfManageSecurity:
    def test_encrypt_and_decrypt(self, pdf_path: str) -> None:
        # 加密
        pdf_handler.pdf_manage_security(pdf_path, action="encrypt", password="secret123")
        reader = PdfReader(pdf_path)
        assert reader.is_encrypted is True

        # 解密
        pdf_handler.pdf_manage_security(pdf_path, action="decrypt", password="secret123")
        info = pdf_handler.pdf_get_info(pdf_path)
        assert info["encrypted"] is False
