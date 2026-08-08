"""PDF 文档处理器 - 19 个工具。

对应方案 5.4 PDF 工具集。使用 pdfplumber + pypdf + reportlab 实现。
"""

from __future__ import annotations

import contextlib
import io
import logging
from pathlib import Path
from typing import Any

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from ..common.error_handler import ToolError
from ..common.file_lock import file_lock_mgr
from ..common.path_guard import path_guard
from ..common.validator import InputValidator

logger = logging.getLogger("timeverse_office_doc_mcp.pdf")


def _validate_pdf(filename: str) -> str:
    """校验 PDF 文件路径。"""
    return path_guard.validate_path(filename, "read")


def _save_pdf(writer: PdfWriter, output: str) -> None:
    """保存 PDF。"""
    validated = path_guard.validate_path(output, "write")
    file_lock_mgr.acquire(validated)
    try:
        with open(validated, "wb") as f:
            writer.write(f)
    finally:
        file_lock_mgr.release(validated)


# ==================== 5.4.1 文档管理（4 个） ====================


def pdf_get_info(filename: str) -> dict[str, Any]:
    """获取 PDF 元信息（页数、作者、标题等）。"""
    validated = _validate_pdf(filename)
    reader = PdfReader(validated)
    encrypted = reader.is_encrypted
    meta: dict[str, Any] = {}
    if not encrypted:
        try:
            meta = reader.metadata or {}
        except Exception:
            meta = {}
    return {
        "filename": filename,
        "page_count": len(reader.pages),
        "title": str(meta.get("/Title", "")) if meta else "",
        "author": str(meta.get("/Author", "")) if meta else "",
        "subject": str(meta.get("/Subject", "")) if meta else "",
        "creator": str(meta.get("/Creator", "")) if meta else "",
        "encrypted": encrypted,
    }


def pdf_merge(files: list[str], output: str) -> dict[str, Any]:
    """合并多个 PDF。"""
    if not files:
        raise ToolError("文件列表不能为空")
    writer = PdfWriter()
    merged_count = 0
    for f in files:
        validated = _validate_pdf(f)
        reader = PdfReader(validated)
        for page in reader.pages:
            writer.add_page(page)
            merged_count += 1
    _save_pdf(writer, output)
    return {"output": output, "files_merged": len(files), "total_pages": merged_count}


def pdf_split(filename: str, page_ranges: str, output_prefix: str) -> dict[str, Any]:
    """拆分 PDF。page_ranges 格式: '1-3,4-6,7-9'。"""
    validated = _validate_pdf(filename)
    reader = PdfReader(validated)
    total_pages = len(reader.pages)
    ranges = _parse_page_ranges(page_ranges, total_pages)
    output_files = []
    for idx, (start, end) in enumerate(ranges):
        writer = PdfWriter()
        for p in range(start - 1, end):
            writer.add_page(reader.pages[p])
        out_path = f"{output_prefix}_{idx + 1}.pdf"
        _save_pdf(writer, out_path)
        output_files.append(out_path)
    return {"output_files": output_files, "count": len(output_files)}


def pdf_rotate_page(filename: str, page_idx: int, angle: int) -> dict[str, Any]:
    """旋转页面（angle: 90/180/270）。"""
    if angle not in (90, 180, 270):
        raise ToolError(f"旋转角度必须是 90/180/270，得到: {angle}")
    validated = _validate_pdf(filename)
    reader = PdfReader(validated)
    if page_idx < 0 or page_idx >= len(reader.pages):
        raise ToolError(f"页面索引超出范围: {page_idx}（共 {len(reader.pages)} 页）")
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        if i == page_idx:
            page.rotate(angle)
        writer.add_page(page)
    _save_pdf(writer, validated)
    return {"filename": filename, "page_idx": page_idx, "angle": angle}


def _parse_page_ranges(ranges_str: str, total: int) -> list[tuple[int, int]]:
    """解析页码范围字符串，返回 [(start, end), ...]。"""
    result = []
    for part in ranges_str.split(","):
        part = part.strip()
        if "-" in part:
            s, e = part.split("-", 1)
            start, end = int(s), int(e)
        else:
            start = end = int(part)
        if start < 1 or end > total or start > end:
            raise ToolError(f"页码范围无效: {part}（总页数 {total}）")
        result.append((start, end))
    return result


# ==================== 5.4.2 内容读取（5 个） ====================


def pdf_extract_text(
    filename: str, page_range: str | None = None, layout_mode: bool = False
) -> dict[str, Any]:
    """提取文本（支持指定页范围）。"""
    validated = _validate_pdf(filename)
    pages_text: list[dict[str, Any]] = []
    with pdfplumber.open(validated) as pdf:
        start, end = _resolve_range(page_range, len(pdf.pages))
        for i in range(start, end):
            page = pdf.pages[i]
            text = page.extract_text(layout=layout_mode) or ""
            pages_text.append({"page": i + 1, "text": text})
    return {"filename": filename, "pages": pages_text, "page_count": len(pages_text)}


def pdf_extract_tables(
    filename: str, page_range: str | None = None, format: str = "json"
) -> dict[str, Any]:
    """提取表格数据。"""
    InputValidator.validate_choice(format, ["json", "csv"], "format")
    validated = _validate_pdf(filename)
    all_tables: list[dict[str, Any]] = []
    with pdfplumber.open(validated) as pdf:
        start, end = _resolve_range(page_range, len(pdf.pages))
        for i in range(start, end):
            page = pdf.pages[i]
            tables = page.extract_tables()
            for t_idx, table in enumerate(tables):
                if format == "csv":
                    csv_lines = [",".join(f'"{c or ""}"' for c in row) for row in table]
                    all_tables.append(
                        {"page": i + 1, "table_idx": t_idx, "csv": "\n".join(csv_lines)}
                    )
                else:
                    all_tables.append({"page": i + 1, "table_idx": t_idx, "data": table})
    return {"filename": filename, "tables": all_tables, "count": len(all_tables)}


def pdf_extract_images(
    filename: str, page_range: str | None = None, output_dir: str | None = None
) -> dict[str, Any]:
    """提取图片。"""
    validated = _validate_pdf(filename)
    out_dir = output_dir or str(Path(validated).parent / f"{Path(validated).stem}_images")
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    extracted: list[str] = []
    with pdfplumber.open(validated) as pdf:
        start, end = _resolve_range(page_range, len(pdf.pages))
        for i in range(start, end):
            page = pdf.pages[i]
            for img_idx, img in enumerate(page.images):
                try:
                    cropped = page.crop((img["x0"], img["y0"], img["x1"], img["y1"]))
                    im = cropped.to_image()
                    out_path = str(Path(out_dir) / f"page{i + 1}_img{img_idx + 1}.png")
                    im.save(out_path)
                    extracted.append(out_path)
                except Exception as e:
                    logger.warning(
                        "Failed to extract image on page %d, img %d: %s", i + 1, img_idx + 1, e
                    )
    return {"filename": filename, "output_dir": out_dir, "images_extracted": len(extracted)}


def pdf_search_text(filename: str, query: str, case_sensitive: bool = False) -> dict[str, Any]:
    """搜索文本。"""
    validated = _validate_pdf(filename)
    InputValidator.validate_text_length(query)
    matches: list[dict[str, Any]] = []
    search_query = query if case_sensitive else query.lower()
    with pdfplumber.open(validated) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            search_text = text if case_sensitive else text.lower()
            idx = 0
            while True:
                pos = search_text.find(search_query, idx)
                if pos == -1:
                    break
                matches.append(
                    {
                        "page": i + 1,
                        "position": pos,
                        "context": text[max(0, pos - 20) : pos + len(query) + 20],
                    }
                )
                idx = pos + 1
    return {"filename": filename, "query": query, "match_count": len(matches), "matches": matches}


def pdf_ocr_text(
    filename: str,
    page_range: str | None = None,
    lang: str = "chi_sim+eng",
    output_format: str = "text",
) -> dict[str, Any]:
    """OCR 文本识别（扫描件/图片型 PDF）。需要系统安装 tesseract。"""
    try:
        import pytesseract  # noqa: F401
    except ImportError:
        raise ToolError(
            "OCR 功能需要安装 pytesseract 和 tesseract。请运行: pip install pytesseract"
        ) from None

    validated = _validate_pdf(filename)
    pages_text: list[dict[str, Any]] = []
    with pdfplumber.open(validated) as pdf:
        start, end = _resolve_range(page_range, len(pdf.pages))
        for i in range(start, end):
            page = pdf.pages[i]
            im = page.to_image(resolution=300)
            import pytesseract
            from PIL import Image

            img = Image.open(io.BytesIO(im.original.data))
            text = pytesseract.image_to_string(img, lang=lang)
            pages_text.append({"page": i + 1, "text": text.strip()})
    return {"filename": filename, "lang": lang, "pages": pages_text}


# ==================== 5.4.3 内容写入（5 个） ====================


def pdf_add_text(
    filename: str,
    page_idx: int,
    text: str,
    x: float = 72,
    y: float = 72,
    font: str = "Helvetica",
    font_size: int = 12,
    output: str | None = None,
) -> dict[str, Any]:
    """添加文本到页面（Overlay 合并模式）。"""
    validated = _validate_pdf(filename)
    reader = PdfReader(validated)
    if page_idx < 0 or page_idx >= len(reader.pages):
        raise ToolError(f"页面索引超出范围: {page_idx}")
    InputValidator.validate_text_length(text)

    # 生成 overlay
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    can.setFont(font, font_size)
    can.drawString(x, y, text)
    can.save()
    packet.seek(0)

    overlay_reader = PdfReader(packet)
    page = reader.pages[page_idx]
    page.merge_page(overlay_reader.pages[0])

    writer = PdfWriter()
    for p in reader.pages:
        writer.add_page(p)
    out = output or validated
    _save_pdf(writer, out)
    return {"filename": out, "page_idx": page_idx, "text": text}


def pdf_add_image(
    filename: str,
    page_idx: int,
    image_path: str,
    x: float = 72,
    y: float = 72,
    width: float = 200,
    height: float = 150,
    output: str | None = None,
) -> dict[str, Any]:
    """添加图片到页面（Overlay 合并模式）。"""
    validated = _validate_pdf(filename)
    img_validated = path_guard.validate_path(image_path, "read")
    if not Path(img_validated).exists():
        raise ToolError(f"图片文件不存在: {image_path}")
    reader = PdfReader(validated)
    if page_idx < 0 or page_idx >= len(reader.pages):
        raise ToolError(f"页面索引超出范围: {page_idx}")

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    can.drawImage(img_validated, x, y, width=width, height=height)
    can.save()
    packet.seek(0)

    overlay_reader = PdfReader(packet)
    page = reader.pages[page_idx]
    page.merge_page(overlay_reader.pages[0])

    writer = PdfWriter()
    for p in reader.pages:
        writer.add_page(p)
    out = output or validated
    _save_pdf(writer, out)
    return {"filename": out, "page_idx": page_idx, "image_path": img_validated}


def pdf_add_watermark(
    filename: str,
    watermark_text: str,
    opacity: float = 0.3,
    font_size: int = 60,
    output: str | None = None,
) -> dict[str, Any]:
    """添加水印（遍历所有页面）。"""
    validated = _validate_pdf(filename)
    reader = PdfReader(validated)

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    can.setFillAlpha(opacity)
    can.setFont("Helvetica", font_size)
    # 旋转 45 度居中
    can.translate(A4[0] / 2, A4[1] / 2)
    can.rotate(45)
    can.drawCentredString(0, 0, watermark_text)
    can.save()
    packet.seek(0)

    overlay_reader = PdfReader(packet)
    writer = PdfWriter()
    for page in reader.pages:
        page.merge_page(overlay_reader.pages[0])
        writer.add_page(page)

    out = output or validated
    _save_pdf(writer, out)
    return {"filename": out, "watermark": watermark_text, "pages_watermarked": len(reader.pages)}


def pdf_add_annotation(
    filename: str,
    page_idx: int,
    annotation_type: str,
    content: str,
    x: float = 72,
    y: float = 72,
    output: str | None = None,
) -> dict[str, Any]:
    """添加注释（annotation_type: highlight/text/link）。"""
    InputValidator.validate_choice(
        annotation_type, ["highlight", "text", "link"], "annotation_type"
    )
    validated = _validate_pdf(filename)
    reader = PdfReader(validated)
    if page_idx < 0 or page_idx >= len(reader.pages):
        raise ToolError(f"页面索引超出范围: {page_idx}")

    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        if i == page_idx:
            from pypdf.annotations import Text

            annot = Text(text=content)
            annot.rect = (x, y, x + 100, y + 20)
            writer.add_page(page)
            writer.add_annotation(page, annot)
        else:
            writer.add_page(page)

    out = output or validated
    _save_pdf(writer, out)
    return {"filename": out, "page_idx": page_idx, "annotation_type": annotation_type}


def pdf_add_bookmark(
    filename: str, title: str, page_idx: int, output: str | None = None
) -> dict[str, Any]:
    """添加书签。"""
    validated = _validate_pdf(filename)
    reader = PdfReader(validated)
    if page_idx < 0 or page_idx >= len(reader.pages):
        raise ToolError(f"页面索引超出范围: {page_idx}")

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.add_outline_item(title, page_idx)

    out = output or validated
    _save_pdf(writer, out)
    return {"filename": out, "title": title, "page_idx": page_idx}


# ==================== 5.4.4 安全与分析（3 个） ====================


def pdf_encrypt(
    filename: str, password: str, permissions: list[str] | None = None, output: str | None = None
) -> dict[str, Any]:
    """加密 PDF。"""
    validated = _validate_pdf(filename)
    reader = PdfReader(validated)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    writer.encrypt(password)
    out = output or validated
    _save_pdf(writer, out)
    return {"filename": out, "encrypted": True}


def pdf_decrypt(filename: str, password: str, output: str | None = None) -> dict[str, Any]:
    """解密 PDF。"""
    validated = _validate_pdf(filename)
    reader = PdfReader(validated)
    if reader.is_encrypted and not reader.decrypt(password):
        raise ToolError("密码错误，解密失败")

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    out = output or validated
    _save_pdf(writer, out)
    return {"filename": out, "decrypted": True}


def pdf_analyze_structure(filename: str) -> dict[str, Any]:
    """分析结构（页面类型、文本密度、表格分布）。"""
    validated = _validate_pdf(filename)
    page_analyses: list[dict[str, Any]] = []
    total_chars = 0
    total_tables = 0
    total_images = 0

    with pdfplumber.open(validated) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            tables = page.extract_tables()
            images = page.images
            char_count = len(text)
            total_chars += char_count
            total_tables += len(tables)
            total_images += len(images)

            # 推断页面类型
            if char_count < 50 and len(images) > 0:
                page_type = "image"
            elif len(tables) > 0:
                page_type = "table"
            elif char_count > 1000:
                page_type = "text-heavy"
            else:
                page_type = "text"

            page_analyses.append(
                {
                    "page": i + 1,
                    "type": page_type,
                    "char_count": char_count,
                    "table_count": len(tables),
                    "image_count": len(images),
                    "width": float(page.width),
                    "height": float(page.height),
                }
            )

    return {
        "filename": filename,
        "page_count": len(page_analyses),
        "total_chars": total_chars,
        "total_tables": total_tables,
        "total_images": total_images,
        "avg_chars_per_page": total_chars // max(len(page_analyses), 1),
        "pages": page_analyses,
    }


# ==================== 5.4.5 模板工具（2 个） ====================


def pdf_create_from_template(
    template_name: str,
    variables: dict[str, Any] | None = None,
    output: str = "output.pdf",
) -> dict[str, Any]:
    """从模板创建 PDF（reportlab 布局模板，变量填充）。"""
    from ..common.template_mgr import template_manager

    fmt, tpl_path = template_manager.resolve_template_path(template_name)
    if fmt != "pdf":
        raise ToolError(f"模板 '{template_name}' 是 {fmt} 格式，不是 pdf")

    validated_out = path_guard.validate_path(output, "write")
    variables = variables or {}

    # 检查是否是 AcroForm 表单 PDF
    reader = PdfReader(tpl_path)
    if reader.get_fields():
        # AcroForm 表单填充
        writer = PdfWriter()
        writer.append(reader)
        for field_name, field_value in variables.items():
            with contextlib.suppress(Exception):
                writer.update_page_form_field_values(
                    writer.pages[0], {field_name: str(field_value)}
                )
        _save_pdf(writer, validated_out)
        return {"output": validated_out, "template": template_name, "mode": "acroform"}

    # reportlab 布局模板：读取模板文本内容并替换占位符后生成 PDF
    text = template_manager._read_template_text(tpl_path, "pdf")
    if variables:
        import re

        text = re.sub(
            r"\{\{([^}]+)\}\}",
            lambda m: str(variables.get(m.group(1), "")),
            text,
        )

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    y = A4[1] - 72
    for line in text.split("\n"):
        if y < 72:
            can.showPage()
            y = A4[1] - 72
        can.drawString(72, y, line)
        y -= 20
    can.save()
    packet.seek(0)

    reader2 = PdfReader(packet)
    writer = PdfWriter()
    for page in reader2.pages:
        writer.add_page(page)
    _save_pdf(writer, validated_out)
    return {"output": validated_out, "template": template_name, "mode": "reportlab"}


def pdf_fill_form(
    filename: str,
    fields: dict[str, str],
    flatten: bool = True,
    output: str | None = None,
) -> dict[str, Any]:
    """填充 PDF 交互式表单字段（AcroForm）。"""
    validated = _validate_pdf(filename)
    reader = PdfReader(validated)
    if not reader.get_fields():
        raise ToolError("该 PDF 不包含交互式表单字段（AcroForm）")

    writer = PdfWriter()
    writer.append(reader)
    filled = 0
    for page in writer.pages:
        try:
            writer.update_page_form_field_values(page, fields)
            filled += 1
        except Exception:
            pass

    if flatten:
        for page in writer.pages:
            # 扁平化：将表单字段转为静态内容
            writer.flatten_page(page)

    out = output or validated
    _save_pdf(writer, out)
    return {"filename": out, "fields_filled": filled, "flattened": flatten}


# ==================== 辅助函数 ====================


def _resolve_range(page_range: str | None, total: int) -> tuple[int, int]:
    """解析页码范围，返回 (start, end)（0-based, end exclusive）。"""
    if page_range is None:
        return 0, total
    parts = page_range.split("-")
    if len(parts) == 1:
        page = int(parts[0])
        return page - 1, page
    start = int(parts[0])
    end = int(parts[1])
    return start - 1, end
