"""跨格式工具处理器 - 模板管理 6 个 + Session 管理 4 个 = 10 个工具。

对应方案 5.5 模板管理工具集 + 5.6 Session 管理工具集。
"""

from __future__ import annotations

import logging
from typing import Any

from ..common.error_handler import ToolError
from ..common.session import session_manager
from ..common.template_mgr import template_manager

logger = logging.getLogger("timeverse_office_doc_mcp.doc")


# ==================== 5.5.1 模板注册与管理（4 个） ====================


def doc_list_templates(format: str | None = None) -> dict[str, Any]:
    """列出所有可用模板，可选按格式过滤。"""
    templates = template_manager.list_templates(format)
    return {
        "count": len(templates),
        "templates": [
            {"name": t["name"], "format": t["format"], "description": t.get("description", "")}
            for t in templates
        ],
    }


def doc_get_template_info(template_name: str) -> dict[str, Any]:
    """获取模板详情（含占位符列表与预览）。"""
    info = template_manager.get_template_info(template_name)
    return {
        "name": info["name"],
        "format": info["format"],
        "description": info.get("description", ""),
        "placeholders": info.get("placeholders", []),
        "path": info["path"],
    }


def doc_register_template(
    name: str,
    format: str,
    file_path: str,
    description: str = "",
    placeholders: list[str] | None = None,
) -> dict[str, Any]:
    """注册新模板到模板库（将外部文件复制到 templates/{format}/ 并写入索引）。"""
    result = template_manager.register_template(
        name=name,
        format=format,
        file_path=file_path,
        description=description,
        placeholders=placeholders,
    )
    return {
        "name": result["name"],
        "format": result["format"],
        "description": result.get("description", ""),
        "placeholders": result.get("placeholders", []),
        "registered": True,
    }


def doc_delete_template(template_name: str) -> dict[str, Any]:
    """从模板库删除模板（删除索引记录与模板库副本，不影响用户原始文件）。"""
    template_manager.delete_template(template_name)
    return {"deleted": template_name}


# ==================== 5.5.2 模板应用与占位符（2 个） ====================


def doc_apply_template(
    template_name: str,
    output_path: str,
    variables: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """从模板创建文档并自动填充变量。

    核心工具：加载模板 -> 替换占位符 -> 保存输出。
    支持所有格式（Word/Excel/PPT/PDF）。
    """
    fmt, tpl_path = template_manager.resolve_template_path(template_name)
    variables = variables or {}

    if fmt == "word":
        return _apply_word_template(template_name, tpl_path, output_path, variables)
    if fmt == "excel":
        return _apply_excel_template(tpl_path, output_path, variables)
    if fmt == "ppt":
        return _apply_ppt_template(tpl_path, output_path, variables)
    if fmt == "pdf":
        from .pdf_handler import pdf_create_from_template

        result = pdf_create_from_template(template_name, variables, output_path)
        return {"template": template_name, "output": result["output"], "format": "pdf"}
    raise ToolError(f"不支持的模板格式: {fmt}")


def _apply_word_template(
    template_name: str, tpl_path: str, output_path: str, variables: dict[str, Any]
) -> dict[str, Any]:
    """应用 Word 模板。"""
    from docx import Document

    from ..common.file_lock import file_lock_mgr
    from ..common.path_guard import path_guard

    validated_out = path_guard.validate_path(output_path, "write")
    doc = Document(tpl_path)
    replaced = _fill_docx_variables(doc, variables)
    file_lock_mgr.acquire(validated_out)
    try:
        doc.save(validated_out)
    finally:
        file_lock_mgr.release(validated_out)
    return {
        "template": template_name,
        "output": validated_out,
        "format": "word",
        "variables_replaced": replaced,
    }


def _fill_docx_variables(doc: Any, variables: dict[str, Any]) -> int:
    """填充 Word 文档中的占位符。"""
    import re

    count = 0
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if "{{" in run.text:
                new_text = re.sub(
                    r"\{\{([^}]+)\}\}",
                    lambda m: str(variables.get(m.group(1), "")),
                    run.text,
                )
                if new_text != run.text:
                    run.text = new_text
                    count += 1
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if "{{" in run.text:
                            new_text = re.sub(
                                r"\{\{([^}]+)\}\}",
                                lambda m: str(variables.get(m.group(1), "")),
                                run.text,
                            )
                            if new_text != run.text:
                                run.text = new_text
                                count += 1
    return count


def _apply_excel_template(
    tpl_path: str, output_path: str, variables: dict[str, Any]
) -> dict[str, Any]:
    """应用 Excel 模板。"""
    import re

    from openpyxl import load_workbook

    from ..common.file_lock import file_lock_mgr
    from ..common.path_guard import path_guard

    validated_out = path_guard.validate_path(output_path, "write")
    wb = load_workbook(tpl_path)
    replaced = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "{{" in cell.value:
                    new_val = re.sub(
                        r"\{\{([^}]+)\}\}",
                        lambda m: str(variables.get(m.group(1), "")),
                        cell.value,
                    )
                    if new_val != cell.value:
                        cell.value = new_val
                        replaced += 1
    file_lock_mgr.acquire(validated_out)
    try:
        wb.save(validated_out)
    finally:
        file_lock_mgr.release(validated_out)
    return {
        "template": "excel",
        "output": validated_out,
        "format": "excel",
        "variables_replaced": replaced,
    }


def _apply_ppt_template(
    tpl_path: str, output_path: str, variables: dict[str, Any]
) -> dict[str, Any]:
    """应用 PPT 模板。"""
    import re

    from pptx import Presentation

    from ..common.file_lock import file_lock_mgr
    from ..common.path_guard import path_guard

    validated_out = path_guard.validate_path(output_path, "write")
    prs = Presentation(tpl_path)
    replaced = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text and "{{" in run.text:
                            new_text = re.sub(
                                r"\{\{([^}]+)\}\}",
                                lambda m: str(variables.get(m.group(1), "")),
                                run.text,
                            )
                            if new_text != run.text:
                                run.text = new_text
                                replaced += 1
    file_lock_mgr.acquire(validated_out)
    try:
        prs.save(validated_out)
    finally:
        file_lock_mgr.release(validated_out)
    return {
        "template": "ppt",
        "output": validated_out,
        "format": "ppt",
        "variables_replaced": replaced,
    }


def doc_extract_placeholders(template_name: str, format: str | None = None) -> dict[str, Any]:
    """扫描提取模板中的占位符变量。"""
    placeholders = template_manager.extract_placeholders(template_name)
    return {
        "template_name": template_name,
        "count": len(placeholders),
        "placeholders": placeholders,
    }


# ==================== 5.6 Session 管理工具集（4 个） ====================


def doc_open_session(filename: str, format: str) -> dict[str, Any]:
    """打开文档到内存 Session。

    返回 session_id，后续编辑工具可通过 session_id 操作内存中的文档。
    """
    from ..common.path_guard import path_guard

    validated = path_guard.validate_path(filename, "read")

    if format == "word":
        from docx import Document

        doc = Document(validated)
    elif format == "excel":
        from openpyxl import load_workbook

        doc = load_workbook(validated)
    elif format == "ppt":
        from pptx import Presentation

        doc = Presentation(validated)
    elif format == "pdf":
        from pypdf import PdfReader

        doc = PdfReader(validated)
    else:
        raise ToolError(f"不支持的格式: {format}")

    session_id = session_manager.open_session(validated, format, doc)
    return {
        "session_id": session_id,
        "filename": validated,
        "format": format,
    }


def doc_save_session(session_id: str, output_path: str | None = None) -> dict[str, Any]:
    """保存 Session 到磁盘。

    不指定 output_path 则保存回原路径。
    """
    from ..common.path_guard import path_guard

    session = session_manager.get_session(session_id)

    if output_path:
        validated_out = path_guard.validate_path(output_path, "write")
    else:
        validated_out = session.filename

    # 按格式保存
    doc = session.document
    if session.format == "word" or session.format == "excel" or session.format == "ppt":
        from ..common.file_lock import file_lock_mgr

        file_lock_mgr.acquire(validated_out)
        try:
            doc.save(validated_out)
        finally:
            file_lock_mgr.release(validated_out)
    elif session.format == "pdf":
        # PDF 的 Session 保存逻辑（如果有 PdfWriter 则写入）
        from ..common.file_lock import file_lock_mgr

        if hasattr(doc, "write"):
            file_lock_mgr.acquire(validated_out)
            try:
                with open(validated_out, "wb") as f:
                    doc.write(f)
            finally:
                file_lock_mgr.release(validated_out)

    session_manager.save_session(session_id, validated_out)
    return {"session_id": session_id, "saved_to": validated_out}


def doc_close_session(session_id: str, save: bool = False) -> dict[str, Any]:
    """关闭 Session。

    save=True 时先保存再关闭。
    """
    if save:
        doc_save_session(session_id)
    session_manager.close_session(session_id)
    return {"session_id": session_id, "closed": True, "saved": save}


def doc_list_sessions() -> dict[str, Any]:
    """列出所有活跃 Session。"""
    sessions = session_manager.list_sessions()
    return {
        "count": len(sessions),
        "sessions": sessions,
    }
