"""MCP Server 主入口 - timeverse-office-doc-mcp。

使用 mcp SDK 原生 Server 类，stdio 传输。
Phase 3 注册 Word(21) + Excel(24) + PPT(19) + PDF(19) + Doc(10) 共 93 个工具。
"""

from __future__ import annotations

import asyncio
import json
import logging
import sys
from typing import Any

from mcp.server import Server
from mcp.types import TextContent, Tool

from .common.audit_log import audit_logger
from .common.error_handler import ToolError
from .config import ensure_dirs
from .handlers import doc_handler, excel_handler, pdf_handler, ppt_handler, word_handler

logger = logging.getLogger("timeverse_office_doc_mcp")

# ==================== MCP Server ====================

server = Server("timeverse-office-doc-mcp")


# ==================== 工具定义 ====================


def _str_param(desc: str, **extra: Any) -> dict[str, Any]:
    """快捷构造 string 参数 schema。"""
    schema = {"type": "string", "description": desc}
    schema.update(extra)
    return schema


def _int_param(desc: str, **extra: Any) -> dict[str, Any]:
    """快捷构造 integer 参数 schema。"""
    schema = {"type": "integer", "description": desc}
    schema.update(extra)
    return schema


def _bool_param(desc: str, **extra: Any) -> dict[str, Any]:
    """快捷构造 boolean 参数 schema。"""
    schema = {"type": "boolean", "description": desc}
    schema.update(extra)
    return schema


def _array_param(desc: str, items: dict[str, Any], **extra: Any) -> dict[str, Any]:
    """快捷构造 array 参数 schema。"""
    schema = {"type": "array", "description": desc, "items": items}
    schema.update(extra)
    return schema


def _obj_param(desc: str, props: dict[str, Any], **extra: Any) -> dict[str, Any]:
    """快捷构造 object 参数 schema。"""
    schema = {"type": "object", "description": desc, "properties": props}
    schema.update(extra)
    return schema


SESSION_ID_PARAM = _str_param("Session ID（传入则在内存中操作，不传则读写磁盘）")

TOOL_DEFINITIONS: list[Tool] = [
    # ==================== 5.1.1 文档管理 ====================
    Tool(
        name="word_create_document",
        description="创建新 Word 文档（支持模板创建与变量填充）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径（.docx）"),
                "title": _str_param("文档标题", default=""),
                "author": _str_param("作者", default=""),
                "template": _str_param("模板名称（可选，从模板库选择）"),
                "variables": _obj_param("模板变量（键值对，用于填充 {{placeholder}}）", {}),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="word_get_info",
        description="获取 Word 文档元信息（段落数、表格数、页数等）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="word_get_text",
        description="提取 Word 全文文本（可选包含表格内容）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "include_tables": _bool_param("是否包含表格文本", default=True),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="word_get_outline",
        description="获取 Word 文档大纲结构（标题层级）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="word_list_documents",
        description="列出指定目录内的所有 Word 文档",
        inputSchema={
            "type": "object",
            "properties": {"directory": _str_param("目录路径")},
            "required": ["directory"],
        },
    ),
    Tool(
        name="word_copy_document",
        description="复制 Word 文档",
        inputSchema={
            "type": "object",
            "properties": {
                "source": _str_param("源文件路径"),
                "destination": _str_param("目标文件路径"),
            },
            "required": ["source", "destination"],
        },
    ),
    # ==================== 5.1.2 内容编辑 ====================
    Tool(
        name="word_add_heading",
        description="添加标题（Heading 样式）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "text": _str_param("标题文本"),
                "level": _int_param("标题级别（0=Title, 1-9=Heading 1-9）", default=1),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "text"],
        },
    ),
    Tool(
        name="word_add_paragraph",
        description="添加段落（支持样式、字号、加粗）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "text": _str_param("段落文本"),
                "style": _str_param("段落样式名（如 Normal, List Bullet 等）"),
                "font_size": _int_param("字号（磅）"),
                "bold": _bool_param("是否加粗", default=False),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "text"],
        },
    ),
    Tool(
        name="word_add_table",
        description="添加表格（支持数据填充与表头样式）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "rows": _int_param("行数"),
                "cols": _int_param("列数"),
                "data": _array_param(
                    "表格数据（二维数组，外层=行，内层=单元格）",
                    {"type": "array", "items": {"type": "string"}},
                ),
                "has_header": _bool_param("首行是否为表头（加粗）", default=True),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "rows", "cols"],
        },
    ),
    Tool(
        name="word_add_image",
        description="插入图片到文档",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "image_path": _str_param("图片文件路径"),
                "width": {"type": "number", "description": "图片宽度（英寸）"},
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "image_path"],
        },
    ),
    Tool(
        name="word_add_page_break",
        description="插入分页符",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="word_add_list",
        description="添加列表（项目符号或编号）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "items": _array_param("列表项", {"type": "string"}),
                "list_style": _str_param(
                    "列表样式（List Bullet / List Number）", default="List Bullet"
                ),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "items"],
        },
    ),
    Tool(
        name="word_set_header_footer",
        description="设置页眉页脚（可选页码）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "header_text": _str_param("页眉文本", default=""),
                "footer_text": _str_param("页脚文本", default=""),
                "include_page_num": _bool_param("页脚是否包含页码", default=False),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="word_generate_toc",
        description="生成目录（Table of Contents，需在 Word 中更新域）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "max_level": _int_param("最大标题级别（1-9）", default=3),
                "styles": _str_param("自定义样式（可选）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    # ==================== 5.1.3 格式化与操作 ====================
    Tool(
        name="word_format_text",
        description="格式化文本片段（加粗、斜体、颜色、字体）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "paragraph_idx": _int_param("段落索引（0-based）"),
                "start": _int_param("起始字符位置（0-based）"),
                "end": _int_param("结束字符位置"),
                "bold": _bool_param("加粗"),
                "italic": _bool_param("斜体"),
                "color": _str_param("颜色（十六进制，如 FF0000）"),
                "font": _str_param("字体名称"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "paragraph_idx", "start", "end"],
        },
    ),
    Tool(
        name="word_format_table",
        description="格式化表格（表头加粗、底纹）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "table_idx": _int_param("表格索引（0-based）"),
                "border_style": _str_param("边框样式", default="single"),
                "header_row": _bool_param("是否格式化表头行", default=True),
                "shading": _str_param("底纹颜色（十六进制，如 D9E2F3）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "table_idx"],
        },
    ),
    Tool(
        name="word_search_replace",
        description="搜索替换文档中的文本",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "find_text": _str_param("查找文本"),
                "replace_text": _str_param("替换文本"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "find_text", "replace_text"],
        },
    ),
    Tool(
        name="word_delete_paragraph",
        description="删除指定段落",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "paragraph_idx": _int_param("段落索引（0-based）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "paragraph_idx"],
        },
    ),
    Tool(
        name="word_create_style",
        description="创建自定义段落样式",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "style_name": _str_param("样式名称"),
                "font": _str_param("字体名称", default="Calibri"),
                "size": _int_param("字号（磅）", default=11),
                "color": _str_param("颜色（十六进制）", default="000000"),
                "bold": _bool_param("是否加粗", default=False),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "style_name"],
        },
    ),
    # ==================== 5.1.4 分析工具 ====================
    Tool(
        name="word_analyze_structure",
        description="分析文档结构（标题层级、段落分布、表格统计、图片统计）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="word_extract_tables",
        description="提取所有表格数据（JSON 或 CSV 格式）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "format": _str_param("输出格式", default="json"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    # ==================== 5.2.1 工作簿管理 ====================
    Tool(
        name="excel_create_workbook",
        description="创建新 Excel 工作簿（支持模板）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径（.xlsx）"),
                "sheet_name": _str_param("默认工作表名", default="Sheet"),
                "template": _str_param("模板名称（可选，从模板库选择）"),
                "variables": _obj_param("模板变量（键值对，用于填充 {{placeholder}}）", {}),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="excel_get_info",
        description="获取 Excel 工作簿元信息（工作表列表、尺寸等）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="excel_list_sheets",
        description="列出所有工作表名称",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="excel_add_sheet",
        description="添加新工作表",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet_name": _str_param("新工作表名"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet_name"],
        },
    ),
    Tool(
        name="excel_delete_sheet",
        description="删除工作表",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet_name": _str_param("待删除工作表名"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet_name"],
        },
    ),
    Tool(
        name="excel_rename_sheet",
        description="重命名工作表",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "old_name": _str_param("原工作表名"),
                "new_name": _str_param("新工作表名"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "old_name", "new_name"],
        },
    ),
    Tool(
        name="excel_copy_sheet",
        description="复制工作表",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "source": _str_param("源工作表名"),
                "target": _str_param("目标工作表名"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "source", "target"],
        },
    ),
    # ==================== 5.2.2 数据读写 ====================
    Tool(
        name="excel_read_cell",
        description="读取单元格值",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "cell_ref": _str_param("单元格引用（如 A1）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet", "cell_ref"],
        },
    ),
    Tool(
        name="excel_write_cell",
        description="写入单元格值",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "cell_ref": _str_param("单元格引用（如 A1）"),
                "value": {"description": "单元格值（字符串、数字或布尔）"},
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet", "cell_ref", "value"],
        },
    ),
    Tool(
        name="excel_read_range",
        description="读取区域数据（二维数组）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "range_str": _str_param("区域范围（如 A1:C10）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet", "range_str"],
        },
    ),
    Tool(
        name="excel_write_range",
        description="批量写入区域数据",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "start_cell": _str_param("起始单元格（如 A1）"),
                "data": _array_param(
                    "写入数据（二维数组，外层=行，内层=单元格值）",
                    {"type": "array", "items": {}},
                ),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet", "start_cell", "data"],
        },
    ),
    Tool(
        name="excel_insert_row",
        description="插入行",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "row_idx": _int_param("行索引（1-based）"),
                "count": _int_param("插入行数", default=1),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet", "row_idx"],
        },
    ),
    Tool(
        name="excel_delete_row",
        description="删除行",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "row_idx": _int_param("行索引（1-based）"),
                "count": _int_param("删除行数", default=1),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet", "row_idx"],
        },
    ),
    Tool(
        name="excel_insert_column",
        description="插入列",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "col_idx": _int_param("列索引（1-based）"),
                "count": _int_param("插入列数", default=1),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet", "col_idx"],
        },
    ),
    Tool(
        name="excel_delete_column",
        description="删除列",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "col_idx": _int_param("列索引（1-based）"),
                "count": _int_param("删除列数", default=1),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet", "col_idx"],
        },
    ),
    # ==================== 5.2.3 格式化与高级 ====================
    Tool(
        name="excel_format_cell",
        description="格式化单元格（字体、颜色、对齐、边框）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "range_str": _str_param("区域范围（如 A1:C10）"),
                "font": _str_param("字体名称"),
                "bold": _bool_param("加粗", default=False),
                "italic": _bool_param("斜体", default=False),
                "font_size": _int_param("字号"),
                "font_color": _str_param("字体颜色（十六进制，如 FF0000）"),
                "bg_color": _str_param("背景色（十六进制，如 D9E2F3）"),
                "alignment": _str_param("对齐方式（left/center/right）"),
                "border_style": _str_param("边框样式（如 single/thin/medium）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet", "range_str"],
        },
    ),
    Tool(
        name="excel_apply_formula",
        description="应用公式到单元格",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "cell_ref": _str_param("单元格引用（如 A1）"),
                "formula": _str_param("公式（如 SUM(B1:B10)）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet", "cell_ref", "formula"],
        },
    ),
    Tool(
        name="excel_create_chart",
        description="创建图表（bar/line/pie）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "chart_type": _str_param("图表类型（bar/line/pie）"),
                "data_range": _str_param("数据区域（如 A1:B10）"),
                "title": _str_param("图表标题", default=""),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet", "chart_type", "data_range"],
        },
    ),
    Tool(
        name="excel_freeze_panes",
        description="冻结窗格",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "cell_ref": _str_param("冻结位置单元格（如 B2）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet", "cell_ref"],
        },
    ),
    Tool(
        name="excel_sort_data",
        description="排序数据区域",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "range_str": _str_param("数据区域（如 A1:C10）"),
                "key_column": _int_param("排序依据列（1-based）"),
                "ascending": _bool_param("是否升序", default=True),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet", "range_str", "key_column"],
        },
    ),
    Tool(
        name="excel_create_pivot_table",
        description="创建数据透视表",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "source_sheet": _str_param("源工作表名"),
                "source_range": _str_param("源数据区域（如 A1:D100）"),
                "target_sheet": _str_param("目标工作表名"),
                "rows": _str_param("行字段（列名）"),
                "cols": _str_param("列字段（列名）"),
                "values": _str_param("值字段（列名）"),
                "agg_func": _str_param("聚合函数（sum/mean/count/max/min）", default="sum"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": [
                "filename",
                "source_sheet",
                "source_range",
                "target_sheet",
                "rows",
                "cols",
                "values",
            ],
        },
    ),
    Tool(
        name="excel_add_conditional_format",
        description="添加条件格式",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "range_str": _str_param("区域范围（如 A1:C10）"),
                "rule_type": _str_param(
                    "规则类型（greater_than/less_than/equal/between/contains_text）"
                ),
                "criteria": _str_param("判断条件（between 用 'min,max'）"),
                "format_color": _str_param("格式颜色（十六进制，如 FF0000）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet", "range_str", "rule_type"],
        },
    ),
    # ==================== 5.2.4 分析工具 ====================
    Tool(
        name="excel_analyze_data",
        description="数据统计分析（描述统计、空值检测、类型推断）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "range_str": _str_param("数据区域（可选，默认全部）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet"],
        },
    ),
    Tool(
        name="excel_find_duplicates",
        description="查找重复数据",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("工作簿文件路径"),
                "sheet": _str_param("工作表名"),
                "columns": _array_param("检查列名列表（可选，默认全部）", {"type": "string"}),
                "threshold": _int_param("重复阈值", default=1),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "sheet"],
        },
    ),
    # ==================== 5.3.1 演示文稿管理 ====================
    Tool(
        name="ppt_create_presentation",
        description="创建新 PPT 演示文稿（支持模板）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径（.pptx）"),
                "template": _str_param("模板名称（可选，从模板库选择）"),
                "variables": _obj_param("模板变量（键值对，用于填充 {{placeholder}}）", {}),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="ppt_get_info",
        description="获取 PPT 演示文稿元信息（幻灯片数、尺寸等）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="ppt_list_slides",
        description="列出所有幻灯片",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="ppt_add_slide",
        description="添加幻灯片（layout: 0-8 对应 slide_layouts）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "layout": _int_param("版式索引（0-8）", default=6),
                "title": _str_param("幻灯片标题（可选）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="ppt_delete_slide",
        description="删除指定幻灯片",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "slide_idx": _int_param("幻灯片索引（0-based）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "slide_idx"],
        },
    ),
    Tool(
        name="ppt_move_slide",
        description="移动幻灯片到新位置",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "slide_idx": _int_param("原幻灯片索引（0-based）"),
                "new_idx": _int_param("目标位置索引（0-based）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "slide_idx", "new_idx"],
        },
    ),
    Tool(
        name="ppt_copy_slide",
        description="复制幻灯片（副本插入到原幻灯片之后）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "slide_idx": _int_param("幻灯片索引（0-based）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "slide_idx"],
        },
    ),
    # ==================== 5.3.2 内容编辑 ====================
    Tool(
        name="ppt_add_text",
        description="添加文本框（单位为英寸）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "slide_idx": _int_param("幻灯片索引（0-based）"),
                "text": _str_param("文本内容"),
                "left": {"type": "number", "description": "左侧位置（英寸）", "default": 1.0},
                "top": {"type": "number", "description": "顶部位置（英寸）", "default": 1.0},
                "width": {"type": "number", "description": "宽度（英寸）", "default": 8.0},
                "height": {"type": "number", "description": "高度（英寸）", "default": 1.5},
                "font_size": _int_param("字号（磅）", default=18),
                "bold": _bool_param("是否加粗", default=False),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "slide_idx", "text"],
        },
    ),
    Tool(
        name="ppt_add_image",
        description="插入图片到幻灯片（单位为英寸）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "slide_idx": _int_param("幻灯片索引（0-based）"),
                "image_path": _str_param("图片文件路径"),
                "left": {"type": "number", "description": "左侧位置（英寸）", "default": 1.0},
                "top": {"type": "number", "description": "顶部位置（英寸）", "default": 1.0},
                "width": {"type": "number", "description": "宽度（英寸）"},
                "height": {"type": "number", "description": "高度（英寸）"},
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "slide_idx", "image_path"],
        },
    ),
    Tool(
        name="ppt_add_table",
        description="添加表格到幻灯片（单位为英寸）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "slide_idx": _int_param("幻灯片索引（0-based）"),
                "rows": _int_param("行数"),
                "cols": _int_param("列数"),
                "data": _array_param(
                    "表格数据（二维数组，外层=行，内层=单元格）",
                    {"type": "array", "items": {"type": "string"}},
                ),
                "left": {"type": "number", "description": "左侧位置（英寸）", "default": 1.0},
                "top": {"type": "number", "description": "顶部位置（英寸）", "default": 2.0},
                "width": {"type": "number", "description": "宽度（英寸）", "default": 8.0},
                "height": {"type": "number", "description": "高度（英寸）", "default": 3.0},
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "slide_idx", "rows", "cols"],
        },
    ),
    Tool(
        name="ppt_add_chart",
        description="添加图表（bar/line/pie；data 含 categories 与 series）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "slide_idx": _int_param("幻灯片索引（0-based）"),
                "chart_type": _str_param("图表类型（bar/line/pie）"),
                "data": _obj_param("图表数据（含 categories 与 series）", {}),
                "title": _str_param("图表标题", default=""),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "slide_idx", "chart_type", "data"],
        },
    ),
    Tool(
        name="ppt_add_shape",
        description="添加形状（rectangle/rounded_rectangle/oval/triangle/textbox）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "slide_idx": _int_param("幻灯片索引（0-based）"),
                "shape_type": _str_param(
                    "形状类型（rectangle/rounded_rectangle/oval/triangle/textbox）"
                ),
                "left": {"type": "number", "description": "左侧位置（英寸）", "default": 1.0},
                "top": {"type": "number", "description": "顶部位置（英寸）", "default": 1.0},
                "width": {"type": "number", "description": "宽度（英寸）", "default": 3.0},
                "height": {"type": "number", "description": "高度（英寸）", "default": 2.0},
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "slide_idx", "shape_type"],
        },
    ),
    Tool(
        name="ppt_set_background",
        description="设置幻灯片背景色",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "slide_idx": _int_param("幻灯片索引（0-based）"),
                "color": _str_param("背景色（十六进制，如 FF0000）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "slide_idx"],
        },
    ),
    Tool(
        name="ppt_apply_theme",
        description="应用主题配色（blue/green/orange/dark）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "theme_name": _str_param("主题名（blue/green/orange/dark）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "theme_name"],
        },
    ),
    Tool(
        name="ppt_set_slide_notes",
        description="设置演讲者备注",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "slide_idx": _int_param("幻灯片索引（0-based）"),
                "notes_text": _str_param("备注文本"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "slide_idx", "notes_text"],
        },
    ),
    # ==================== 5.3.3 分析工具 ====================
    Tool(
        name="ppt_extract_text",
        description="提取所有幻灯片文本",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="ppt_get_slide_notes",
        description="获取指定幻灯片的演讲者备注",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "slide_idx": _int_param("幻灯片索引（0-based）"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename", "slide_idx"],
        },
    ),
    Tool(
        name="ppt_analyze_structure",
        description="分析演示文稿结构（幻灯片分布、形状统计等）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="ppt_get_structure",
        description="获取演示文稿结构（幻灯片与形状详情）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("演示文稿文件路径"),
                "session_id": SESSION_ID_PARAM,
            },
            "required": ["filename"],
        },
    ),
    # ==================== 5.4.1 文档管理 ====================
    Tool(
        name="pdf_get_info",
        description="获取 PDF 元信息（页数、作者、标题等）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径（.pdf）"),
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="pdf_merge",
        description="合并多个 PDF 文件",
        inputSchema={
            "type": "object",
            "properties": {
                "files": _array_param("待合并的 PDF 文件路径列表", {"type": "string"}),
                "output": _str_param("合并后输出文件路径"),
            },
            "required": ["files", "output"],
        },
    ),
    Tool(
        name="pdf_split",
        description="拆分 PDF（page_ranges 格式: 1-3,4-6,7-9）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "page_ranges": _str_param("页码范围（如 1-3,4-6,7-9）"),
                "output_prefix": _str_param("输出文件前缀"),
            },
            "required": ["filename", "page_ranges", "output_prefix"],
        },
    ),
    Tool(
        name="pdf_rotate_page",
        description="旋转指定页面（angle: 90/180/270）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "page_idx": _int_param("页面索引（0-based）"),
                "angle": _int_param("旋转角度（90/180/270）"),
            },
            "required": ["filename", "page_idx", "angle"],
        },
    ),
    # ==================== 5.4.2 内容读取 ====================
    Tool(
        name="pdf_extract_text",
        description="提取 PDF 文本（支持指定页范围与布局模式）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "page_range": _str_param("页码范围（如 1-3，可选）"),
                "layout_mode": _bool_param("是否保留布局", default=False),
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="pdf_extract_tables",
        description="提取 PDF 表格数据（JSON 或 CSV）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "page_range": _str_param("页码范围（如 1-3，可选）"),
                "format": _str_param("输出格式（json/csv）", default="json"),
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="pdf_extract_images",
        description="提取 PDF 中的图片",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "page_range": _str_param("页码范围（如 1-3，可选）"),
                "output_dir": _str_param("图片输出目录（可选）"),
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="pdf_search_text",
        description="搜索 PDF 中的文本",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "query": _str_param("搜索关键词"),
                "case_sensitive": _bool_param("是否区分大小写", default=False),
            },
            "required": ["filename", "query"],
        },
    ),
    Tool(
        name="pdf_ocr_text",
        description="OCR 文本识别（扫描件/图片型 PDF，需安装 tesseract）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "page_range": _str_param("页码范围（如 1-3，可选）"),
                "lang": _str_param("OCR 语言（如 chi_sim+eng）", default="chi_sim+eng"),
                "output_format": _str_param("输出格式（text）", default="text"),
            },
            "required": ["filename"],
        },
    ),
    # ==================== 5.4.3 内容写入 ====================
    Tool(
        name="pdf_add_text",
        description="添加文本到指定页面（Overlay 合并模式）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "page_idx": _int_param("页面索引（0-based）"),
                "text": _str_param("文本内容"),
                "x": {"type": "number", "description": "X 坐标（磅）", "default": 72},
                "y": {"type": "number", "description": "Y 坐标（磅）", "default": 72},
                "font": _str_param("字体名称", default="Helvetica"),
                "font_size": _int_param("字号（磅）", default=12),
                "output": _str_param("输出文件路径（可选，默认覆盖原文件）"),
            },
            "required": ["filename", "page_idx", "text"],
        },
    ),
    Tool(
        name="pdf_add_image",
        description="添加图片到指定页面（Overlay 合并模式）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "page_idx": _int_param("页面索引（0-based）"),
                "image_path": _str_param("图片文件路径"),
                "x": {"type": "number", "description": "X 坐标（磅）", "default": 72},
                "y": {"type": "number", "description": "Y 坐标（磅）", "default": 72},
                "width": {"type": "number", "description": "图片宽度（磅）", "default": 200},
                "height": {"type": "number", "description": "图片高度（磅）", "default": 150},
                "output": _str_param("输出文件路径（可选，默认覆盖原文件）"),
            },
            "required": ["filename", "page_idx", "image_path"],
        },
    ),
    Tool(
        name="pdf_add_watermark",
        description="添加水印（遍历所有页面）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "watermark_text": _str_param("水印文本"),
                "opacity": {"type": "number", "description": "透明度（0-1）", "default": 0.3},
                "font_size": _int_param("字号（磅）", default=60),
                "output": _str_param("输出文件路径（可选，默认覆盖原文件）"),
            },
            "required": ["filename", "watermark_text"],
        },
    ),
    Tool(
        name="pdf_add_annotation",
        description="添加注释（highlight/text/link）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "page_idx": _int_param("页面索引（0-based）"),
                "annotation_type": _str_param("注释类型（highlight/text/link）"),
                "content": _str_param("注释内容"),
                "x": {"type": "number", "description": "X 坐标（磅）", "default": 72},
                "y": {"type": "number", "description": "Y 坐标（磅）", "default": 72},
                "output": _str_param("输出文件路径（可选，默认覆盖原文件）"),
            },
            "required": ["filename", "page_idx", "annotation_type", "content"],
        },
    ),
    Tool(
        name="pdf_add_bookmark",
        description="添加书签",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "title": _str_param("书签标题"),
                "page_idx": _int_param("目标页面索引（0-based）"),
                "output": _str_param("输出文件路径（可选，默认覆盖原文件）"),
            },
            "required": ["filename", "title", "page_idx"],
        },
    ),
    # ==================== 5.4.4 安全与分析 ====================
    Tool(
        name="pdf_encrypt",
        description="加密 PDF",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "password": _str_param("加密密码"),
                "permissions": _array_param("权限列表（可选）", {"type": "string"}),
                "output": _str_param("输出文件路径（可选，默认覆盖原文件）"),
            },
            "required": ["filename", "password"],
        },
    ),
    Tool(
        name="pdf_decrypt",
        description="解密 PDF",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "password": _str_param("解密密码"),
                "output": _str_param("输出文件路径（可选，默认覆盖原文件）"),
            },
            "required": ["filename", "password"],
        },
    ),
    Tool(
        name="pdf_analyze_structure",
        description="分析 PDF 结构（页面类型、文本密度、表格分布）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
            },
            "required": ["filename"],
        },
    ),
    # ==================== 5.4.5 模板工具 ====================
    Tool(
        name="pdf_create_from_template",
        description="从模板创建 PDF（支持变量填充）",
        inputSchema={
            "type": "object",
            "properties": {
                "template_name": _str_param("模板名称"),
                "variables": _obj_param("模板变量（键值对，用于填充 {{placeholder}}）", {}),
                "output": _str_param("输出文件路径", default="output.pdf"),
            },
            "required": ["template_name"],
        },
    ),
    Tool(
        name="pdf_fill_form",
        description="填充 PDF 交互式表单字段（AcroForm）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("PDF 文件路径"),
                "fields": _obj_param("表单字段（字段名到值的映射）", {}),
                "flatten": _bool_param("是否扁平化表单字段", default=True),
                "output": _str_param("输出文件路径（可选，默认覆盖原文件）"),
            },
            "required": ["filename", "fields"],
        },
    ),
    # ==================== 5.5.1 模板管理 ====================
    Tool(
        name="doc_list_templates",
        description="列出所有可用模板（可选按格式过滤）",
        inputSchema={
            "type": "object",
            "properties": {
                "format": _str_param("按格式过滤（word/excel/ppt/pdf，可选）"),
            },
            "required": [],
        },
    ),
    Tool(
        name="doc_get_template_info",
        description="获取模板详情（含占位符列表与预览）",
        inputSchema={
            "type": "object",
            "properties": {
                "template_name": _str_param("模板名称"),
            },
            "required": ["template_name"],
        },
    ),
    Tool(
        name="doc_register_template",
        description="注册新模板到模板库",
        inputSchema={
            "type": "object",
            "properties": {
                "name": _str_param("模板名称"),
                "format": _str_param("模板格式（word/excel/ppt/pdf）"),
                "file_path": _str_param("模板文件路径"),
                "description": _str_param("模板描述", default=""),
                "placeholders": _array_param("占位符列表（可选）", {"type": "string"}),
            },
            "required": ["name", "format", "file_path"],
        },
    ),
    Tool(
        name="doc_delete_template",
        description="从模板库删除模板",
        inputSchema={
            "type": "object",
            "properties": {
                "template_name": _str_param("模板名称"),
            },
            "required": ["template_name"],
        },
    ),
    # ==================== 5.5.2 模板应用 ====================
    Tool(
        name="doc_apply_template",
        description="从模板创建文档并自动填充变量（支持所有格式）",
        inputSchema={
            "type": "object",
            "properties": {
                "template_name": _str_param("模板名称"),
                "output_path": _str_param("输出文件路径"),
                "variables": _obj_param("模板变量（键值对，用于填充 {{placeholder}}）", {}),
            },
            "required": ["template_name", "output_path"],
        },
    ),
    Tool(
        name="doc_extract_placeholders",
        description="扫描提取模板中的占位符变量",
        inputSchema={
            "type": "object",
            "properties": {
                "template_name": _str_param("模板名称"),
                "format": _str_param("模板格式（可选，自动推断）"),
            },
            "required": ["template_name"],
        },
    ),
    # ==================== 5.6 Session 管理 ====================
    Tool(
        name="doc_open_session",
        description="打开文档到内存 Session（返回 session_id）",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": _str_param("文档文件路径"),
                "format": _str_param("文档格式（word/excel/ppt/pdf）"),
            },
            "required": ["filename", "format"],
        },
    ),
    Tool(
        name="doc_save_session",
        description="保存 Session 到磁盘（不指定路径则保存回原路径）",
        inputSchema={
            "type": "object",
            "properties": {
                "session_id": _str_param("待保存的 Session ID"),
                "output_path": _str_param("输出文件路径（可选，默认原路径）"),
            },
            "required": ["session_id"],
        },
    ),
    Tool(
        name="doc_close_session",
        description="关闭 Session（save=True 时先保存再关闭）",
        inputSchema={
            "type": "object",
            "properties": {
                "session_id": _str_param("待关闭的 Session ID"),
                "save": _bool_param("是否先保存再关闭", default=False),
            },
            "required": ["session_id"],
        },
    ),
    Tool(
        name="doc_list_sessions",
        description="列出所有活跃 Session",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
]


# ==================== 工具路由表 ====================

TOOL_HANDLERS: dict[str, Any] = {
    "word_create_document": word_handler.word_create_document,
    "word_get_info": word_handler.word_get_info,
    "word_get_text": word_handler.word_get_text,
    "word_get_outline": word_handler.word_get_outline,
    "word_list_documents": word_handler.word_list_documents,
    "word_copy_document": word_handler.word_copy_document,
    "word_add_heading": word_handler.word_add_heading,
    "word_add_paragraph": word_handler.word_add_paragraph,
    "word_add_table": word_handler.word_add_table,
    "word_add_image": word_handler.word_add_image,
    "word_add_page_break": word_handler.word_add_page_break,
    "word_add_list": word_handler.word_add_list,
    "word_set_header_footer": word_handler.word_set_header_footer,
    "word_generate_toc": word_handler.word_generate_toc,
    "word_format_text": word_handler.word_format_text,
    "word_format_table": word_handler.word_format_table,
    "word_search_replace": word_handler.word_search_replace,
    "word_delete_paragraph": word_handler.word_delete_paragraph,
    "word_create_style": word_handler.word_create_style,
    "word_analyze_structure": word_handler.word_analyze_structure,
    "word_extract_tables": word_handler.word_extract_tables,
    # ==================== Excel ====================
    "excel_create_workbook": excel_handler.excel_create_workbook,
    "excel_get_info": excel_handler.excel_get_info,
    "excel_list_sheets": excel_handler.excel_list_sheets,
    "excel_add_sheet": excel_handler.excel_add_sheet,
    "excel_delete_sheet": excel_handler.excel_delete_sheet,
    "excel_rename_sheet": excel_handler.excel_rename_sheet,
    "excel_copy_sheet": excel_handler.excel_copy_sheet,
    "excel_read_cell": excel_handler.excel_read_cell,
    "excel_write_cell": excel_handler.excel_write_cell,
    "excel_read_range": excel_handler.excel_read_range,
    "excel_write_range": excel_handler.excel_write_range,
    "excel_insert_row": excel_handler.excel_insert_row,
    "excel_delete_row": excel_handler.excel_delete_row,
    "excel_insert_column": excel_handler.excel_insert_column,
    "excel_delete_column": excel_handler.excel_delete_column,
    "excel_format_cell": excel_handler.excel_format_cell,
    "excel_apply_formula": excel_handler.excel_apply_formula,
    "excel_create_chart": excel_handler.excel_create_chart,
    "excel_freeze_panes": excel_handler.excel_freeze_panes,
    "excel_sort_data": excel_handler.excel_sort_data,
    "excel_create_pivot_table": excel_handler.excel_create_pivot_table,
    "excel_add_conditional_format": excel_handler.excel_add_conditional_format,
    "excel_analyze_data": excel_handler.excel_analyze_data,
    "excel_find_duplicates": excel_handler.excel_find_duplicates,
    # ==================== PPT ====================
    "ppt_create_presentation": ppt_handler.ppt_create_presentation,
    "ppt_get_info": ppt_handler.ppt_get_info,
    "ppt_list_slides": ppt_handler.ppt_list_slides,
    "ppt_add_slide": ppt_handler.ppt_add_slide,
    "ppt_delete_slide": ppt_handler.ppt_delete_slide,
    "ppt_move_slide": ppt_handler.ppt_move_slide,
    "ppt_copy_slide": ppt_handler.ppt_copy_slide,
    "ppt_add_text": ppt_handler.ppt_add_text,
    "ppt_add_image": ppt_handler.ppt_add_image,
    "ppt_add_table": ppt_handler.ppt_add_table,
    "ppt_add_chart": ppt_handler.ppt_add_chart,
    "ppt_add_shape": ppt_handler.ppt_add_shape,
    "ppt_set_background": ppt_handler.ppt_set_background,
    "ppt_apply_theme": ppt_handler.ppt_apply_theme,
    "ppt_set_slide_notes": ppt_handler.ppt_set_slide_notes,
    "ppt_extract_text": ppt_handler.ppt_extract_text,
    "ppt_get_slide_notes": ppt_handler.ppt_get_slide_notes,
    "ppt_analyze_structure": ppt_handler.ppt_analyze_structure,
    "ppt_get_structure": ppt_handler.ppt_get_structure,
    # ==================== PDF ====================
    "pdf_get_info": pdf_handler.pdf_get_info,
    "pdf_merge": pdf_handler.pdf_merge,
    "pdf_split": pdf_handler.pdf_split,
    "pdf_rotate_page": pdf_handler.pdf_rotate_page,
    "pdf_extract_text": pdf_handler.pdf_extract_text,
    "pdf_extract_tables": pdf_handler.pdf_extract_tables,
    "pdf_extract_images": pdf_handler.pdf_extract_images,
    "pdf_search_text": pdf_handler.pdf_search_text,
    "pdf_ocr_text": pdf_handler.pdf_ocr_text,
    "pdf_add_text": pdf_handler.pdf_add_text,
    "pdf_add_image": pdf_handler.pdf_add_image,
    "pdf_add_watermark": pdf_handler.pdf_add_watermark,
    "pdf_add_annotation": pdf_handler.pdf_add_annotation,
    "pdf_add_bookmark": pdf_handler.pdf_add_bookmark,
    "pdf_encrypt": pdf_handler.pdf_encrypt,
    "pdf_decrypt": pdf_handler.pdf_decrypt,
    "pdf_analyze_structure": pdf_handler.pdf_analyze_structure,
    "pdf_create_from_template": pdf_handler.pdf_create_from_template,
    "pdf_fill_form": pdf_handler.pdf_fill_form,
    # ==================== Doc ====================
    "doc_list_templates": doc_handler.doc_list_templates,
    "doc_get_template_info": doc_handler.doc_get_template_info,
    "doc_register_template": doc_handler.doc_register_template,
    "doc_delete_template": doc_handler.doc_delete_template,
    "doc_apply_template": doc_handler.doc_apply_template,
    "doc_extract_placeholders": doc_handler.doc_extract_placeholders,
    "doc_open_session": doc_handler.doc_open_session,
    "doc_save_session": doc_handler.doc_save_session,
    "doc_close_session": doc_handler.doc_close_session,
    "doc_list_sessions": doc_handler.doc_list_sessions,
}


# ==================== MCP 请求处理 ====================


@server.list_tools()  # type: ignore[untyped-decorator]
async def list_tools() -> list[Tool]:
    """返回所有可用工具。"""
    return TOOL_DEFINITIONS


@server.call_tool()  # type: ignore[untyped-decorator]
async def call_tool(
    name: str,
    arguments: dict[str, Any],
) -> list[TextContent]:
    """处理工具调用：路由到对应 handler 并返回结果。"""
    import time

    handler = TOOL_HANDLERS.get(name)
    if handler is None:
        return [TextContent(type="text", text=f"Error: 未知工具 '{name}'")]

    start = time.time()
    success = True
    error_msg: str | None = None

    try:
        result = handler(**arguments)
        if result is None:
            text = "操作完成"
        elif isinstance(result, dict):
            text = json.dumps(result, ensure_ascii=False, default=str)
        else:
            text = str(result)
    except ToolError as e:
        success = False
        error_msg = str(e)
        text = f"Error: {e}"
    except Exception as e:
        success = False
        error_msg = str(e)
        logger.exception("Unhandled error in tool %s", name)
        text = f"Error: {e}"

    duration_ms = int((time.time() - start) * 1000)
    audit_logger.log_operation(
        tool_name=name,
        args=arguments,
        result=text[:200],
        duration_ms=duration_ms,
        success=success,
        error=error_msg,
    )

    return [TextContent(type="text", text=text)]


# ==================== 传输层 ====================


async def _run_stdio() -> None:
    """通过 stdio 传输启动。"""
    from mcp.server.stdio import stdio_server

    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )


def main() -> None:
    """CLI 入口点。"""
    logging.basicConfig(
        level=logging.WARNING,
        stream=sys.stderr,
        format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    )
    ensure_dirs()
    logger.info("Starting timeverse-office-doc-mcp (stdio)...")
    asyncio.run(_run_stdio())


if __name__ == "__main__":
    main()
