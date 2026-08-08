"""跨格式统一的模板占位符填充工具。

各 handler 共享的 {{variable}} 替换逻辑，消除重复代码。
"""

from __future__ import annotations

import re
from typing import Any

_PLACEHOLDER_RE = re.compile(r"\{\{([^}]+)\}\}")


def replace_placeholders(text: str, variables: dict[str, Any]) -> str:
    """替换文本中的 {{variable}} 占位符，未找到的变量替换为空字符串。"""

    def replacer(match: re.Match[str]) -> str:
        key = match.group(1).strip()
        value = variables.get(key)
        if value is None:
            return ""
        return str(value)

    return _PLACEHOLDER_RE.sub(replacer, text)


def fill_word_variables(doc: Any, variables: dict[str, Any]) -> int:
    """填充 Word 文档（python-docx Document）中的占位符，返回替换数量。"""
    count = 0
    for paragraph in doc.paragraphs:
        if "{{" in paragraph.text:
            count += _replace_in_runs(list(paragraph.runs), variables)
    # 表格单元格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if "{{" in paragraph.text:
                        count += _replace_in_runs(list(paragraph.runs), variables)
    return count


def fill_excel_variables(wb: Any, variables: dict[str, Any]) -> int:
    """填充 Excel 工作簿（openpyxl Workbook）中的占位符，返回替换数量。"""
    count = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "{{" in cell.value:
                    new_val = replace_placeholders(cell.value, variables)
                    if new_val != cell.value:
                        cell.value = new_val
                        count += 1
    return count


def fill_ppt_variables(prs: Any, variables: dict[str, Any]) -> int:
    """填充 PPT 演示文稿（python-pptx Presentation）中的占位符，返回替换数量。"""
    count = 0
    for slide in prs.slides:
        count += fill_ppt_slide_variables(slide, variables)
    return count


def fill_ppt_slide_variables(slide: Any, variables: dict[str, Any]) -> int:
    """填充单张幻灯片（python-pptx Slide）中的占位符，返回替换数量。"""
    count = 0
    for shape in slide.shapes:
        if shape.has_text_frame:
            count += _replace_in_text_frame(shape.text_frame, variables)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    count += _replace_in_text_frame(cell.text_frame, variables)
    return count


def _replace_in_text_frame(text_frame: Any, variables: dict[str, Any]) -> int:
    """在文本框中替换占位符，返回替换次数。"""
    count = 0
    for paragraph in text_frame.paragraphs:
        runs = list(paragraph.runs)
        if runs and "{{" in "".join(run.text for run in runs):
            count += _replace_in_runs(runs, variables)
    return count


def _replace_in_runs(runs: list[Any], variables: dict[str, Any]) -> int:
    """在一个段落/文本行的 run 序列中替换占位符。

    优先在单个 run 内替换以保留原有格式；
    若占位符被拆分到多个 run（如 {{ti | tle}}），则重组到首个 run 并清空其余 run。
    """
    if not runs:
        return 0

    count = 0
    for run in runs:
        if "{{" in run.text:
            new_text = replace_placeholders(run.text, variables)
            if new_text != run.text:
                run.text = new_text
                count += 1

    # 逐 run 替换后仍有未闭合占位符，说明被拆分为多个 run，进行段落级重组
    joined = "".join(run.text for run in runs)
    if "{{" in joined:
        replaced = replace_placeholders(joined, variables)
        if replaced != joined:
            runs[0].text = replaced
            for run in runs[1:]:
                run.text = ""
            count += 1
    return count
