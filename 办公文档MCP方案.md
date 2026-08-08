# 办公文档读写分析 MCP 方案

> **定位**：面向 Word / Excel / PowerPoint / PDF 四大格式的全场景读写与分析 MCP Server
> **边界**：不包含格式转换（Word→PDF 等），由独立的转换 MCP 负责
> **日期**：2026-07-29

---

## 目录

1. [背景与目标](#1-背景与目标)
2. [竞品调研分析](#2-竞品调研分析)
3. [技术选型](#3-技术选型)
4. [整体架构设计](#4-整体架构设计)
5. [工具集设计](#5-工具集设计)
6. [资源与提示词设计](#6-资源与提示词设计)
7. [安全与防护设计](#7-安全与防护设计)
8. [性能与可靠性](#8-性能与可靠性)
9. [部署方案](#9-部署方案)
10. [实施路线图](#10-实施路线图)
11. [附录](#11-附录)

---

## 1. 背景与目标

### 1.1 问题背景

在企业办公场景中，AI 助手需要频繁操作 Word、Excel、PPT、PDF 等文档。当前生态中存在多个零散的 MCP Server，但存在以下问题：

| 问题 | 说明 |
|------|------|
| **功能碎片化** | Word、Excel、PPT、PDF 各自独立为不同 MCP Server，配置和维护成本高 |
| **能力不均衡** | Word/Excel 生态成熟，PDF 和 PPT 的读写+分析能力普遍薄弱 |
| **分析能力缺失** | 大多数 MCP 只做"读写"，缺乏文档结构分析、内容提取、智能统计等分析能力 |
| **工具粒度不合理** | 要么过粗（一个工具做太多事），要么过细（调用链冗长） |
| **安全边界模糊** | 路径越权、大文件 OOM、敏感信息泄露等风险缺乏系统性防护 |

### 1.2 设计目标

构建一个**统一的办公文档 MCP Server**，实现：

- ✅ **四格式全覆盖**：Word (.docx) / Excel (.xlsx) / PowerPoint (.pptx) / PDF (.pdf)
- ✅ **读+写+分析三维度**：不止读写，更提供结构分析、内容提取、统计摘要等分析能力
- ✅ **模板系统**：四格式均支持模板传入与变量填充，统一占位符规范 `{{variable}}`
- ✅ **原子化工具设计**：每个工具单一职责，LLM 可灵活组合调用
- ✅ **安全可控**：路径沙箱、输入校验、大小限制、操作审计
- ✅ **高性能**：大文件流式处理、Session 内存编辑、合理的超时机制
- ✅ **不含格式转换**：格式转换由独立 MCP 负责，职责清晰

### 1.3 核心设计原则

```
┌─────────────────────────────────────────────────────┐
│              设计原则金字塔                           │
├─────────────────────────────────────────────────────┤
│  1. 原子性 — 每个工具做一件事，做好一件事             │
│  2. 可组合 — LLM 能自由串联工具完成复杂工作流          │
│  3. 安全优先 — 路径沙箱 + 输入校验 + 操作审计         │
│  4. 描述驱动 — 工具描述和 Schema 即文档，LLM 靠它理解 │
│  5. 统一命名 — {format}_{verb}_{object} 命名规范     │
│  6. 渐进增强 — 先核心读写，再分析能力，再高级特性      │
└─────────────────────────────────────────────────────┘
```

---

## 2. 竞品调研分析

### 2.1 现有方案全景

对当前主流的办公文档 MCP Server 进行了系统调研，按覆盖范围和技术栈分类如下：

#### 2.1.1 单格式专项 MCP

| 项目 | 格式 | 技术栈 | 工具数 | 特点 | Star |
|------|------|--------|--------|------|------|
| GongRzhe/Office-Word-MCP-Server | Word | Python + FastMCP + python-docx | ~16 | 模块化架构，PyPI 发布，uvx 支持 | 960+ |
| GongRzhe/Office-PowerPoint-MCP-Server | PPT | Python + FastMCP + python-pptx | 32 | v2.0，11 个模块，模板系统，图片效果 | 300+ |
| haris-musa/excel-mcp-server | Excel | Python + FastMCP + openpyxl | ~20 | 双传输(stdio+SSE)，无需安装 Excel | 200+ |

#### 2.1.2 多格式合并 MCP

| 项目 | 格式覆盖 | 技术栈 | 特点 | 不足 |
|------|----------|--------|------|------|
| theWDY/office-editor-mcp | Word + Excel + PPT | Python + FastMCP | 含 OCR、文档比较、翻译、加密 | 无 PDF 支持；工具描述较简略 |
| cuoicungtui/office-mcp | Word + Excel + PPT | Python + FastMCP | 参考 GongRzhe 方案，基础功能完整 | 无 PDF；无分析能力 |
| ForLegalAI/mcp-ms-office-documents | Word + Excel + PPT + Email + XML | Python + Docker | Markdown→文档，模板系统，云存储上传 | 只写不读；无 PDF；无分析 |
| Aspose MCP Server | Word + Excel + PPT + PDF | .NET 8.0 + Aspose.Total | 88 个工具，Session 管理，认证机制 | 商业授权；含格式转换(越界) |

#### 2.1.3 优劣势总结

```
优势共性                          劣势共性
─────────────────────────────────────────────────────
✅ 基于 python-docx/openpyxl/     ❌ PDF 支持普遍薄弱
   python-pptx，技术栈统一          （仅 Aspose 有完整 PDF，
✅ FastMCP 装饰器模式，开发效率高      但需商业授权）
✅ 支持 stdio/SSE 双传输            ❌ "分析"能力几乎为零
✅ uvx 一键安装，部署便捷           ❌ 工具命名不统一，
✅ 活跃社区，持续维护                  跨格式协作困难
                                  ❌ 安全防护不足
                                     （路径越权、无大小限制）
                                  ❌ 碎片化严重，
                                     用户需配置多个 MCP
```

### 2.2 关键发现

1. **PDF 是最大短板**：现有方案要么不支持 PDF，要么仅支持基础读取。PyMuPDF 性能最强但 AGPL 许可证是商业风险，pdfplumber 表格提取优秀但无写入能力，需要组合方案。

2. **"分析"是蓝海**：几乎没有 MCP 提供文档结构分析、内容摘要统计、智能数据提取等分析能力。这是与竞品差异化的关键。

3. **统一 MCP 优于多个分散 MCP**：用户更倾向于一个配置搞定所有格式，而非为 Word/Excel/PPT/PDF 分别配置 4 个 MCP。

4. **Aspose 方案值得参考**：其 Session 管理（内存编辑避免频繁磁盘 IO）、按需启用的设计思路很好，但 .NET + 商业授权的门槛太高。

5. **FastMCP 是事实标准**：几乎所有 Python 生态的办公 MCP 都基于 FastMCP，生态成熟，文档完善，是技术选型的安全选择。

---

## 3. 技术选型

### 3.1 技术栈总览

| 层面 | 选型 | 理由 |
|------|------|------|
| **语言** | Python 3.12+ | 办公文档库生态最丰富，FastMCP 原生支持 |
| **MCP 框架** | FastMCP | 事实标准，装饰器模式，自动 Schema 生成，支持 stdio/HTTP/SSE |
| **Word 处理** | python-docx | MIT 许可，成熟稳定，支持段落/表格/样式/图片 |
| **Excel 处理** | openpyxl | MIT 许可，支持读写/公式/图表/样式/数据透视表 |
| **PPT 处理** | python-pptx | MIT/BSD 许可，支持幻灯片/形状/图表/表格/模板 |
| **PDF 读取** | pdfplumber + pypdf | pdfplumber 精准文本/表格提取，pypdf 处理合并/拆分/加密 |
| **PDF 写入** | reportlab | 生成 PDF 报告/文档，支持复杂排版 |
| **PDF 高性能** | PyMuPDF (可选) | AGPL 许可，仅在对性能有极高要求时启用 |
| **数据增强** | pandas | Excel 数据分析、统计、透视 |
| **图像处理** | Pillow | PPT 图片增强、PDF 图片提取后处理 |

### 3.2 库选型决策矩阵

#### Word (.docx)

| 库 | 读 | 写 | 分析 | 许可证 | 选型 |
|----|----|----|------|--------|------|
| python-docx | ✅ | ✅ | 基础 | MIT | ✅ 主选 |
| docx2txt | ✅ | ❌ | 文本提取 | MIT | 辅助 |
| spire-doc | ✅ | ✅ | 丰富 | 商业(免费版限10页) | ❌ 不选 |

#### Excel (.xlsx)

| 库 | 读 | 写 | 分析 | 许可证 | 选型 |
|----|----|----|------|--------|------|
| openpyxl | ✅ | ✅ | 基础 | MIT | ✅ 主选 |
| pandas | ✅ | ✅ | 强大 | BSD | ✅ 辅助(数据分析) |
| xlsxwriter | ❌ | ✅ | ❌ | BSD | ❌ 仅写不读 |
| xlrd/xlwt | .xls | .xls | ❌ | MIT | ❌ 旧格式不选 |

#### PowerPoint (.pptx)

| 库 | 读 | 写 | 分析 | 许可证 | 选型 |
|----|----|----|------|--------|------|
| python-pptx | ✅ | ✅ | 基础 | MIT/BSD | ✅ 主选 |
| Spire.Presentation | ✅ | ✅ | 丰富 | 商业(免费版限10页) | ❌ 不选 |

#### PDF (.pdf)

| 库 | 读 | 写 | 分析 | 许可证 | 选型 |
|----|----|----|------|--------|------|
| pdfplumber | ✅ | ❌ | 表格/文本 | MIT | ✅ 主选(读取) |
| pypdf | ✅ | ✅ | 基础 | MIT | ✅ 主选(操作) |
| reportlab | ❌ | ✅ | ❌ | Commercial/Free | ✅ 主选(生成) |
| PyMuPDF | ✅ | ✅ | 强大 | AGPL/Commercial | ⚠️ 可选(高性能场景) |
| pymupdf4llm | ✅ | ❌ | Markdown | AGPL/Commercial | ⚠️ 可选(LLM优化) |
| camelot | ✅ | ❌ | 表格 | MIT | ❌ 依赖 ghostscript |

### 3.3 PDF 组合策略

PDF 是最复杂的格式，单一库无法覆盖所有需求，采用**分层组合策略**：

```
PDF 处理分层架构
┌──────────────────────────────────────────┐
│            MCP 工具层                      │
├──────────┬──────────┬──────────┬─────────┤
│ 文本提取  │ 表格提取  │ 文档操作  │ PDF生成 │
│pdfplumber│pdfplumber│  pypdf   │reportlab│
│          │          │          │         │
│  +pypdf  │          │ 合并/拆分 │ 报告/   │
│ (metadata│          │ 旋转/加密 │ 文档    │
│  /pages) │          │ 水印     │         │
└──────────┴──────────┴──────────┴─────────┘
         ↓ 可选高性能层 ↓
┌──────────────────────────────────────────┐
│      PyMuPDF (AGPL, 需商业授权)            │
│  高性能文本提取 + 页面渲染 + 表格 + OCR    │
│  仅在处理 >1000 页或需要 OCR 时启用       │
└──────────────────────────────────────────┘
```

### 3.4 PDF 写入技术路径（Overlay 合并模式）

PDF 与 Word/Excel/PPT 不同，已有 PDF 的原地修改能力极有限。pypdf 仅支持页面级操作（合并/拆分/旋转/加密），reportlab 只能从零生成。向**已有 PDF 添加内容**（文本/图片/水印/注释）采用 **Overlay 合并模式**：

```
PDF 写入 Overlay 合合流程
┌──────────────────────────────────────────────────────┐
│  1. 使用 pypdf.PdfReader 读取原始 PDF                  │
│  2. 使用 reportlab 在内存中生成 overlay 页              │
│     ├─ io.BytesIO 作为 reportlab canvas 输出缓冲        │
│     ├─ 在空白页面上绘制要添加的内容（文本/图片/水印）     │
│     └─ 保存为临时 PDF（仅含 overlay 层）                │
│  3. 使用 pypdf.PageObject.merge_page() 将 overlay       │
│     叠加到原始页面上                                    │
│  4. 使用 pypdf.PdfWriter 写入合并后的 PDF               │
└──────────────────────────────────────────────────────┘
```

```python
# PDF Overlay 合并示例：向已有 PDF 添加水印
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import io

def pdf_add_overlay(original_path: str, overlay_func, output_path: str):
    """向已有 PDF 添加 overlay 内容"""
    reader = PdfReader(original_path)
    writer = PdfWriter()

    for page in reader.pages:
        # 1. 生成 overlay 页
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        overlay_func(can, page.mediabox)  # 回调函数绘制内容
        can.save()
        packet.seek(0)

        # 2. 合并 overlay 到原页面
        overlay_page = PdfReader(packet).pages[0]
        page.merge_page(overlay_page)
        writer.add_page(page)

    # 3. 写入输出
    with open(output_path, "wb") as f:
        writer.write(f)
```

| PDF 写入工具 | 实现方式 | 说明 |
|-------------|---------|------|
| `pdf_add_text` | reportlab canvas 绘制文本 + pypdf merge_page | 在指定页面位置叠加文本 |
| `pdf_add_image` | reportlab canvas drawImage + pypdf merge_page | 在指定页面位置叠加图片 |
| `pdf_add_watermark` | reportlab canvas 全页绘制水印 + pypdf merge_page | 遍历所有页面叠加水印 |
| `pdf_add_annotation` | pypdf AnnotationBuilder | pypdf 原生支持注释（高亮/链接/文本批注） |
| `pdf_add_bookmark` | pypdf writer.add_outline_item | pypdf 原生支持大纲/书签层级结构 |

---

## 4. 整体架构设计

### 4.1 架构总览

```
                    ┌─────────────────────┐
                    │   MCP Client (LLM)  │
                    │  Claude/Cursor/etc  │
                    └─────────┬───────────┘
                              │ JSON-RPC 2.0
                              │ (stdio / SSE / HTTP)
                    ┌─────────▼───────────┐
                    │   FastMCP Server    │
                    │  (统一入口 + 路由)   │
                    └─────────┬───────────┘
                              │
           ┌──────────┬───────┴───────┬──────────┐
           │          │               │          │
    ┌──────▼──┐ ┌─────▼────┐ ┌───────▼──┐ ┌────▼─────┐
    │  Word   │ │  Excel   │ │    PPT   │ │   PDF    │
    │ Handler │ │ Handler  │ │ Handler  │ │ Handler  │
    │ Module  │ │ Module   │ │ Module   │ │ Module   │
    └────┬────┘ └────┬─────┘ └────┬─────┘ └────┬─────┘
         │           │            │            │
    ┌────▼────┐ ┌────▼─────┐ ┌───▼────┐ ┌────▼─────┐
    │python-  │ │openpyxl +│ │python- │ │pdfplumber│
    │  docx   │ │  pandas  │ │  pptx  │ │+ pypdf + │
    │         │ │          │ │        │ │reportlab │
    └─────────┘ └──────────┘ └────────┘ └──────────┘

    ┌──────────────────────────────────────────────┐
    │               公共服务层                      │
    │  ┌─────────┐ ┌──────────┐ ┌───────────────┐ │
    │  │路径沙箱  │ │输入校验   │ │ Session 管理  │ │
    │  │PathGuard│ │Validator │ │ SessionMgr   │ │
    │  └─────────┘ └──────────┘ └───────────────┘ │
    │  ┌─────────┐ ┌──────────┐ ┌───────────────┐ │
    │  │审计日志  │ │错误处理   │ │ 资源/模板管理 │ │
    │  │ AuditLog│ │ErrHandler│ │ ResourceMgr  │ │
    │  └─────────┘ └──────────┘ └───────────────┘ │
    └──────────────────────────────────────────────┘
```

### 4.2 项目结构

```
timeverse-office-doc-mcp/
├── pyproject.toml              # 项目配置 + 依赖
├── README.md
├── src/
│   └── timeverse_office_doc_mcp/
│       ├── __init__.py
│       ├── server.py           # FastMCP Server 主入口
│       ├── config.py           # 配置管理（路径白名单、大小限制等）
│       │
│       ├── handlers/           # 四大格式处理器
│       │   ├── __init__.py
│       │   ├── word_handler.py     # Word 工具实现
│       │   ├── excel_handler.py    # Excel 工具实现
│       │   ├── ppt_handler.py      # PowerPoint 工具实现
│       │   └── pdf_handler.py      # PDF 工具实现
│       │
│       ├── common/             # 公共服务层
│       │   ├── __init__.py
│       │   ├── path_guard.py       # 路径安全沙箱
│       │   ├── validator.py        # 输入校验器
│       │   ├── session.py          # Session 管理（内存编辑 + TTL 清理）
│       │   ├── audit_log.py        # 操作审计日志
│       │   ├── error_handler.py    # 统一错误处理
│       │   ├── resource_mgr.py     # 资源 & 模板管理
│       │   ├── file_lock.py        # 文件并发锁
│       │   ├── auth.py             # HTTP/SSE 认证中间件
│       │   ├── rate_limiter.py     # 速率限制器
│       │   ├── sanitizer.py        # 敏感数据脱敏
│       │   └── metrics.py          # 监控指标收集
│       │
│       ├── resources/          # MCP Resources
│       │   ├── __init__.py
│       │   └── document_resources.py
│       │
│       └── prompts/            # MCP Prompts
│           ├── __init__.py
│           └── document_prompts.py
│
├── templates/                  # 文档模板库
│   ├── word/
│   ├── excel/
│   ├── ppt/
│   └── pdf/
│
├── tests/
│   ├── test_word_handler.py
│   ├── test_excel_handler.py
│   ├── test_ppt_handler.py
│   ├── test_pdf_handler.py
│   └── test_common.py
│
└── docker/
    ├── Dockerfile
    └── docker-compose.yml
```

### 4.3 命名规范

所有工具统一采用 `{format}_{verb}_{object}` 命名模式：

| 前缀 | 格式 | 示例 |
|------|------|------|
| `word_` | Word 文档 | `word_create_document`, `word_add_paragraph` |
| `excel_` | Excel 表格 | `excel_create_workbook`, `excel_write_cell` |
| `ppt_` | PowerPoint | `ppt_create_presentation`, `ppt_add_slide` |
| `pdf_` | PDF 文档 | `pdf_extract_text`, `pdf_add_watermark` |
| `doc_` | 跨格式通用 | `doc_list_templates`, `doc_apply_template`, `doc_analyze_structure` |

---

## 5. 工具集设计

### 5.1 Word 工具集（21 个工具）

#### 5.1.1 文档管理

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `word_create_document` | 创建新文档（支持模板） | filename, title, author, template, variables | 写 |
| `word_get_info` | 获取文档元信息 | filename | 读 |
| `word_get_text` | 提取全文文本 | filename, include_tables | 读 |
| `word_get_outline` | 获取文档大纲结构 | filename | 读 |
| `word_list_documents` | 列出目录内文档 | directory | 读 |
| `word_copy_document` | 复制文档 | source, destination | 写 |

#### 5.1.2 内容编辑

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `word_add_heading` | 添加标题 | filename, text, level | 写 |
| `word_add_paragraph` | 添加段落 | filename, text, style, font_size, bold | 写 |
| `word_add_table` | 添加表格 | filename, rows, cols, data, has_header | 写 |
| `word_add_image` | 插入图片 | filename, image_path, width | 写 |
| `word_add_page_break` | 插入分页符 | filename | 写 |
| `word_add_list` | 添加列表 | filename, items, list_style | 写 |
| `word_set_header_footer` | 设置页眉页脚 | filename, header_text, footer_text, include_page_num | 写 |
| `word_generate_toc` | 生成目录（Table of Contents） | filename, max_level, styles | 写 |

#### 5.1.3 格式化与操作

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `word_format_text` | 格式化文本片段 | filename, paragraph_idx, start, end, bold, italic, color, font | 写 |
| `word_format_table` | 格式化表格 | filename, table_idx, border_style, header_row, shading | 写 |
| `word_search_replace` | 搜索替换 | filename, find_text, replace_text | 写 |
| `word_delete_paragraph` | 删除段落 | filename, paragraph_idx | 写 |
| `word_create_style` | 创建自定义样式 | filename, style_name, font, size, color, bold | 写 |

#### 5.1.4 分析工具

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `word_analyze_structure` | 分析文档结构（标题层级、段落分布、表格统计） | filename | 分析 |
| `word_extract_tables` | 提取所有表格数据 | filename, format(json/csv) | 分析 |

### 5.2 Excel 工具集（24 个工具）

#### 5.2.1 工作簿管理

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `excel_create_workbook` | 创建工作簿（支持模板） | filename, sheet_name, template, variables | 写 |
| `excel_get_info` | 获取工作簿信息 | filename | 读 |
| `excel_list_sheets` | 列出工作表 | filename | 读 |
| `excel_add_sheet` | 添加工作表 | filename, sheet_name | 写 |
| `excel_delete_sheet` | 删除工作表 | filename, sheet_name | 写 |
| `excel_rename_sheet` | 重命名工作表 | filename, old_name, new_name | 写 |
| `excel_copy_sheet` | 复制工作表 | filename, source, target | 写 |

#### 5.2.2 数据读写

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `excel_read_cell` | 读取单元格 | filename, sheet, cell_ref | 读 |
| `excel_write_cell` | 写入单元格 | filename, sheet, cell_ref, value | 写 |
| `excel_read_range` | 读取区域数据 | filename, sheet, range_str | 读 |
| `excel_write_range` | 批量写入区域 | filename, sheet, start_cell, data | 写 |
| `excel_insert_row` | 插入行 | filename, sheet, row_idx, count | 写 |
| `excel_delete_row` | 删除行 | filename, sheet, row_idx, count | 写 |
| `excel_insert_column` | 插入列 | filename, sheet, col_idx, count | 写 |
| `excel_delete_column` | 删除列 | filename, sheet, col_idx, count | 写 |

#### 5.2.3 格式化与高级

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `excel_format_cell` | 格式化单元格 | filename, sheet, range, font, fill, border, alignment | 写 |
| `excel_apply_formula` | 应用公式 | filename, sheet, cell_ref, formula | 写 |
| `excel_create_chart` | 创建图表 | filename, sheet, chart_type, data_range, title | 写 |
| `excel_freeze_panes` | 冻结窗格 | filename, sheet, cell_ref | 写 |
| `excel_sort_data` | 排序数据 | filename, sheet, range, key_column, ascending | 写 |
| `excel_create_pivot_table` | 创建数据透视表 | filename, source_sheet, source_range, target_sheet, rows, cols, values, agg_func | 写 |
| `excel_add_conditional_format` | 添加条件格式 | filename, sheet, range, rule_type, criteria, format | 写 |

#### 5.2.4 分析工具

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `excel_analyze_data` | 数据统计分析（描述统计、空值检测、类型推断） | filename, sheet, range | 分析 |
| `excel_find_duplicates` | 查找重复数据 | filename, sheet, columns, threshold | 分析 |

### 5.3 PowerPoint 工具集（19 个工具）

#### 5.3.1 演示文稿管理

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `ppt_create_presentation` | 创建演示文稿（支持模板） | filename, template, variables | 写 |
| `ppt_get_info` | 获取演示文稿信息 | filename | 读 |
| `ppt_list_slides` | 列出幻灯片概览 | filename | 读 |
| `ppt_add_slide` | 添加幻灯片 | filename, layout, title | 写 |
| `ppt_delete_slide` | 删除幻灯片 | filename, slide_idx | 写 |
| `ppt_move_slide` | 移动幻灯片 | filename, slide_idx, new_idx | 写 |
| `ppt_copy_slide` | 复制幻灯片 | filename, slide_idx | 写 |

#### 5.3.2 内容编辑

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `ppt_add_text` | 添加文本框 | filename, slide_idx, text, position, font_size, bold | 写 |
| `ppt_add_image` | 插入图片 | filename, slide_idx, image_path, position, size | 写 |
| `ppt_add_table` | 添加表格 | filename, slide_idx, rows, cols, data | 写 |
| `ppt_add_chart` | 添加图表 | filename, slide_idx, chart_type, data, title | 写 |
| `ppt_add_shape` | 添加形状 | filename, slide_idx, shape_type, position, size | 写 |
| `ppt_set_background` | 设置幻灯片背景 | filename, slide_idx, color or image | 写 |
| `ppt_apply_theme` | 应用主题配色 | filename, theme_name | 写 |
| `ppt_set_slide_notes` | 设置演讲者备注 | filename, slide_idx, notes_text | 写 |

#### 5.3.3 分析工具

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `ppt_extract_text` | 提取所有幻灯片文本 | filename | 分析 |
| `ppt_get_slide_notes` | 获取演讲者备注 | filename, slide_idx | 分析 |
| `ppt_analyze_structure` | 分析结构（幻灯片分布、元素统计、布局分析） | filename | 分析 |
| `ppt_get_structure` | 获取完整结构树 | filename | 分析 |

### 5.4 PDF 工具集（19 个工具）

#### 5.4.1 文档管理

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `pdf_get_info` | 获取 PDF 元信息（页数、作者、标题等） | filename | 读 |
| `pdf_merge` | 合并多个 PDF | files[], output | 写 |
| `pdf_split` | 拆分 PDF | filename, page_ranges, output_prefix | 写 |
| `pdf_rotate_page` | 旋转页面 | filename, page_idx, angle | 写 |

#### 5.4.2 内容读取

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `pdf_extract_text` | 提取文本（支持指定页范围） | filename, page_range, layout_mode | 分析 |
| `pdf_extract_tables` | 提取表格数据 | filename, page_range, format | 分析 |
| `pdf_extract_images` | 提取图片 | filename, page_range, output_dir | 分析 |
| `pdf_search_text` | 搜索文本 | filename, query, case_sensitive | 分析 |
| `pdf_ocr_text` | OCR 文本识别（扫描件/图片型 PDF） | filename, page_range, lang, output_format | 分析 |

#### 5.4.3 内容写入

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `pdf_add_text` | 添加文本到页面 | filename, page_idx, text, position, font, size | 写 |
| `pdf_add_image` | 添加图片到页面 | filename, page_idx, image_path, position, size | 写 |
| `pdf_add_watermark` | 添加水印 | filename, watermark_text, opacity, font_size | 写 |
| `pdf_add_annotation` | 添加注释 | filename, page_idx, annotation_type, content, position | 写 |
| `pdf_add_bookmark` | 添加书签 | filename, title, page_idx | 写 |

#### 5.4.4 安全与分析

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `pdf_encrypt` | 加密 PDF | filename, password, permissions | 写 |
| `pdf_decrypt` | 解密 PDF | filename, password | 写 |
| `pdf_analyze_structure` | 分析结构（页面类型、文本密度、表格分布） | filename | 分析 |

#### 5.4.5 模板工具

| 工具名 | 功能 | 关键参数 | 读/写 |
|--------|------|----------|-------|
| `pdf_create_from_template` | 从模板创建 PDF（变量填充） | template_name, variables, output | 写 |
| `pdf_fill_form` | 填充 PDF 交互式表单字段（AcroForm） | filename, fields, flatten | 写 |

### 5.5 模板管理工具集（6 个工具）

> 跨格式的统一模板管理系统，支持注册、预览、应用和占位符提取。所有格式共用同一套模板管理接口，降低 LLM 的认知负担。

#### 5.5.1 模板注册与管理

| 工具名 | 功能 | 关键参数 | 说明 |
|--------|------|----------|------|
| `doc_list_templates` | 列出所有可用模板 | format (可选，过滤格式) | 返回模板名称、格式、描述 |
| `doc_get_template_info` | 获取模板详情（含占位符列表与预览） | template_name | 返回占位符变量清单 + 结构预览 |
| `doc_register_template` | 注册新模板到模板库 | name, format, file_path, description, placeholders | 将外部文件注册为模板 |
| `doc_delete_template` | 从模板库删除模板 | template_name | 删除注册索引记录与模板库副本，不影响用户原始文件 |

**模板存储设计**：

- **模板库目录**：`templates/{word|excel|ppt|pdf}/`，按格式分子目录存放模板文件（见 4.2 项目结构）。
- **注册语义**：`doc_register_template` 将外部 `file_path` **复制**到 `templates/{format}/` 下并按 `name` 命名，随后把 `{name, format, path, description, placeholders}` 写入注册索引 `templates/registry.json`。
- **索引解析**：`template_name` 为唯一逻辑键，`doc_apply_template`、`doc_list_templates`、`templates://` 资源统一通过索引解析模板文件路径，对 LLM 不暴露磁盘路径。
- **删除语义**：`doc_delete_template` 删除注册索引记录并清理模板库副本；注册时传入的用户原始文件不受影响。
- **安全**：模板副本位于 PathGuard 白名单目录内，天然满足路径沙箱约束，无需对外部源路径额外放行。

#### 5.5.2 模板应用与占位符

| 工具名 | 功能 | 关键参数 | 说明 |
|--------|------|----------|------|
| `doc_apply_template` | 从模板创建文档并自动填充变量 | template_name, output_path, variables | 核心工具：加载模板 → 替换占位符 → 保存输出 |
| `doc_extract_placeholders` | 扫描提取模板中的占位符变量 | template_name, format | 返回所有 `{{...}}` 占位符及其类型推断 |

### 5.6 Session 管理工具集（4 个工具）

> 跨格式的 Session 管理系统，支持在内存中编辑文档，避免每次操作都读写磁盘。适用于需要连续多次编辑同一文档的场景（如批量添加段落、表格、图片）。

| 工具名 | 功能 | 关键参数 | 说明 |
|--------|------|----------|------|
| `doc_open_session` | 打开文档到内存 Session | filename, format | 返回 session_id，后续编辑工具可通过 session_id 操作内存中的文档 |
| `doc_save_session` | 保存 Session 到磁盘 | session_id, output_path (可选) | 将内存中的修改写入文件；不指定 output_path 则保存回原路径 |
| `doc_close_session` | 关闭 Session | session_id, save (默认 False) | 关闭并释放内存；save=True 时先保存再关闭 |
| `doc_list_sessions` | 列出所有活跃 Session | 无 | 返回所有未关闭的 Session 列表（含格式、路径、修改状态） |

**Session 与编辑工具的协作模式**：

所有编辑类工具（如 `word_add_paragraph`、`excel_write_cell` 等）新增可选参数 `session_id`：

```python
# 模式一：传统模式（无 session_id，每次操作读写磁盘）
word_add_heading(filename="report.docx", text="第一章", level=1)
word_add_paragraph(filename="report.docx", text="内容...")
# 每次调用：打开文件 -> 修改 -> 保存 -> 关闭（磁盘 IO: 2次/调用）

# 模式二：Session 模式（传入 session_id，在内存中操作）
session_id = doc_open_session(filename="report.docx", format="word")
word_add_heading(filename="report.docx", text="第一章", level=1, session_id=session_id)
word_add_paragraph(filename="report.docx", text="内容...", session_id=session_id)
word_add_table(filename="report.docx", rows=3, cols=3, data=data, session_id=session_id)
doc_save_session(session_id=session_id)
doc_close_session(session_id=session_id)
# 仅 1 次读取 + 1 次写入，中间操作全部在内存中完成
```

> **工具内部逻辑**：当 `session_id` 参数存在时，从 SessionManager 获取内存中的文档对象直接操作；当 `session_id` 不存在时，走传统的「打开-修改-保存」流程。对 LLM 透明，不传 `session_id` 也能正常工作。

### 5.7 模板填充机制设计

#### 5.7.1 占位符规范

模板通过统一的占位符语法实现动态内容填充，所有格式共用同一套规范：

| 占位符语法 | 示例 | 说明 | 适用格式 |
|-----------|------|------|---------|
| `{{variable}}` | `{{title}}`, `{{author}}` | 简单文本替换 | Word / Excel / PPT / PDF |
| `{{date:format}}` | `{{date:%Y-%m-%d}}` | 日期格式化填充 | Word / Excel / PPT / PDF |
| `{{number:format}}` | `{{number:,.2f}}` | 数字格式化（千分位/小数） | Word / Excel / PPT / PDF |
| `{{table:name}}` | `{{table:sales_data}}` | 表格数据填充（二维数组） | Word / PPT / PDF |
| `{{image:name}}` | `{{image:logo}}` | 图片插入（路径或 base64） | Word / PPT |
| `{{condition:expr}}` | `{{condition:show_summary}}` | 条件显隐（布尔值控制段落/行） | Word / PPT |

#### 5.7.2 各格式实现方式

| 格式 | 模板加载方式 | 占位符替换范围 | 技术实现 |
|------|-------------|--------------|---------|
| **Word** | `Document(template_path)` 打开 .docx | 段落 / 表格单元格 / 页眉 / 页脚 / 文本框 | 遍历 `paragraphs` + `tables`，正则匹配 `{{...}}`，保留原样式替换 |
| **Excel** | `load_workbook(template_path)` 加载 .xlsx | 单元格值 / 公式 / 批注 | 遍历所有 worksheet 的 cells，匹配 `{{...}}` 替换值 |
| **PPT** | `Presentation(template_path)` 加载 .pptx | 文本框 / 表格 / 备注页 | 遍历 `slides` → `shapes` → `text_frame`，匹配并替换，保留原格式 |
| **PDF** | AcroForm 表单字段填充 + reportlab 布局模板 | 表单字段 / 文本区域 / 图片区域 | `pypdf` 填充交互表单；`reportlab` 使用预定义 Flowable 模板填充变量内容 |

#### 5.7.3 变量定义与填充流程

```python
# 模板变量定义示例
variables = {
    # 简单文本
    "title": "2026年第三季度运营报告",
    "author": "数据分析团队",
    "department": "运营部",
    
    # 日期格式化
    "date": "2026-09-30",           # 原始值
    # 模板中 {{date:%Y年%m月%d日}} → "2026年09月30日"
    
    # 数字格式化
    "total_revenue": 1250000.5,     # 原始值
    # 模板中 {{total_revenue:,.2f}} → "1,250,000.50"
    
    # 表格数据填充
    "table:sales_summary": [
        ["区域", "销售额", "增长率"],
        ["华东", "¥450,000", "+12.5%"],
        ["华南", "¥380,000", "+8.3%"],
        ["华北", "¥420,000", "+15.1%"],
    ],
    
    # 图片插入
    "image:logo": "/path/to/company_logo.png",
    "image:chart": "/path/to/q3_chart.png",
    
    # 条件显隐
    "condition:show_summary": True,   # True → 显示该段落/幻灯片
    "condition:show_appendix": False,  # False → 隐藏该段落/幻灯片
}

# 调用模板应用工具
doc_apply_template(
    template_name="quarterly_report",
    output_path="Q3_report.docx",
    variables=variables
)
```

**填充执行流程**：

```
doc_apply_template 调用流程
┌─────────────────────────────────────────────────────┐
│ 1. 解析 template_name → 获取格式 + 模板文件路径      │
│ 2. 路径安全校验（PathGuard）                         │
│ 3. 按格式加载模板到内存                               │
│    ├─ Word:  Document(template_path)                │
│    ├─ Excel: load_workbook(template_path)           │
│    ├─ PPT:   Presentation(template_path)            │
│    └─ PDF:   AcroForm 检测 / reportlab 模板加载     │
│ 4. 遍历文档元素，正则匹配 {{...}} 占位符              │
│ 5. 按占位符类型分派替换                               │
│    ├─ 文本/日期/数字 → 直接替换值                     │
│    ├─ table:name → 插入表格（保留模板表格样式）       │
│    ├─ image:name → 插入图片到占位位置                 │
│    └─ condition:expr → 按布尔值显隐段落/幻灯片        │
│ 6. 保存到 output_path                                │
│ 7. 返回填充结果摘要（替换变量数 + 输出路径）          │
└─────────────────────────────────────────────────────┘
```

#### 5.7.4 PDF 模板两种模式

PDF 模板支持两种互补模式，覆盖不同场景：

| 模式 | 适用场景 | 技术实现 | 优势 | 局限 |
|------|---------|---------|------|------|
| **AcroForm 表单填充** | 已有 PDF 表单（合同、申请表、发票模板） | `pypdf` 读取表单字段 → 填值 → 可选扁平化(flatten) | 精确填充预设字段，保留原版式 | 需源 PDF 已含表单字段 |
| **reportlab 布局模板** | 从零生成结构化报告（月报、分析报告） | 预定义 Flowable 布局（标题/段落/表格/图表区域）→ 变量注入 | 完全可编程，灵活度高 | 需预先编写布局模板代码 |

```python
# PDF AcroForm 表单填充示例
pdf_fill_form(
    filename="contract_template.pdf",
    fields={
        "party_a": "腾讯科技（深圳）有限公司",
        "party_b": "XX科技有限公司",
        "contract_date": "2026-07-30",
        "amount": "¥500,000",
    },
    flatten=True  # 填充后扁平化，防止二次编辑
)

# PDF reportlab 布局模板示例
pdf_create_from_template(
    template_name="monthly_report",  # 预注册的 reportlab 布局模板
    variables={
        "title": "2026年7月运营月报",
        "summary": "本月整体运营情况良好...",
        "table:kpi_data": [["指标", "目标", "实际", "达成率"], ...],
        "chart_image": "/path/to/chart.png",
    },
    output="2026_07_monthly_report.pdf"
)
```

### 5.7.5 模板边缘情况处理规范

#### 变量缺失行为

| 场景 | 行为 | 说明 |
|------|------|------|
| 变量未提供 | 替换为空字符串 `""` | 不报错，静默清除占位符 |
| 变量值为 `None` | 替换为空字符串 `""` | 与未提供一致 |
| 变量值为 `False`（条件占位符） | 隐藏对应段落/行/幻灯片 | `{{condition:xxx}}` 专属行为 |
| 变量类型不匹配 | 尝试 str() 转换，失败则替换为空 | 如模板期望数字但收到字符串 |

#### 占位符转义

当文档内容本身包含 `{{` 时，使用双花括号转义：

| 原文 | 渲染结果 | 说明 |
|------|---------|------|
| `\{\{not_a_variable\}\}` | `{{not_a_variable}}` | 转义后的字面量 |
| `{{variable}}` | 变量值 | 正常占位符 |

#### 循环/列表填充

支持动态数量的列表项填充，适用于不定数量的段落、列表行、幻灯片：

| 占位符语法 | 示例 | 说明 | 适用格式 |
|-----------|------|------|---------|
| `{{loop:name}}...{{end:name}}` | `{{loop:items}}`- {{item.name}}: {{item.price}}`{{end:items}}` | 遍历数组，为每个元素渲染模板片段 | Word / PPT / PDF |

```python
# 循环填充示例
variables = {
    "loop:items": [
        {"name": "产品A", "price": "¥99"},
        {"name": "产品B", "price": "¥199"},
        {"name": "产品C", "price": "¥299"},
    ]
}
# 模板中：
# {{loop:items}}
# - {{name}}: {{price}}
# {{end:items}}
# 渲染后：
# - 产品A: ¥99
# - 产品B: ¥199
# - 产品C: ¥299
```

#### Excel 动态行填充

Excel 模板支持动态行扩展，通过特殊占位符标记数据起始行：

| 占位符语法 | 说明 |
|-----------|------|
| `{{row:data}}` | 标记数据插入起始位置，传入二维数组自动向下扩展行 |

```python
# Excel 动态行填充
variables = {
    "row:data": [
        ["张三", "工程部", "高级工程师"],
        ["李四", "产品部", "产品经理"],
        ["王五", "设计部", "UI设计师"],
    ]
}
# 模板中 A2 单元格为 {{row:data}}，填充后从 A2 开始向下扩展 3 行
```

### 5.8 工具总数统计

| 格式 | 管理类 | 编辑类 | 分析类 | 模板类 | 合计 |
|------|--------|--------|--------|--------|------|
| Word | 6 | 13 | 2 | - | 21 |
| Excel | 7 | 15 | 2 | - | 24 |
| PowerPoint | 7 | 8 | 4 | - | 19 |
| PDF | 4 | 7 | 6 | 2 | 19 |
| 跨格式模板管理 | - | - | - | 6 | 6 |
| 跨格式Session管理 | 4 | - | - | - | 4 |
| **合计** | **28** | **43** | **14** | **8** | **93** |

> 共 93 个工具，覆盖读、写、分析、模板、Session 五大维度。

---

## 6. 资源与提示词设计

### 6.1 Resources（资源）

Resources 是只读的、可寻址的数据，适合暴露文档元信息和模板：

```python
# 文档元信息资源
@mcp.resource("doc://{filename}/metadata")
def get_document_metadata(filename: str) -> dict:
    """获取任意格式文档的元信息"""
    # 自动检测格式并返回对应元数据

# 文档内容预览资源
@mcp.resource("doc://{filename}/preview")
def get_document_preview(filename: str) -> str:
    """获取文档内容预览（前500字符 + 结构概要）"""

# 模板库资源
@mcp.resource("templates://available")
def list_all_templates() -> list:
    """列出所有格式的可用模板"""

@mcp.resource("templates://word/available")
def list_word_templates() -> list:
    """列出可用的 Word 模板"""

@mcp.resource("templates://excel/available")
def list_excel_templates() -> list:
    """列出可用的 Excel 模板（含财务报表/数据看板等）"""

@mcp.resource("templates://ppt/available")
def list_ppt_templates() -> list:
    """列出可用的 PPT 模板（含预览描述）"""

@mcp.resource("templates://pdf/available")
def list_pdf_templates() -> list:
    """列出可用的 PDF 模板（含 AcroForm 表单 + reportlab 布局模板）"""

@mcp.resource("templates://{template_name}/info")
def get_template_detail(template_name: str) -> dict:
    """获取指定模板的详情：格式、占位符列表、结构预览"""

# 工作目录文件列表
@mcp.resource("workspace://files")
def list_workspace_files() -> list:
    """列出工作目录中的所有支持格式文件"""
```

### 6.2 Prompts（提示词）

Prompts 是可复用的对话模板，引导 LLM 以正确的方式使用工具：

```python
@mcp.prompt()
def create_word_report(title: str, topic: str, use_template: bool = False) -> str:
    """引导 LLM 创建一份结构完整的 Word 报告"""
    if use_template:
        return f"""
        请基于模板创建一份关于「{topic}」的 Word 报告，标题为「{title}」。
        
        步骤：
        1. 使用 doc_list_templates(format="word") 查看可用模板
        2. 使用 doc_get_template_info 查看模板占位符
        3. 使用 doc_apply_template 从模板创建文档，填充以下变量：
           - title: {title}
           - topic: {topic}
           - date: 当天日期
           - 其他占位符根据模板要求填充
        4. 使用 word_add_paragraph / word_add_table 补充模板未覆盖的内容
        5. 使用 word_analyze_structure 确认文档结构完整
        
        文件名：{title}.docx
        """
    return f"""
    请创建一份关于「{topic}」的 Word 报告，标题为「{title}」。
    
    要求：
    1. 使用 word_create_document 创建文档
    2. 使用 word_add_heading 添加一级标题（报告标题）
    3. 使用 word_add_heading 添加二级标题（各章节）
    4. 每个章节使用 word_add_paragraph 添加2-3段内容
    5. 在适当位置使用 word_add_table 添加数据表格
    6. 最后使用 word_analyze_structure 确认文档结构完整
    
    文件名：{title}.docx
    """

@mcp.prompt()
def analyze_excel_data(filename: str) -> str:
    """引导 LLM 分析 Excel 数据并生成报告"""
    return f"""
    请分析 Excel 文件「{filename}」中的数据：
    
    步骤：
    1. 使用 excel_get_info 获取工作簿基本信息
    2. 使用 excel_list_sheets 查看所有工作表
    3. 对每个工作表使用 excel_analyze_data 进行统计分析
    4. 使用 excel_find_duplicates 检查重复数据
    5. 汇总分析结果，指出数据质量问题和关键发现
    """

@mcp.prompt()
def extract_pdf_content(filename: str) -> str:
    """引导 LLM 提取 PDF 全部内容并结构化"""
    return f"""
    请从 PDF 文件「{filename}」中提取所有内容：
    
    步骤：
    1. 使用 pdf_get_info 获取文档基本信息
    2. 使用 pdf_extract_text 提取全部文本
    3. 使用 pdf_extract_tables 提取所有表格
    4. 使用 pdf_extract_images 提取图片（保存到输出目录）
    5. 使用 pdf_analyze_structure 分析文档结构
    6. 将提取的内容整理为结构化的摘要报告
    """

@mcp.prompt()
def create_presentation(topic: str, slide_count: int = 10, use_template: bool = False) -> str:
    """引导 LLM 从主题生成完整 PPT"""
    if use_template:
        return f"""
        请基于模板创建一个关于「{topic}」的 {slide_count} 页演示文稿。
        
        步骤：
        1. 使用 doc_list_templates(format="ppt") 查看可用 PPT 模板
        2. 使用 doc_get_template_info 查看模板占位符列表
        3. 使用 doc_apply_template 从模板创建演示文稿，填充：
           - title: {topic}
           - slide_count: {slide_count}
           - 其他占位符根据模板要求填充
        4. 使用 ppt_add_slide / ppt_add_text 补充模板未覆盖的幻灯片
        5. 使用 ppt_analyze_structure 确认结构完整
        """
    return f"""
    请围绕「{topic}」创建一个 {slide_count} 页的演示文稿。
    
    要求：
    1. 使用 ppt_create_presentation 创建文件
    2. 第一页为标题页
    3. 使用 ppt_add_slide 添加内容页，选择合适的布局
    4. 每页使用 ppt_add_text 添加要点文本
    5. 在适当位置使用 ppt_add_chart 添加数据图表
    6. 使用 ppt_apply_theme 应用专业配色方案
    7. 最后使用 ppt_analyze_structure 确认结构完整
    """
```

---

## 7. 安全与防护设计

### 7.1 路径沙箱（PathGuard）

**核心原则**：所有文件操作限制在配置的白名单目录内，防止路径遍历攻击。

```python
class PathGuard:
    """路径安全守卫 — 防止路径遍历和越权访问"""
    
    def __init__(self, allowed_dirs: list[str], blocked_patterns: list[str],
                 max_file_size: int = 100 * 1024 * 1024):
        self.allowed_dirs = [Path(d).resolve() for d in allowed_dirs]
        self.blocked_patterns = blocked_patterns  # 如 ~/.ssh, ~/.aws 等
        self.max_file_size = max_file_size  # 最大文件大小（字节）
    
    def validate_path(self, file_path: str, operation: str = "read") -> str:
        """
        校验文件路径是否安全
        
        检查项：
        1. 路径规范化（resolve .. 和符号链接）
        2. 是否在允许目录内
        3. 是否匹配黑名单模式
        4. 写操作时检查文件扩展名白名单
        5. 文件大小是否超过限制
        """
        resolved = Path(file_path).resolve()
        
        # 检查是否在允许目录内
        if not any(resolved.is_relative_to(d) for d in self.allowed_dirs):
            raise ToolError(f"路径 '{file_path}' 不在允许的目录范围内")
        
        # 检查黑名单
        for pattern in self.blocked_patterns:
            if pattern in str(resolved):
                raise ToolError(f"路径 '{file_path}' 被安全策略阻止")
        
        # 写操作检查扩展名
        if operation == "write":
            allowed_exts = {'.docx', '.xlsx', '.pptx', '.pdf'}
            if resolved.suffix not in allowed_exts:
                raise ToolError(f"不支持的文件格式: {resolved.suffix}")
        
        # 文件大小检查
        if resolved.exists() and resolved.stat().st_size > self.max_file_size:
            raise ToolError(f"文件大小超过限制 ({self.max_file_size // 1024 // 1024}MB)")
        
        return str(resolved)
```

### 7.2 输入校验（Validator）

```python
class InputValidator:
    """统一输入校验器"""
    
    @staticmethod
    def validate_filename(filename: str) -> str:
        """文件名安全校验：禁止特殊字符、控制长度"""
        if not filename or len(filename) > 255:
            raise ToolError("文件名无效或过长")
        if any(c in filename for c in ['..', '\x00', '\n', '\r']):
            raise ToolError("文件名包含非法字符")
        return filename
    
    @staticmethod
    def validate_range(range_str: str) -> str:
        """Excel 区域字符串校验：如 A1:C10"""
        import re
        if not re.match(r'^[A-Z]+\d+:[A-Z]+\d+$', range_str):
            raise ToolError(f"无效的区域格式: {range_str}")
        return range_str
    
    @staticmethod
    def validate_table_data(data: list, max_rows: int = 1000, max_cols: int = 100) -> list:
        """表格数据校验：限制行列数"""
        if len(data) > max_rows:
            raise ToolError(f"表格行数超过限制 ({max_rows})")
        for row in data:
            if len(row) > max_cols:
                raise ToolError(f"表格列数超过限制 ({max_cols})")
        return data
    
    @staticmethod
    def validate_text_length(text: str, max_length: int = 100000) -> str:
        """文本长度校验"""
        if len(text) > max_length:
            raise ToolError(f"文本长度超过限制 ({max_length} 字符)")
        return text
```

### 7.3 操作审计（AuditLog）

```python
class AuditLogger:
    """操作审计日志 — 记录所有文件操作"""
    
    def __init__(self, log_path: str):
        self.log_path = log_path
    
    def log_operation(self, tool_name: str, args: dict, 
                      result: str, duration_ms: int, 
                      success: bool, error: str = None):
        """
        记录操作日志
        - 工具名
        - 参数摘要（脱敏）
        - 执行结果状态
        - 耗时
        - 错误信息（如有）
        """
        entry = {
            "timestamp": datetime.now().isoformat(),
            "tool": tool_name,
            "args_summary": self._summarize_args(args),
            "success": success,
            "duration_ms": duration_ms,
            "error": error,
        }
        # 追加写入日志文件（JSON Lines 格式）
        with open(self.log_path, 'a') as f:
            f.write(json.dumps(entry, ensure_ascii=False) + '\n')
```

### 7.4 安全配置

```python
# config.py — 安全配置
class SecurityConfig:
    # 路径白名单（MCP 启动时配置，支持逗号分隔多目录）
    # 前缀匹配：允许目录下的所有子目录/文件，天然覆盖用户动态创建的项目目录
    ALLOWED_DIRS = [d for d in (
        *os.environ.get("OFFICE_ALLOWED_DIRS", "").split(","),  # 自定义根目录（可多个）
        os.environ.get("OFFICE_WORKSPACE", "./workspace"),
        os.environ.get("OFFICE_OUTPUT", "./output"),
        os.environ.get("OFFICE_TEMPLATES", "./templates"),
        os.getcwd(),  # 启动时工作目录兜底
    ) if d]
    
    # 路径黑名单
    BLOCKED_PATTERNS = [
        ".ssh", ".aws", ".config", ".env",
        "AppData", "Library", "/etc", "/sys",
    ]
    
    # 文件大小限制
    MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB
    
    # 表格大小限制
    MAX_TABLE_ROWS = 1000
    MAX_TABLE_COLS = 100
    
    # 文本长度限制
    MAX_TEXT_LENGTH = 100000  # 10万字符
    
    # 操作超时
    OPERATION_TIMEOUT = 60  # 秒
    
    # 允许的文件扩展名
    ALLOWED_EXTENSIONS = {'.docx', '.xlsx', '.pptx', '.pdf'}
```

**动态自定义目录处理**：

白名单采用**根目录前缀匹配**（`Path.is_relative_to`），允许目录下的所有子目录与文件，因此：

- 用户只需把自定义项目的**根目录**配置进白名单（`OFFICE_ALLOWED_DIRS` 逗号分隔多目录），该根下动态新建的任何子目录/文件自动合法，无需逐个配置或重启。
- 启动时的**工作目录（cwd）自动加入**白名单兜底：用户在任意项目目录下启动 MCP，即可操作该项目树内的一切文件。
- 白名单之外（如 `~/.ssh`、`/etc`）一律拒绝；**不提供运行时热更新白名单的工具**，新增目录需更新环境变量后重启 MCP，避免未授权越权访问。
- 客户端配置示例见 9.1 节。

### 7.5 HTTP/SSE 认证与速率限制

当通过 HTTP/SSE 传输部署时，必须启用认证和速率限制，防止未授权访问和滥用：

```python
class AuthMiddleware:
    """HTTP/SSE 传输认证中间件"""
    
    def __init__(self, api_keys: list[str]):
        self.api_keys = set(api_keys)
    
    def authenticate(self, headers: dict) -> bool:
        """验证 Bearer Token"""
        auth = headers.get("Authorization", "")
        if auth.startswith("Bearer "):
            token = auth[7:]
            return token in self.api_keys
        return False

class RateLimiter:
    """滑动窗口速率限制器"""
    
    def __init__(self, max_requests: int = 60, window_seconds: int = 60):
        self.max_requests = max_requests
        self.window = window_seconds
        self.requests = {}  # client_id -> [timestamps]
    
    def check(self, client_id: str) -> bool:
        """检查是否超过速率限制"""
        now = time.time()
        reqs = self.requests.get(client_id, [])
        # 清理过期记录
        reqs = [t for t in reqs if now - t < self.window]
        if len(reqs) >= self.max_requests:
            return False
        reqs.append(now)
        self.requests[client_id] = reqs
        return True
```

**认证配置**：

```python
# config.py 补充
class SecurityConfig:
    # HTTP/SSE 认证（stdio 模式无需配置）
    AUTH_ENABLED = os.environ.get("AUTH_ENABLED", "false").lower() == "true"
    API_KEYS = os.environ.get("API_KEYS", "").split(",")  # 逗号分隔的 API Key 列表
    
    # 速率限制
    RATE_LIMIT_ENABLED = os.environ.get("RATE_LIMIT_ENABLED", "true").lower() == "true"
    RATE_LIMIT_MAX = int(os.environ.get("RATE_LIMIT_MAX", "60"))  # 每分钟最大请求数
```

### 7.6 文件并发锁

防止多个 MCP 客户端或 Session 同时写同一文件导致数据损坏：

```python
import filelock

class FileLockManager:
    """文件级并发锁管理器"""
    
    def __init__(self, lock_dir: str = None):
        self.lock_dir = lock_dir or tempfile.gettempdir()
        self._locks = {}  # file_path -> FileLock
    
    def acquire(self, file_path: str, timeout: int = 30) -> filelock.FileLock:
        """获取文件锁"""
        lock_path = Path(self.lock_dir) / (Path(file_path).name + ".lock")
        lock = filelock.FileLock(str(lock_path), timeout=timeout)
        lock.acquire()
        return lock
    
    def release(self, lock: filelock.FileLock):
        """释放文件锁"""
        lock.release()

# 在写操作工具中使用
@mcp.tool
def word_add_paragraph(filename: str, text: str, ...):
    path = path_guard.validate_path(filename, "write")
    lock = file_lock_mgr.acquire(path)
    try:
        doc = Document(path)
        # ... 添加段落 ...
        doc.save(path)
    finally:
        file_lock_mgr.release(lock)
```

> **Session 模式下的锁**：Session 模式下文档在内存中编辑，不涉及磁盘并发写入，无需文件锁。锁仅在 `save_session` 写入磁盘时获取。

### 7.7 敏感数据脱敏

文档内容提取后返回给 LLM 时，对敏感信息进行可选脱敏处理：

```python
import re

class DataSanitizer:
    """敏感数据脱敏处理器"""
    
    PATTERNS = {
        "phone": (r'1[3-9]\d{9}', '1**********'),
        "email": (r'[\w.-]+@[\w.-]+\.\w+', '***@***.***'),
        "id_card": (r'\d{17}[\dXx]', '******************'),
        "bank_card": (r'\d{16,19}', '****************'),
    }
    
    def __init__(self, enabled: bool = False, fields: list[str] = None):
        self.enabled = enabled
        self.fields = fields or list(self.PATTERNS.keys())
    
    def sanitize(self, text: str) -> str:
        """脱敏处理"""
        if not self.enabled:
            return text
        for field in self.fields:
            if field in self.PATTERNS:
                pattern, replacement = self.PATTERNS[field]
                text = re.sub(pattern, replacement, text)
        return text

# 配置
class SecurityConfig:
    # 敏感数据脱敏（默认关闭，按需启用）
    SANITIZE_ENABLED = os.environ.get("SANITIZE_ENABLED", "false").lower() == "true"
    SANITIZE_FIELDS = os.environ.get("SANITIZE_FIELDS", "phone,email,id_card,bank_card").split(",")
```

---

## 8. 性能与可靠性

### 8.1 Session 管理（内存编辑）

借鉴 Aspose MCP Server 的 Session 管理思路，避免频繁磁盘读写：

```python
class SessionManager:
    """文档 Session 管理 — 内存中编辑，避免频繁磁盘 IO"""
    
    def __init__(self, ttl: int = 3600):
        self.sessions = {}  # session_id -> {doc, format, path, modified, created_at, last_access}
        self.ttl = ttl  # Session 过期时间（秒）
    
    def open_session(self, file_path: str, format: str) -> str:
        """打开文档到内存 Session"""
        session_id = str(uuid.uuid4())
        
        if format == "word":
            doc = Document(file_path)
        elif format == "excel":
            doc = load_workbook(file_path)
        elif format == "ppt":
            doc = Presentation(file_path)
        elif format == "pdf":
            # PDF 读取用 pypdf.PdfReader（支持后续写入），不用 pdfplumber（只读无 save）
            from pypdf import PdfReader
            doc = PdfReader(file_path)
        
        self.sessions[session_id] = {
            "doc": doc,
            "format": format,
            "path": file_path,
            "modified": False,
            "created_at": time.time(),
            "last_access": time.time(),
        }
        return session_id
    
    def save_session(self, session_id: str, output_path: str = None):
        """保存 Session 到磁盘"""
        session = self.sessions[session_id]
        save_path = output_path or session["path"]
        fmt = session["format"]
        doc = session["doc"]
        
        if fmt == "pdf":
            # PDF 不支持原地修改，使用 PdfWriter 写出
            from pypdf import PdfWriter
            writer = PdfWriter()
            for page in doc.pages:
                writer.add_page(page)
            with open(save_path, "wb") as f:
                writer.write(f)
        else:
            # Word/Excel/PPT 均支持 .save() 方法
            doc.save(save_path)
        
        session["modified"] = False
    
    def close_session(self, session_id: str, save: bool = False):
        """关闭 Session"""
        if save:
            self.save_session(session_id)
        # 清理资源（部分文档对象需要显式关闭）
        session = self.sessions.get(session_id)
        if session and hasattr(session["doc"], "close"):
            session["doc"].close()
        del self.sessions[session_id]
    
    def get_session(self, session_id: str):
        """获取 Session 中的文档对象"""
        session = self.sessions[session_id]
        session["last_access"] = time.time()  # 更新最后访问时间
        return session["doc"]
    
    def cleanup_expired(self):
        """清理过期的 Session（由后台定时任务调用）"""
        now = time.time()
        expired = [
            sid for sid, s in self.sessions.items()
            if now - s["last_access"] > self.ttl
        ]
        for sid in expired:
            self.close_session(sid, save=False)
        return len(expired)
```

**Session 使用模式**：

```
传统模式（每次操作读写磁盘）           Session 模式（内存编辑）
┌──────────────────────────┐       ┌──────────────────────────┐
│ 1. word_add_heading      │       │ 1. session_open(doc.docx)│
│    → 打开→修改→保存→关闭 │       │    → 加载到内存           │
│ 2. word_add_paragraph    │       │ 2. word_add_heading      │
│    → 打开→修改→保存→关闭 │       │    → 内存修改（无IO）     │
│ 3. word_add_table        │       │ 3. word_add_paragraph    │
│    → 打开→修改→保存→关闭 │       │    → 内存修改（无IO）     │
│                          │       │ 4. word_add_table        │
│ 磁盘IO: 6次（3读3写）    │       │    → 内存修改（无IO）     │
│ 耗时: ~3秒               │       │ 5. session_save          │
│                          │       │    → 一次写入磁盘         │
│                          │       │ 磁盘IO: 2次（1读1写）    │
│                          │       │ 耗时: ~0.5秒             │
└──────────────────────────┘       └──────────────────────────┘
```

### 8.2 大文件处理策略

| 格式 | 问题 | 策略 |
|------|------|------|
| Excel (>50MB) | openpyxl 全量加载导致 OOM | 使用 `read_only=True` 流式读取；写入使用 `write_only=True` |
| PDF (>100页) | 文本提取耗时 | 支持分页提取（page_range 参数）；设置超时 |
| Word (>500段) | 全文提取返回过大 | 分段返回；支持 offset/limit 分页 |
| PPT (>100页) | 结构分析耗时 | 支持指定 slide_range；摘要模式 |

### 8.3 错误处理策略

```python
# 统一错误处理
class ToolError(Exception):
    """工具错误 — 期望内的错误，对 LLM 可见"""
    pass

class InternalError(Exception):
    """内部错误 — 非期望错误，对 LLM 隐藏细节"""
    pass

# FastMCP 配置
mcp = FastMCP(
    "Office Document MCP",
    mask_error_details=True,  # 隐藏内部错误细节
)

# 工具中的错误处理模式
@mcp.tool
def word_add_table(filename: str, rows: int, cols: int, data: list):
    try:
        path = path_guard.validate_path(filename, "write")
        validator.validate_table_data(data, rows, cols)
        doc = Document(path)
        # ... 添加表格逻辑 ...
        doc.save(path)
        return f"成功在 {filename} 中添加 {rows}x{cols} 表格"
    except ToolError:
        raise  # 期望错误，直接抛出
    except Exception as e:
        logger.error(f"word_add_table 内部错误: {e}", exc_info=True)
        raise InternalError("处理文档时发生内部错误，请检查文件是否损坏")
```

### 8.4 超时与取消

```python
import threading
from contextlib import contextmanager

@contextmanager
def timeout_handler(seconds: int):
    """跨平台操作超时处理器（支持 Win/Mac/Linux）
    
    使用 threading.Timer 替代 signal.SIGALRM（后者仅限 Unix）。
    超时后通过抛出异常中断工作线程。
    """
    timer = threading.Timer(
        seconds,
        lambda: (_ for _ in ()).throw(
            ToolError(f"操作超时（{seconds}秒），请尝试处理更小的文件或范围")
        )
    )
    timer.daemon = True
    timer.start()
    try:
        yield
    finally:
        timer.cancel()

# 在耗时工具中使用
@mcp.tool
def pdf_extract_text(filename: str, page_range: str = None):
    with timeout_handler(config.OPERATION_TIMEOUT):
        # ... 提取逻辑 ...
```

### 8.5 监控与可观测性

提供运行时指标收集和健康检查，便于运维监控：

```python
class MetricsCollector:
    """运行时指标收集器"""
    
    def __init__(self):
        self.counters = {}   # tool_name -> 调用次数
        self.latencies = {}  # tool_name -> [耗时列表]
        self.errors = {}     # tool_name -> 错误次数
    
    def record(self, tool_name: str, duration_ms: int, success: bool):
        """记录一次工具调用"""
        self.counters[tool_name] = self.counters.get(tool_name, 0) + 1
        self.latencies.setdefault(tool_name, []).append(duration_ms)
        if not success:
            self.errors[tool_name] = self.errors.get(tool_name, 0) + 1
    
    def get_summary(self) -> dict:
        """获取指标摘要"""
        return {
            name: {
                "calls": self.counters.get(name, 0),
                "errors": self.errors.get(name, 0),
                "avg_latency_ms": sum(lats) / len(lats) if lats else 0,
            }
            for name in self.counters
        }

# 健康检查端点（HTTP 模式）
@mcp.tool
def doc_health_check() -> dict:
    """健康检查 - 返回服务状态和指标摘要"""
    return {
        "status": "healthy",
        "active_sessions": len(session_mgr.sessions),
        "metrics": metrics.get_summary(),
    }
```

**监控指标清单**：

| 指标 | 说明 | 告警阈值 |
|------|------|---------|
| 工具调用次数 | 每个工具的总调用次数 | - |
| 工具错误率 | 错误次数 / 总调用次数 | > 10% |
| 平均延迟 | 每个工具的平均执行耗时 | > 5秒 |
| 活跃 Session 数 | 当前打开的 Session 数量 | > 100 |
| 内存占用 | Python 进程内存使用 | > 80% |

---

## 9. 部署方案

### 9.1 本地部署（推荐）

适合个人用户和开发环境，通过 stdio 传输与 MCP 客户端通信：

```json
// ~/.config/claude/claude_desktop_config.json
{
  "mcpServers": {
    "timeverse-office-doc-mcp": {
      "command": "uvx",
      "args": ["timeverse-office-doc-mcp"],
      "env": {
        "OFFICE_ALLOWED_DIRS": "/path/to/custom_projects",
        "OFFICE_WORKSPACE": "/path/to/workspace",
        "OFFICE_OUTPUT": "/path/to/output",
        "OFFICE_TEMPLATES": "/path/to/templates"
      }
    }
  }
}
```

### 9.2 Docker 部署

适合团队共享和企业环境，通过 HTTP/SSE 传输：

```yaml
# docker-compose.yml
version: '3.8'
services:
  timeverse-office-doc-mcp:
    build: .
    ports:
      - "8000:8000"
    environment:
      - OFFICE_WORKSPACE=/data/workspace
      - OFFICE_OUTPUT=/data/output
      - MCP_TRANSPORT=http
      - MCP_PORT=8000
      - MAX_FILE_SIZE=104857600
      - AUTH_ENABLED=true
      - API_KEYS=your-secret-key-1,your-secret-key-2
      - RATE_LIMIT_ENABLED=true
      - RATE_LIMIT_MAX=60
      - SESSION_TTL=3600
      - SANITIZE_ENABLED=false
    volumes:
      - ./workspace:/data/workspace
      - ./output:/data/output
      - ./templates:/app/templates
    restart: unless-stopped
```

### 9.3 配置参数

```python
# 环境变量配置
class ServerConfig:
    # 传输方式: stdio | http | sse
    TRANSPORT = os.environ.get("MCP_TRANSPORT", "stdio")
    
    # HTTP/SSE 端口
    PORT = int(os.environ.get("MCP_PORT", "8000"))
    
    # 工作目录
    WORKSPACE_DIR = os.environ.get("OFFICE_WORKSPACE", "./workspace")
    OUTPUT_DIR = os.environ.get("OFFICE_OUTPUT", "./output")
    
    # 模板目录
    TEMPLATE_DIR = os.environ.get("OFFICE_TEMPLATES", "./templates")
    
    # 安全配置
    MAX_FILE_SIZE = int(os.environ.get("MAX_FILE_SIZE", str(100 * 1024 * 1024)))
    OPERATION_TIMEOUT = int(os.environ.get("OPERATION_TIMEOUT", "60"))
    
    # Session 配置
    SESSION_ENABLED = os.environ.get("SESSION_ENABLED", "true").lower() == "true"
    SESSION_TTL = int(os.environ.get("SESSION_TTL", "3600"))  # 1小时
    
    # 日志级别
    LOG_LEVEL = os.environ.get("LOG_LEVEL", "INFO")
    
    # 按需启用格式
    ENABLED_FORMATS = os.environ.get("ENABLED_FORMATS", "word,excel,ppt,pdf").split(",")
    
    # 可选: 启用 PyMuPDF 高性能模式
    USE_PYMUPDF = os.environ.get("USE_PYMUPDF", "false").lower() == "true"
    
    # HTTP/SSE 认证（stdio 模式无需配置）
    AUTH_ENABLED = os.environ.get("AUTH_ENABLED", "false").lower() == "true"
    API_KEYS = os.environ.get("API_KEYS", "").split(",") if AUTH_ENABLED else []
    
    # 速率限制
    RATE_LIMIT_ENABLED = os.environ.get("RATE_LIMIT_ENABLED", "true").lower() == "true"
    RATE_LIMIT_MAX = int(os.environ.get("RATE_LIMIT_MAX", "60"))  # 每分钟最大请求数
    
    # 敏感数据脱敏
    SANITIZE_ENABLED = os.environ.get("SANITIZE_ENABLED", "false").lower() == "true"
    SANITIZE_FIELDS = os.environ.get("SANITIZE_FIELDS", "phone,email,id_card,bank_card").split(",")
```

### 9.4 按需启用

支持按格式启用，减少资源占用（借鉴 Aspose MCP Server 的设计）：

```bash
# 仅启用 Word 和 Excel
ENABLED_FORMATS=word,excel uvx timeverse-office-doc-mcp

# 全部启用（默认）
uvx timeverse-office-doc-mcp

# 启用高性能 PDF 模式
USE_PYMUPDF=true ENABLED_FORMATS=pdf uvx timeverse-office-doc-mcp
```

---

## 10. 实施路线图

### 10.1 阶段划分

```
Phase 1: 基础框架 + 核心读写 (4周)
├── Week 1: 项目脚手架 + FastMCP 集成 + 安全层（PathGuard + 文件锁）
├── Week 2: Word 处理器（21个工具）
├── Week 3: Excel 处理器（24个工具）
└── Week 4: PPT 处理器（19个工具）

Phase 2: PDF 支持 + 分析能力 (3周)
├── Week 5: PDF 读取工具（pdfplumber + pypdf）+ OCR 工具
├── Week 6: PDF 写入工具（Overlay 合并模式: reportlab + pypdf）+ 安全工具
└── Week 7: 四格式分析工具完善

Phase 3: 高级特性 + 优化 (3周)
├── Week 8: Session 管理（4个工具 + TTL 清理）+ 大文件流式处理
├── Week 9: 资源(Resources) + 提示词(Prompts) + 模板系统（占位符引擎 + 边缘处理 + 四格式模板工具 + 模板注册管理）
└── Week 10: 性能优化 + HTTP/SSE 认证 + 速率限制 + 监控指标 + 测试覆盖 + 文档

Phase 4: 发布 + 生态 (2周)
├── Week 11: PyPI 发布 + Docker 镜像
└── Week 12: 文档完善 + 示例 + 社区推广
```

### 10.2 优先级矩阵

```
        高价值
          │
    ┌─────┼─────┐
    │ P0  │ P1  │
    │核心  │分析 │
    │读写  │能力 │
    │     │     │
 ───┼─────┼─────┼─── 难度
    │ P2  │ P3  │
    │模板  │Session│
    │系统  │管理  │
    │     │     │
    └─────┼─────┘
          │
        低价值

P0 (Phase 1-2): 必须有 — 核心读写工具
P1 (Phase 2):  差异化 — 分析能力
P2 (Phase 3):  核心特性 — 模板系统（四格式模板 + 占位符引擎）
P3 (Phase 3):  性能提升 — Session 管理
```

### 10.3 测试策略

| 测试类型 | 覆盖范围 | 工具 |
|----------|----------|------|
| 单元测试 | 每个工具的输入/输出/边界 | pytest |
| 集成测试 | 多工具串联工作流 | pytest + 真实文件 |
| 安全测试 | 路径越权、大文件、注入 | pytest + 安全用例 |
| 性能测试 | 大文件处理耗时 | pytest-benchmark |
| 兼容测试 | 多平台 (Win/Mac/Linux) | CI/CD 矩阵 |

---

## 11. 附录

### 11.1 竞品参考

| 项目 | 参考价值 |
|------|----------|
| GongRzhe/Office-Word-MCP-Server | Word 工具设计、模块化架构 |
| GongRzhe/Office-PowerPoint-MCP-Server | PPT 工具设计、模板系统 |
| haris-musa/excel-mcp-server | Excel 工具设计、双传输模式 |
| Aspose MCP Server | Session 管理、按需启用、认证机制 |
| theWDY/office-editor-mcp | 多格式合并方案、高级功能(OCR/比较) |
| ForLegalAI/mcp-ms-office-documents | Markdown→文档、模板系统、Docker部署 |

### 11.2 技术库版本参考

| 库 | 版本 | 许可证 | 说明 |
|----|------|--------|------|
| python-docx | 1.1+ | MIT | Word 读写 |
| openpyxl | 3.1+ | MIT | Excel 读写 |
| python-pptx | 1.0+ | MIT/BSD | PPT 读写 |
| pdfplumber | 0.11+ | MIT | PDF 文本/表格提取 |
| pypdf | 5.0+ | MIT | PDF 操作(合并/拆分/加密) |
| reportlab | 4.0+ | Commercial/Free | PDF 生成 |
| pandas | 2.2+ | BSD | 数据分析 |
| Pillow | 10.0+ | MIT-CMU | 图像处理 |
| FastMCP | 2.0+ | Apache-2.0 | MCP 框架 |
| PyMuPDF (可选) | 1.25+ | AGPL/Commercial | 高性能 PDF / OCR |
| filelock | 3.13+ | Unlicense | 文件并发锁 |

### 11.3 关键设计决策记录

| 决策 | 选择 | 替代方案 | 理由 |
|------|------|----------|------|
| 语言 | Python | Node.js/Go | 办公文档库生态最丰富 |
| MCP 框架 | FastMCP | 官方 SDK | 装饰器模式，开发效率高 |
| PDF 读取 | pdfplumber + pypdf | PyMuPDF | 避免 AGPL 许可风险 |
| PDF 写入 | reportlab + pypdf Overlay 合并 | PyMuPDF / pypdf 原地修改 | pypdf 不支持原地添加内容，reportlab 只能从零生成；Overlay 合并模式兼顾两者 |
| 架构 | 单一统一 Server | 多 Server 分离 | 用户配置简单 |
| 命名 | {format}_{verb}_{object} | 功能分组命名 | LM 理解更清晰 |
| Session | 可选启用 + TTL 自动清理 | 强制启用 / 无清理 | 兼顾简单场景和性能；防止内存泄漏 |
| 模板系统 | 统一占位符 `{{...}}` + 跨格式管理 | 各格式独立模板接口 | LLM 统一认知，降低调用复杂度 |
| PDF 模板 | AcroForm + reportlab 双模式 | 单一方案 | 覆盖表单填充 + 从零生成两种场景 |
| 超时实现 | threading.Timer | signal.SIGALRM | 跨平台支持（Win/Mac/Linux） |
| HTTP 认证 | Bearer Token | 无认证 / OAuth | 轻量级，适合 MCP 场景 |
| 文件并发 | filelock 文件锁 | 无锁 / DB 锁 | 防止多客户端同时写入损坏文件 |

### 11.4 与格式转换 MCP 的协作边界

```
本 MCP 的职责                      格式转换 MCP 的职责
┌──────────────────────┐          ┌──────────────────────┐
│ • Word 读/写/分析     │          │ • Word → PDF         │
│ • Excel 读/写/分析    │          │ • Excel → PDF        │
│ • PPT 读/写/分析      │  ←互补→  │ • PPT → PDF          │
│ • PDF 读/写/分析      │          │ • PDF → Word         │
│ • 格式内操作          │          │ • 格式间转换          │
│ • 文档内容提取        │          │ • 批量转换            │
│ • 结构分析            │          │ • 转换质量保证        │
└──────────────────────┘          └──────────────────────┘
```

**协作场景示例**：
1. 用户："读取这个 Word 文档的内容"→ 本 MCP（word_get_text）
2. 用户："将这个 Word 文档转为 PDF"→ 转换 MCP
3. 用户："分析这个 PDF 的表格数据"→ 本 MCP（pdf_extract_tables）
4. 用户："把这三份 PDF 合并成一份"→ 本 MCP（pdf_merge）
5. 用户："把这份 Excel 数据导出为 PDF 报告"→ 转换 MCP（格式转换）+ 本 MCP（pdf_add_watermark 添加水印）

---

*方案版本: v1.2 | 撰写日期: 2026-07-29 | 更新日期: 2026-08-01（v1.2 修复 Session/PDF/安全/统计缺陷，新增 10 个工具至 93 个）*
