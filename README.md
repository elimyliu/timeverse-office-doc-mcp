# timeverse-office-doc-mcp

[![Python](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/)
[![CI](https://github.com/timeverse/timeverse-office-doc-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/timeverse/timeverse-office-doc-mcp/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

面向 **Word / Excel / PowerPoint / PDF** 四大格式的全场景读写与分析 MCP Server。提供 **74 个工具**，覆盖文档创建、内容编辑、格式化、数据分析、模板管理、Session 内存编辑等全链路能力，让 AI 模型直接操控办公文档。

---

## 目录

- [工具简介](#工具简介)
- [核心特点和优势](#核心特点和优势)
- [详细的能力清单](#详细的能力清单)
  - [Word 工具集（19 个）](#word-工具集19-个)
  - [Excel 工具集（16 个）](#excel-工具集16-个)
  - [PowerPoint 工具集（15 个）](#powerpoint-工具集15-个)
  - [PDF 工具集（16 个）](#pdf-工具集16-个)
  - [跨格式工具集（8 个）](#跨格式工具集8-个)
- [架构概览](#架构概览)
- [安装方式](#安装方式)
  - [TimeVerse Studio](#timeverse-studio)
  - [其他 AI 客户端](#其他-ai-客户端)
- [使用示例](#使用示例)
- [环境变量](#环境变量)
- [路径沙箱](#路径沙箱)
- [开发](#开发)
- [License](#license)

---

## 工具简介

**timeverse-office-doc-mcp** 是一个基于 [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) 的办公文档处理服务器。它将 Word、Excel、PowerPoint、PDF 四大主流办公格式的读写、编辑、分析能力封装为 74 个标准化工具，任何支持 MCP 的 AI 客户端都可以直接调用。

无论是让 AI 帮你生成一份格式规范的 Word 报告、处理 Excel 数据统计与排序、制作 PPT 演示文稿，还是对 PDF 进行合并拆分、加水印、提取表格——这个 MCP Server 都能让 AI 直接完成，无需人工切换软件。

**一句话定位**：让 AI 成为你的全能文档助手。

---

## 核心特点和优势

### 1. 四大格式全覆盖，74 个工具

一个 Server 搞定 Word（19）、Excel（16）、PowerPoint（15）、PDF（16）以及跨格式模板与 Session 管理（8），无需为每种格式单独部署。

### 2. Session 内存编辑模式

所有编辑类工具均支持 `session_id` 参数：

- **传入 session_id**：文档对象常驻内存，连续编辑零磁盘 IO，完成后一次性保存
- **不传 session_id**：走传统的「打开 → 修改 → 保存」流程

对 LLM 完全透明，按需选择，大幅提升多步编辑场景的性能。

### 3. 模板引擎与变量填充

统一的 `{{variable}}` 占位符语法，跨 Word / Excel / PPT / PDF 四种格式通用。注册模板时自动扫描提取占位符并推断类型（text / date / number / table / image 等），AI 只需提供变量值即可批量生成文档。

**多章节扩页（PPT）**：`doc_apply_template` 支持 `sections` 参数，自动识别模板中的章节页/内容页原型并复制为每章节一组（封面/目录 + 章节×N + 结尾），一次调用即可生成完整演示文稿；`ppt_fill_variables` 支持对单页或全文档随时补充变量。占位符即使被 Word/PPT 拆分为多个 run 也能正确替换。

### 4. 路径沙箱安全机制

- 白名单根目录前缀匹配，子目录自动合法
- 黑名单拦截敏感路径（`.ssh`、`.aws`、`/etc` 等）
- 写操作校验文件扩展名白名单
- 文件大小限制（默认 100MB）
- 相对路径统一锚定到项目根目录（`OFFICE_BASE_DIR` 可覆盖），不依赖进程 cwd
- 启动时自动将 cwd 加入白名单

### 5. 文件并发锁

基于 `filelock` 的文件级锁机制，防止多进程/多请求并发写同一文件导致损坏。所有修改类工具在 handler 调用周期内整体持锁，「读-改-写」全程原子，避免并发读写相互覆盖导致数据丢失。

### 6. 审计日志

所有工具调用自动记录 JSON Lines 格式审计日志（`audit.log`），包含工具名、参数摘要、耗时、成功/失败状态，便于追溯与调试。

### 7. 输入校验

对文件名、Excel 区域引用、表格行列数、文本长度等进行严格校验，防止注入与资源耗尽。

---

## 详细的能力清单

### Word 工具集（19 个）

#### 文档管理（5 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `word_create_document` | 创建新文档（支持模板创建与变量填充） | `filename` | `title`, `author`, `template`, `variables`, `session_id` |
| `word_get_info` | 获取文档元信息（`detailed=true` 时附带结构分析：标题层级、样式分布、表格详情等） | `filename` | `detailed`(默认 false), `session_id` |
| `word_extract_text` | 提取文本（`extract_type`: text=全文 / outline=仅大纲 / all=两者） | `filename` | `extract_type`(默认 "text"), `include_tables`(默认 true), `session_id` |
| `word_list_documents` | 列出目录内所有 Word 文档 | `directory` | - |
| `word_copy_document` | 复制文档 | `source`, `destination` | - |

#### 内容编辑（8 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `word_add_heading` | 添加标题（Heading 样式，level=0 为 Title；支持对齐、字号覆盖与段前/段后间距） | `filename`, `text` | `level`(默认 1), `align`, `font_size`, `space_before`, `space_after`, `session_id` |
| `word_add_paragraph` | 添加段落（支持样式、字号、加粗、对齐与段前/段后间距；`page_break=true` 时同时插入分页符） | `filename` | `text`, `style`, `font_size`, `bold`, `align`, `space_before`, `space_after`, `page_break`(默认 false), `session_id` |
| `word_add_cover` | 添加居中版式封面（标题/副标题/作者/日期/组织），完成后自动分页 | `filename`, `title` | `subtitle`, `author`, `date`, `org`, `session_id` |
| `word_add_table` | 添加表格（支持数据填充与表头样式） | `filename`, `rows`, `cols` | `data`, `has_header`(默认 true), `session_id` |
| `word_add_image` | 插入图片 | `filename`, `image_path` | `width`(英寸), `session_id` |
| `word_add_list` | 添加列表（项目符号或编号） | `filename`, `items` | `list_style`(默认 "List Bullet"), `session_id` |
| `word_set_header_footer` | 设置页眉页脚（可选页码） | `filename` | `header_text`, `footer_text`, `include_page_num`, `session_id` |
| `word_generate_toc` | 生成目录（Table of Contents，需在 Word 中更新域） | `filename` | `max_level`(默认 3), `styles`, `session_id` |

#### 格式化与操作（5 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `word_format_text` | 格式化文本片段（加粗、斜体、颜色、字体） | `filename`, `paragraph_idx`, `start`, `end` | `bold`, `italic`, `color`, `font`, `session_id` |
| `word_format_table` | 格式化表格（表头加粗、底纹） | `filename`, `table_idx` | `border_style`, `header_row`, `shading`, `session_id` |
| `word_search_replace` | 搜索替换文档中的文本 | `filename`, `find_text`, `replace_text` | `session_id` |
| `word_delete_paragraph` | 删除指定段落 | `filename`, `paragraph_idx` | `session_id` |
| `word_create_style` | 创建自定义段落样式 | `filename`, `style_name` | `font`, `size`, `color`, `bold`, `session_id` |

#### 分析工具（1 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `word_extract_tables` | 提取所有表格数据（JSON 或 CSV） | `filename` | `format`(默认 "json"), `session_id` |

> `word_get_info(detailed=true)` 输出文档结构分析（标题层级、样式分布）；`word_extract_text` 提取全文并支持大纲模式；分页符通过 `word_add_paragraph(page_break=true)` 插入。

---

### Excel 工具集（16 个）

#### 工作簿管理（3 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `excel_create_workbook` | 创建新工作簿（支持模板创建与变量填充） | `filename` | `sheet_name`(默认 "Sheet"), `template`, `variables`, `session_id` |
| `excel_get_overview` | 获取工作簿概览（工作表列表及各表行列数） | `filename` | `session_id` |
| `excel_manage_sheet` | 管理工作表（action: add / delete / rename / copy） | `filename`, `action` | `sheet_name`, `new_name`, `source`, `target`, `session_id` |

#### 数据读写（5 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `excel_read` | 读取数据（`range_str` 为单格如 "A1" 返回单值，为区域如 "A1:C10" 返回二维数据） | `filename`, `sheet`, `range_str` | `session_id` |
| `excel_write` | 批量写入区域（data 为二维数组，单格传 `[[value]]`；`header_bold`/`header_bg_color` 可直接应用表头样式） | `filename`, `sheet`, `start_cell`, `data` | `header_bold`, `header_bg_color`, `session_id` |
| `excel_write_multi` | 一次向多个工作表批量写入数据（不存在的 sheet 自动创建，支持统一表头样式） | `filename`, `sheets` | `start_cell`, `header_bold`, `header_bg_color`, `session_id` |
| `excel_modify_row` | 插入或删除行（action: insert / delete） | `filename`, `sheet`, `row_idx` | `action`(默认 "insert"), `count`(默认 1), `session_id` |
| `excel_modify_column` | 插入或删除列（action: insert / delete） | `filename`, `sheet`, `col_idx` | `action`(默认 "insert"), `count`(默认 1), `session_id` |

#### 格式化与高级（6 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `excel_format_cell` | 格式化单元格（字体、颜色、对齐、边框；格式化后自动调整列宽） | `filename`, `sheet`, `range_str` | `font`, `bold`, `italic`, `font_size`, `font_color`, `bg_color`, `alignment`, `border_style`, `auto_fit`(默认 true), `session_id` |
| `excel_apply_formula` | 应用公式 | `filename`, `sheet`, `cell_ref`, `formula` | `session_id` |
| `excel_create_chart` | 创建图表（bar / line / pie） | `filename`, `sheet`, `chart_type`, `data_range` | `title`, `session_id` |
| `excel_freeze_panes` | 冻结窗格 | `filename`, `sheet`, `cell_ref` | `session_id` |
| `excel_sort_data` | 排序数据区域 | `filename`, `sheet`, `range_str`, `key_column` | `ascending`(默认 true), `session_id` |
| `excel_add_conditional_format` | 添加条件格式 | `filename`, `sheet`, `range_str`, `rule_type` | `criteria`, `format_color`, `session_id` |

#### 分析工具（2 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `excel_analyze_data` | 数据统计分析（描述统计、频次分布等） | `filename`, `sheet` | `range_str`, `session_id` |
| `excel_find_duplicates` | 查找重复数据 | `filename`, `sheet` | `columns`, `threshold`(默认 1), `session_id` |

> 工作表管理统一由 `excel_manage_sheet`（add / delete / rename / copy）完成；单元格与行列读写由 `excel_read`/`excel_write`/`excel_modify_row`/`excel_modify_column` 覆盖；工作簿信息通过 `excel_get_overview` 获取。

---

### PowerPoint 工具集（15 个）

#### 演示文稿管理（5 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `ppt_create_presentation` | 创建新演示文稿（支持模板创建与变量填充） | `filename` | `template`, `variables`, `session_id` |
| `ppt_get_overview` | 获取演示文稿概览（尺寸、幻灯片数等；`list_slides=true` 时列出每张幻灯片概览） | `filename` | `list_slides`(默认 false), `session_id` |
| `ppt_add_slide` | 添加幻灯片 | `filename` | `layout`(默认 6), `title`, `session_id` |
| `ppt_manage_slide` | 管理幻灯片（action: delete / move / copy） | `filename`, `action`, `slide_idx` | `new_idx`(move 时必填), `session_id` |
| `ppt_apply_theme` | 应用主题配色（blue / green / orange / dark） | `filename`, `theme_name` | `session_id` |

#### 内容编辑（8 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `ppt_add_text` | 添加文本框 | `filename`, `slide_idx`, `text` | `left`, `top`, `width`, `height`, `font_size`, `bold`, `session_id` |
| `ppt_add_image` | 插入图片 | `filename`, `slide_idx`, `image_path` | `left`, `top`, `width`, `height`, `session_id` |
| `ppt_add_table` | 添加表格 | `filename`, `slide_idx`, `rows`, `cols` | `data`, `left`, `top`, `width`, `height`, `session_id` |
| `ppt_add_chart` | 添加图表（bar / line / pie） | `filename`, `slide_idx`, `chart_type`, `data` | `title`, `session_id` |
| `ppt_add_shape` | 添加形状（rectangle / oval / triangle 等） | `filename`, `slide_idx`, `shape_type` | `left`, `top`, `width`, `height`, `session_id` |
| `ppt_set_background` | 设置幻灯片背景色 | `filename`, `slide_idx` | `color`, `session_id` |
| `ppt_slide_notes` | 获取或设置演讲者备注（action: get / set） | `filename`, `slide_idx` | `action`(默认 "get"), `notes_text`, `session_id` |
| `ppt_fill_variables` | 填充 {{占位符}} 变量（slide_idx 指定则仅填充该页，否则全文档；复制模板页后补充不同内容） | `filename` | `variables`, `slide_idx`, `session_id` |

#### 分析工具（2 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `ppt_extract_text` | 提取所有幻灯片文本 | `filename` | `session_id` |
| `ppt_get_structure` | 获取完整结构树（`analyze=true` 时附带结构分析：形状/布局分布） | `filename` | `analyze`(默认 false), `session_id` |

> 演示文稿信息通过 `ppt_get_overview` 获取；幻灯片删除/移动/复制由 `ppt_manage_slide` 完成；备注读写由 `ppt_slide_notes` 覆盖；`ppt_get_structure(analyze=true)` 输出结构分析（形状/布局分布）；占位符填充由 `ppt_fill_variables` 按页或全文档完成。

---

### PDF 工具集（16 个）

#### 文档管理（4 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `pdf_get_info` | 获取元信息（页数、作者等；`analyze=true` 时附带页面结构分析） | `filename` | `analyze`(默认 false) |
| `pdf_merge` | 合并多个 PDF | `files`, `output` | - |
| `pdf_split` | 拆分 PDF（格式如 `1-3,4-6`） | `filename`, `page_ranges`, `output_prefix` | - |
| `pdf_rotate_page` | 旋转页面（90 / 180 / 270） | `filename`, `page_idx`, `angle` | - |

#### 内容读取（5 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `pdf_extract_text` | 提取文本 | `filename` | `page_range`, `layout_mode` |
| `pdf_extract_tables` | 提取表格数据 | `filename` | `page_range`, `format`(默认 "json") |
| `pdf_extract_images` | 提取图片 | `filename` | `page_range`, `output_dir` |
| `pdf_search_text` | 搜索文本 | `filename`, `query` | `case_sensitive` |
| `pdf_ocr_text` | OCR 识别（RapidOCR 引擎，纯 Python 无需系统依赖） | `filename` | `page_range`, `lang`, `output_format` |

#### 内容写入（4 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `pdf_add_text` | 添加文本（Overlay 模式） | `filename`, `page_idx`, `text` | `x`, `y`, `font`, `font_size`, `output` |
| `pdf_add_image` | 添加图片 | `filename`, `page_idx`, `image_path` | `x`, `y`, `width`, `height`, `output` |
| `pdf_add_watermark` | 添加水印（遍历所有页） | `filename`, `watermark_text` | `opacity`(默认 0.3), `font_size`(默认 60), `output` |
| `pdf_add_annotation` | 添加注释或书签（annotation_type: highlight / text / link / bookmark） | `filename`, `page_idx`, `annotation_type`, `content` | `x`, `y`, `output` |

#### 安全（1 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `pdf_manage_security` | 管理 PDF 安全（action: encrypt / decrypt） | `filename`, `action`, `password` | `permissions`, `output` |

#### 模板工具（2 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `pdf_create_from_template` | 从模板创建 PDF | `template_name` | `variables`, `output`(默认 "output.pdf") |
| `pdf_fill_form` | 填充 AcroForm 表单字段 | `filename`, `fields` | `flatten`(默认 true), `output` |

> PDF 加密/解密由 `pdf_manage_security`（action: encrypt / decrypt）完成；书签通过 `pdf_add_annotation(annotation_type="bookmark")` 添加；`pdf_get_info(analyze=true)` 输出文档结构分析；OCR 使用 RapidOCR 引擎（无需系统级依赖）。

---

### 跨格式工具集（8 个）

#### 模板管理（4 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `doc_list_templates` | 列出所有模板 | - | `format` |
| `doc_get_template_info` | 获取模板详情（含占位符列表） | `template_name` | - |
| `doc_manage_template` | 模板注册与删除（action: register / delete） | `action`, `name` | `format`, `file_path`, `description`, `placeholders` |
| `doc_apply_template` | 从模板创建文档并填充变量（PPT 模板支持 `sections` 按章节扩展页数） | `template_name`, `output_path` | `variables`, `sections` |

#### 占位符（1 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `doc_extract_placeholders` | 扫描提取模板占位符 | `template_name` | `format` |

#### Session 管理（3 个）

| 工具 | 描述 | 必需参数 | 可选参数 |
|------|------|----------|----------|
| `doc_open_session` | 打开文档到内存 Session | `filename`, `format` | - |
| `doc_close_session` | 关闭 Session（`save=true` 时先保存；`output_path` 指定保存位置） | `session_id` | `save`(默认 false), `output_path` |
| `doc_list_sessions` | 列出所有活跃 Session | - | - |

> 模板注册/删除由 `doc_manage_template` 完成；Session 内容保存通过 `doc_close_session(save=true, output_path=...)` 落盘。

---

## 架构概览

```
┌─────────────────────────────────────────────────────────┐
│                    AI 客户端 (MCP Client)                 │
│              TimeVerse Studio / Claude / Cursor ...       │
└────────────────────────┬────────────────────────────────┘
                         │ stdio (JSON-RPC)
┌────────────────────────▼────────────────────────────────┐
│                     MCP Server (stdio)                    │
│  ┌─────────────────────────────────────────────────────┐ │
│  │              server.py — 工具注册与路由               │ │
│  │   TOOL_DEFINITIONS (74 个工具 Schema)                │ │
│  │   TOOL_HANDLERS (工具名 → handler 映射)              │ │
│  └──────────────────────────┬──────────────────────────┘ │
│                             │                             │
│  ┌───────────┬───────────┬──┴────────┬──────────┬──────┐ │
│  │  Word     │  Excel    │  PPT      │  PDF     │ Doc  │ │
│  │  Handler  │  Handler  │  Handler  │  Handler │Handler│ │
│  │  (18)     │  (15)     │  (15)     │  (16)    │ (8)  │ │
│  └─────┬─────└─────┬─────└─────┬─────└─────┬────└──┬───┘ │
│        │           │           │           │       │     │
│  ┌─────▼───────────▼───────────▼───────────▼───────▼───┐ │
│  │                 公共服务层 (common/)                   │ │
│  │  ┌────────────┐ ┌──────────┐ ┌──────────┐           │ │
│  │  │ PathGuard  │ │Validator │ │ErrorHandler│          │ │
│  │  │ 路径沙箱    │ │ 输入校验  │ │ 错误处理   │          │ │
│  │  └────────────┘ └──────────┘ └──────────┘           │ │
│  │  ┌────────────┐ ┌──────────┐ ┌──────────┐           │ │
│  │  │ AuditLog   │ │ FileLock │ │ Session  │           │ │
│  │  │ 审计日志    │ │ 文件锁    │ │ 内存编辑  │          │ │
│  │  └────────────┘ └──────────┘ └──────────┘           │ │
│  │  ┌────────────────────┐                              │ │
│  │  │ TemplateManager    │                              │ │
│  │  │ 模板管理            │                              │ │
│  │  └────────────────────┘                              │ │
│  └──────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────┘
                         │
          ┌──────────────┼──────────────┐
          ▼              ▼              ▼
    ┌──────────┐  ┌──────────┐  ┌──────────┐
    │ workspace│  │ output   │  │ templates│
    │ 工作目录  │  │ 输出目录  │  │ 模板库    │
    └──────────┘  └──────────┘  └──────────┘
```

### 项目结构

```
timeverse-office-doc-mcp/
├── src/timeverse_office_doc_mcp/
│   ├── __init__.py              # 版本号
│   ├── server.py                # MCP Server 入口（74 个工具定义 + 路由）
│   ├── config.py                # 配置管理（ServerConfig + SecurityConfig）
│   ├── handlers/
│   │   ├── word_handler.py      # Word 19 个工具
│   │   ├── excel_handler.py     # Excel 16 个工具
│   │   ├── ppt_handler.py       # PowerPoint 15 个工具
│   │   ├── pdf_handler.py       # PDF 16 个工具
│   │   └── doc_handler.py       # 跨格式 8 个工具（模板 + Session）
│   └── common/
│       ├── path_guard.py        # 路径沙箱（白名单 + 黑名单 + 扩展名校验）
│       ├── validator.py         # 输入校验（文件名、区域、表格、文本长度）
│       ├── error_handler.py     # 错误处理（ToolError + 装饰器）
│       ├── audit_log.py         # 审计日志（JSON Lines 格式）
│       ├── file_lock.py         # 文件并发锁（filelock 实现）
│       ├── session.py           # Session 内存编辑（TTL 自动清理）
│       └── template_mgr.py      # 模板管理（注册/查找/占位符提取）
├── templates/                   # 文档模板库（registry.json 索引）
├── tests/                       # 测试（6 个文件，覆盖全部 handler）
├── .github/workflows/           # CI/CD（ci.yml + release.yml）
└── pyproject.toml
```

### 技术栈

| 层 | 技术 |
|----|------|
| MCP 协议 | `mcp` SDK（stdio 传输） |
| Word | `python-docx` |
| Excel | `openpyxl` + `pandas`（统计分析） |
| PowerPoint | `python-pptx` |
| PDF | `pdfplumber`（提取）+ `pypdf`（读写/合并/加密）+ `reportlab`（生成）+ `rapidocr-onnxruntime`（OCR） |
| 文件锁 | `filelock` |
| Python | >= 3.10，支持 3.10 / 3.11 / 3.12 |

---

## 安装方式

### TimeVerse Studio

[TimeVerse Studio](https://timeverse.ai) 是本项目的首选 AI 客户端，提供开箱即用的 MCP 集成体验。

#### 方式一：uvx 运行（推荐，无需预装）

在 TimeVerse Studio 的 MCP Server 配置中添加：

```json
{
  "mcpServers": {
    "timeverse-office-doc-mcp": {
      "command": "uvx",
      "args": ["timeverse-office-doc-mcp"],
      "env": {
        "OFFICE_ALLOWED_DIRS": "/path/to/your/documents",
        "OFFICE_WORKSPACE": "/path/to/workspace",
        "OFFICE_OUTPUT": "/path/to/output",
        "OFFICE_TEMPLATES": "/path/to/templates"
      }
    }
  }
}
```

> uvx 会自动创建隔离环境并拉取最新版本，无需手动安装 Python 包。

#### 方式二：直接安装后运行

```bash
pip install timeverse-office-doc-mcp
```

安装后可直接使用 `timeverse-office-doc-mcp` 命令。在 TimeVerse Studio 中配置：

```json
{
  "mcpServers": {
    "timeverse-office-doc-mcp": {
      "command": "timeverse-office-doc-mcp",
      "env": {
        "OFFICE_ALLOWED_DIRS": "/path/to/your/documents",
        "OFFICE_WORKSPACE": "/path/to/workspace",
        "OFFICE_OUTPUT": "/path/to/output",
        "OFFICE_TEMPLATES": "/path/to/templates"
      }
    }
  }
}
```

---

### 其他 AI 客户端

本项目兼容所有支持 MCP 协议的 AI 客户端，以下以 Claude Desktop 和 Cursor 为例。

#### Claude Desktop

编辑 Claude Desktop 配置文件（macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`）：

**uvx 方式：**

```json
{
  "mcpServers": {
    "timeverse-office-doc-mcp": {
      "command": "uvx",
      "args": ["timeverse-office-doc-mcp"],
      "env": {
        "OFFICE_ALLOWED_DIRS": "/path/to/your/documents"
      }
    }
  }
}
```

**直接安装方式：**

```json
{
  "mcpServers": {
    "timeverse-office-doc-mcp": {
      "command": "timeverse-office-doc-mcp",
      "env": {
        "OFFICE_ALLOWED_DIRS": "/path/to/your/documents"
      }
    }
  }
}
```

#### Cursor

在 Cursor 设置 → MCP 中添加新的 MCP Server：

**uvx 方式：**

```json
{
  "mcpServers": {
    "timeverse-office-doc-mcp": {
      "command": "uvx",
      "args": ["timeverse-office-doc-mcp"],
      "env": {
        "OFFICE_ALLOWED_DIRS": "/path/to/your/documents"
      }
    }
  }
}
```

**直接安装方式：**

```bash
pip install timeverse-office-doc-mcp
```

然后在 Cursor MCP 配置中填写命令为 `timeverse-office-doc-mcp`。

#### 从源码安装（开发者）

```bash
git clone https://github.com/timeverse/timeverse-office-doc-mcp.git
cd timeverse-office-doc-mcp
pip install -e ".[dev]"
```

开发模式下可使用以下命令直接启动：

```bash
python -m timeverse_office_doc_mcp.server
# 或
timeverse-office-doc-mcp
```

---

## 使用示例

### 示例 1：创建 Word 报告

对 AI 说：

> "帮我创建一份 Word 文档，标题是'2026 年度项目总结'，添加一级标题'项目概述'，下面写一段介绍文字，再添加一个 3 行 2 列的表格，表头是'指标'和'数值'。"

AI 会依次调用：

```
word_create_document(filename="2026年度项目总结.docx", title="2026 年度项目总结")
word_add_heading(filename="2026年度项目总结.docx", text="项目概述", level=1)
word_add_paragraph(filename="2026年度项目总结.docx", text="本项目旨在...")
word_add_table(filename="2026年度项目总结.docx", rows=3, cols=2, data=[["指标","数值"],["完成度","95%"],["预算执行","88%"]], has_header=true)
```

创建带封面（居中版式，自动分页）的文档：

```
word_add_cover(filename="2026年度项目总结.docx", title="2026年度项目总结", subtitle="技术部年度汇报", author="张三", date="2026-08-08", org="某科技有限公司")
```

### 示例 2：Session 模式连续编辑

对 AI 说：

> "打开 report.docx，先搜索替换把'旧公司名'改成'新公司名'，然后在末尾添加一个分页符，最后保存。"

AI 会使用 Session 模式减少磁盘 IO：

```
doc_open_session(filename="report.docx", format="word")
  → 返回 session_id: "a1b2c3d4e5f6"

word_search_replace(filename="report.docx", find_text="旧公司名", replace_text="新公司名", session_id="a1b2c3d4e5f6")
word_add_paragraph(filename="report.docx", page_break=true, session_id="a1b2c3d4e5f6")
doc_close_session(session_id="a1b2c3d4e5f6", save=true, output_path="report.docx")
```

### 示例 3：Excel 数据分析

对 AI 说：

> "读取 sales_data.xlsx 的 Sheet1，分析 A1:D100 区域的数据，然后按'地区'列升序排序，再查找重复的记录。"

AI 会调用：

```
excel_analyze_data(filename="sales_data.xlsx", sheet="Sheet1", range_str="A1:D100")
excel_sort_data(filename="sales_data.xlsx", sheet="Sheet1", range_str="A1:D100", key_column="地区", ascending=true)
excel_find_duplicates(filename="sales_data.xlsx", sheet="Sheet1", columns=["地区", "产品"])
```

### 示例 4：PDF 合并与加水印

对 AI 说：

> "把 report1.pdf 和 report2.pdf 合并成一个文件，然后加上'机密'水印。"

AI 会调用：

```
pdf_merge(files=["report1.pdf", "report2.pdf"], output="merged_report.pdf")
pdf_add_watermark(filename="merged_report.pdf", watermark_text="机密", opacity=0.3, font_size=60, output="merged_report_watermarked.pdf")
```

### 示例 5：模板填充生成文档

对 AI 说：

> "用'月度报告'模板生成一份文档，变量填充：项目名='Apollo'，日期='2026-08'，负责人='张三'。"

AI 会调用：

```
doc_apply_template(
    template_name="月度报告",
    output_path="output/Apollo_2026-08.docx",
    variables={"项目名": "Apollo", "日期": "2026-08", "负责人": "张三"}
)
```

### 示例 6：PowerPoint 快速制作

对 AI 说：

> "创建一个 PPT，第一页标题'产品发布'，添加第二页放一个柱状图，数据是 Q1:100、Q2:150、Q3:200。最后应用蓝色主题。"

AI 会调用：

```
ppt_create_presentation(filename="产品发布.pptx")
ppt_add_slide(filename="产品发布.pptx", layout=6, title="产品发布")
ppt_add_slide(filename="产品发布.pptx", layout=6, title="季度数据")
ppt_add_chart(
    filename="产品发布.pptx",
    slide_idx=1,
    chart_type="bar",
    data={"categories": ["Q1", "Q2", "Q3"], "series": {"销售额": [100, 150, 200]}},
    title="季度销售趋势"
)
ppt_apply_theme(filename="产品发布.pptx", theme_name="blue")
```

### 示例 7：模板多章节生成 PPT 演示文稿

对 AI 说：

> "用'银杉新经-商务模板'模板做一份公司介绍 PPT，封面标题'公司介绍'，目录四个章节：公司概况、主营业务、核心优势、联系我们，并给出每个章节的内容要点。"

AI 只需调用一次 `doc_apply_template`，通过 `sections` 自动将模板中的章节页/内容页原型复制为每章节一组（封面/目录 + 章节页×4 + 内容页×4 + 结尾页，共 11 页）：

```
doc_apply_template(
    template_name="银杉新经-商务模板",
    output_path="四川银杉新经科技有限公司-公司介绍.pptx",
    variables={
        "title": "公司介绍", "subtitle": "四川银杉新经科技有限公司",
        "date": "2026-08-08",
        "item1": "公司概况", "item2": "主营业务", "item3": "核心优势", "item4": "联系我们",
        "contact": "四川天府新区华阳街道华府大道一段1号1单元17层11号"
    },
    sections=[
        {"section_no": "01", "section_title": "公司概况", "slide_title": "公司概况",
         "point1": "成立于 2019 年 4 月", "point2": "注册资本 500 万元",
         "point3": "法定代表人刘欢", "point4": "统一社会信用代码 91510100MA66YJJB15"},
        {"section_no": "02", "section_title": "主营业务", "slide_title": "主营业务",
         "point1": "计算机软硬件开发", "point2": "数据处理及存储",
         "point3": "信息系统集成", "point4": "信息技术咨询"},
        {"section_no": "03", "section_title": "核心优势", "slide_title": "核心优势",
         "point1": "信用良好，自身/关联风险均为 0", "point2": "全链条 IT 服务能力",
         "point3": "小微轻量主体，决策灵活", "point4": "业务范围覆盖多元方向"},
        {"section_no": "04", "section_title": "联系我们", "slide_title": "联系我们",
         "point1": "公司全称：四川银杉新经科技有限公司", "point2": "注册地址：四川天府新区",
         "point3": "统一社会信用代码：91510100MA66YJJB15", "point4": "信息来源：国家企业信用信息公示系统"},
    ]
)
```

生成后如需微调某一页内容，可复制页面后用 `ppt_fill_variables` 按页补充变量：

```
ppt_manage_slide(filename="四川银杉新经科技有限公司-公司介绍.pptx", action="copy", slide_idx=7)
ppt_fill_variables(
    filename="四川银杉新经科技有限公司-公司介绍.pptx",
    slide_idx=7,
    variables={"slide_title": "新增章节", "point1": "要点一", "point2": "要点二", "point3": "要点三", "point4": "要点四"}
)
```

---

## 环境变量

本项目设计为零配置开箱即用，以下环境变量均为可选：

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `OFFICE_ALLOWED_DIRS` | - | 自定义白名单根目录（逗号分隔多个），默认已包含启动目录 |
| `OFFICE_BASE_DIR` | 项目根目录 | 相对路径解析基准目录（MCP 传入相对路径时统一锚定到此目录） |
| `OFFICE_TEMPLATES` | `./templates` | 模板库目录 |
| `SANITIZE_ENABLED` | `false` | 审计日志敏感数据脱敏开关 |
| `SANITIZE_FIELDS` | `phone,email,id_card,bank_card` | 脱敏字段类型 |

> **为什么不需要配置速率限制、超时、Session TTL 等？**
>
> 本项目基于 MCP stdio 协议，单客户端顺序通信，天然不存在并发争抢，无需速率限制。
> 文件大小、表格行列数、文本长度等安全限制已硬编码在代码中，无需用户调整。
> Session 过期时间（1 小时）作为内部常量管理，用户无需关心。
> 未来若增加 HTTP/SSE 多客户端传输，相关配置会随之引入。

---

## 路径沙箱

所有文件操作限制在白名单目录内，采用根目录前缀匹配：

- 配置的白名单目录（`OFFICE_ALLOWED_DIRS`）下的所有子目录/文件自动合法
- `workspace`、`output`、`templates` 三个内置目录自动加入白名单
- 启动时工作目录（cwd）自动加入白名单
- 相对路径统一锚定到项目根目录（`OFFICE_BASE_DIR` 可覆盖），解析结果可预测，不依赖进程 cwd
- 黑名单拦截敏感路径：`.ssh`、`.aws`、`.config`、`.env`、`AppData`、`Library`、`/etc`、`/sys`
- 写操作校验文件扩展名白名单：`.docx`、`.xlsx`、`.pptx`、`.pdf`
- 白名单之外一律拒绝

---

## 开发

```bash
# 安装开发依赖
pip install -e ".[dev]"

# 代码检查
ruff check src/ tests/
ruff format src/ tests/
mypy src/

# 运行测试
pytest -v
```

### CI/CD

- **ci.yml**：push/PR 到 main 时触发，lint（ruff + mypy）+ 矩阵测试（Ubuntu/macOS/Windows × Python 3.10/3.11/3.12）
- **release.yml**：打 `v*` tag 时触发，自动构建并发布到 PyPI

---

## License

[MIT](LICENSE)
