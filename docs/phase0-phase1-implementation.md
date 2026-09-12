# 阶段 0 与阶段 1 实施记录

实施日期：2026-09-11

## 1. 阶段目标

- 阶段 0：统一 Document IR，补齐逐页路由、置信度、警告和质量回执。
- 阶段 1：逐页路由、同一 DOCX 混合文本/图片/OCR 页、文本路线保留页面内嵌图片、基础多栏阅读顺序、页眉页脚和页码过滤。

## 2. 已完成内容

### 2.1 统一 Document IR

新增 `src/document_ir.py`，提供：

- `IRWarning`：文档级和页级质量警告。
- `IRBlock`：标题、段落、列表、公式、表格、图片、页面保真图等统一块类型。
- `IRPage`：单页 route、尺寸、blocks、confidence、warnings、editable。
- `IRDocument`：文档级 blocks、metadata、质量报告。
- `IRDocument.quality_report()`：生成 `route_summary`、`page_results`、`needs_review_pages`、`table_count`、`media_count` 等字段。

新增 `src/ir_builder.py`，负责把文本布局、OCR 结果和页面保真图组装为 IR。

### 2.2 逐页路由

`src/pdf_worker.py` 已改为逐页路由：

- 每页独立判断 `text`、`page_image` 或 `ocr`。
- 文本层完整且质量高的页面走 `text`。
- 文本层不完整、全页图片或低质量页面走 `page_image`。
- 文档没有可用文本层时，整份文档走 `ocr`。
- 混合文档中无文本但有图片的页面走 `page_image`，避免因为单页图片而加载 OCR 模型。
- 同一 DOCX 可以同时包含文本页、页面保真页和 OCR 页。

### 2.3 页面内嵌图片保留

- `src/pdf_layout.py` 新增 `PdfImageBlock`，从 pdfium 页面对象提取图片、位置和 PNG 数据。
- 文本路线生成的 IR 会包含 `image` 块，导出时按原始位置插入 Word。
- 外部样本验证：`arxiv_attention_is_all_you_need.pdf` 转换后 `media_count = 3`，之前为 0。
- 外部样本验证：`who_phis_toolkit.pdf` 转换后 `media_count = 18`，并保持 18 个 Word 表格。

### 2.4 基础多栏阅读顺序

- `src/pdf_layout.py` 增加基于跨行稳定空白带的列边界检测。
- 跨栏标题作为全宽分隔行保留，栏内文本按列内顺序读取。
- 当前覆盖简单双栏和规则多栏；对异常 CMap、字符盒缺失的复杂学术 PDF 仍可能回退为单栏顺序。

### 2.5 页眉页脚和页码过滤

- 对跨页重复的顶部/底部文本进行归一化和计数。
- 多页文档中重复出现的内容被标记为 `is_header_footer`。
- 单页文档只过滤独立页码文本，避免误删标题。
- 文本 IR 不导出页眉页脚，质量报告中记录 `header_footer_filtered` 信息。

### 2.6 任务质量回执

`Job` 和任务记录新增：

- `quality`：完整质量报告对象。
- `route_mode`：任务级路由模式。
- `export_mode`：任务级导出模式。

`GET /api/pdf-to-word/jobs/{job_id}` 现在返回：

- `quality.route_summary`
- `quality.page_results`
- `quality.warnings`
- `quality.needs_review_pages`
- `quality.table_count`
- `quality.media_count`
- `quality.editable_page_count`
- `quality.visual_only_page_count`

质量报告同时写入 SQLite，服务重启后仍可恢复。

### 2.7 创建任务参数

`POST /api/pdf-to-word/jobs` 新增可选表单字段：

- `route_mode`：`auto`、`text`、`ocr`。
- `output_mode`：`text`、`editable`、`balanced`、`hybrid`。

`editable` 和 `balanced` 当前映射为可编辑文本导出；`hybrid` 在 OCR 页面前追加页面图像。

## 3. 质量和回归验证

当前自动化测试：

```text
Ran 35 tests in 5.683s
OK
```

新增测试覆盖：

- Document IR 质量报告统计。
- 双栏阅读顺序。
- 跨页页眉页脚过滤。
- 页面内嵌图片提取和坐标。
- 文本路线图片保留。
- 文本页 + 页面保真页混合导出。
- 文本页 + OCR 页混合导出。
- 服务重启后质量报告恢复。

更新后的文本线路快速路径已经通过现有测试和真实样本抽检。

## 4. 当前边界

- 页面保真页仍然是不可编辑图片，质量报告会将其标记为 `needs_review`。
- OCR 混合路径目前仍对整份 PDF 调用一次 OCR pipeline，只把 OCR 页结果写入 IR；后续应按 OCR 页渲染局部图片并只识别这些页。
- 无边框表格、复杂合并单元格、公式 OMML、书签、目录、页眉页脚写入 Word header/footer 仍属于阶段 2。
- 多栏检测对简单双栏有效，对异常字体和字符盒缺失的 PDF 仍需引入版面模型或更可靠的字符级定位。
- 页面数量变化还没有在导出后通过 Word/LibreOffice 渲染回读校验，`page_delta` 目前是占位值。

## 5. 下一阶段建议

1. 为 OCR 混合路径增加按页渲染和按页识别，避免整份 PDF 推理。
2. 增加无边框表格和复杂合并表格检测。
3. 增加公式 OMML 和公式图片兜底。
4. 将页眉页脚写入 Word 原生 header/footer，页码转 `PAGE` 域。
5. 将 PDF outline 映射为 Word 标题或书签。
6. 增加 DOCX 渲染回读，用于检测真实页数变化、空白页和内容溢出。

## 6. 外部样本抽检

使用新的逐页路由和 IR 导出路径抽检三个公开/现有样本：

| 样本 | 新结果 | 说明 |
|---|---|---|
| `weknora_paper_test_report.pdf` | 7 个文本页，9 个 Word 表格，0 个媒体对象，约 1.16 秒 | 保持原表格回归基线 |
| `who_phis_toolkit.pdf` | 27 个文本页 + 1 个页面保真页，18 个 Word 表格，18 个媒体对象，约 6.88 秒 | 不再把整份文件送入 OCR，图片和图表被保留 |
| `arxiv_attention_is_all_you_need.pdf` | 15 个文本页，3 个 Word 表格，3 个媒体对象，约 4.9 秒 | 页面内图片不再丢失；双栏阅读顺序仍受字符盒质量限制 |

这些结果来自本机 CPU 运行，只用于验证阶段性行为，不作为正式性能承诺。

## 7. 图片编码优化（P0/P1）

针对内嵌图片编码开销，已完成：

- PNG 默认不再使用 `optimize=True`；只有显式配置 `PDF_SERVICE_EMBEDDED_IMAGE_PNG_OPTIMIZE=true` 时才启用高成本压缩。
- 原始 JPEG 图片优先直通，避免解码后重新编码。
- 非 JPEG 图片先按 `PDF_SERVICE_EMBEDDED_IMAGE_MAX_PIXELS`（默认 6000000 像素）等比缩放，再选择 PNG 或 JPEG。
- 大尺寸、颜色丰富的照片类图片使用 JPEG，质量由 `PDF_SERVICE_EMBEDDED_IMAGE_JPEG_QUALITY` 控制（默认 85）。
- 截图、线条图、透明图片和颜色数量较少的图形继续使用 PNG。
- `extract_pdf_layout()` 支持 `include_page_images`，Worker 只提取 `text` 路线的页面图片；`page_image` 和 OCR 页面不再重复提取内嵌图片。

本机抽检 `who_phis_toolkit.pdf`：

```text
优化前：extract_pdf_layout 全量提取约 2.79 秒
优化后：全量提取约 0.86 秒
Worker 端到端：约 6.26 秒，输出约 1.54 MiB
```

优化后仍保留 18 个媒体对象，其中 17 个为文本页内嵌图片，1 个为页面保真图。

新增测试覆盖：

- 内嵌图片像素上限缩放。
- 原始 JPEG 流直通。
- 按页面范围过滤内嵌图片提取。

当前完整测试结果为 38 个测试通过。
