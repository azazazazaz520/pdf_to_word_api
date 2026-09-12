# Agent 交接文档：PDF → Word 高保真转换

> 目标读者：第一次接手本项目的 agent / 工程师。
> 读完本文即可运行、调试、扩展本转换器，不需要先读历史文档。

---

## 0. 一句话说明

把 PDF 逐页重建成 **Word（DOCX）**：一页一个 section、所有内容按源坐标**绝对定位**，
并使用 PDF 内嵌字体渲染；重建不了的区域自动贴原区域图片并写入质量报告。

仓库位置：`D:\pdf_validation`（独立验证项目，不属于 Prism 主工程）。
运行环境：Windows + Python 3.11（venv）+ **必须安装 Microsoft Word**（渲染回读用 COM）。

```powershell
cd D:\pdf_validation
$py = ".\.venv\Scripts\python.exe"     # 一定用 venv，不要用系统 Python
& $py -m unittest discover -s tests    # 81 个测试，约 13 秒
```

---

## 1. 快速上手

### 1.1 转换一份 PDF（最快路径）

把下面脚本存成 `tmp/convert.py`，改 `SOURCE` 后运行：

```python
import json, sys, time
from pathlib import Path

sys.path.insert(0, r"D:\pdf_validation")
from pypdf import PdfReader
from src.pdf_worker import process_job

SOURCE = Path(r"C:\path\to\input.pdf")
RUN = Path(r"D:\pdf_validation\tmp\my_run"); RUN.mkdir(parents=True, exist_ok=True)

payload = {
    "job_id": "my-run", "filename": SOURCE.name,
    "input_path": str(SOURCE), "output_path": str(RUN / "result.docx"),
    "progress_path": str(RUN / "progress.json"),
    "stage_log_path": str(RUN / "stages.jsonl"),
    "cancel_path": str(RUN / "cancel.requested"),
    "page_count": len(PdfReader(str(SOURCE)).pages),
    # 路由（逐页 auto / 强制 text / 强制 ocr）
    "route_mode": "auto", "engine": "structure-lite",
    "text_min_page_chars": 20, "text_min_page_ratio": 0.6,
    "text_high_quality_ratio": 0.8, "text_full_page_image_min_pixels": 300000,
    "text_garbled_char_ratio": 0.05,
    # 图片
    "page_image_max_pixels": 4194304, "page_image_jpeg_quality": 88,
    "embedded_image_max_pixels": 6000000, "embedded_image_jpeg_quality": 85,
    "embedded_image_png_optimize": False,
    # 渲染回读与验收
    "render_validation": True, "render_timeout_seconds": 900.0,
    "render_ssim": True, "render_ssim_dpi": 110.0,
    "render_text_compare": True, "render_coverage": True,
    "ssim_threshold": 0.98,             # 验收目标（只记录）
    "ssim_fallback_threshold": 0.80,    # 低于此值才整页贴图
    "coverage_fallback_threshold": 0.90,# 文字覆盖率低于此值才整页贴图
    "fidelity_auto_fallback": True,
    "fidelity_fallback_dpi": 110.0, "fidelity_page_image_max_pixels": 8000000,
    "fidelity_revalidate_after_fallback": True,
    # 字体
    "embed_pdf_fonts": True, "calibrate_font_metrics": True,
    "font_metrics_cache": str(RUN / "font_metrics_cache.json"),
    "include_toc": False, "include_bookmarks": True,
    "task_timeout_seconds": 1800.0, "ocr_time_budget_seconds": 600.0,
    "quality_gate_enabled": True,
}
started = time.perf_counter()
result = process_job(payload)
(RUN / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
print(result["status"], round(time.perf_counter() - started, 1), "s")
```

运行：`& .\.venv\Scripts\python.exe tmp\convert.py`

产物都在 `RUN` 目录：

| 文件 | 说明 |
|---|---|
| `result.docx` | 转换结果 |
| `result.rendered.pdf` | Word 渲染回读的 PDF（用于和源 PDF 逐页对比） |
| `result.json` | 完整返回，含 `quality` 质量报告 |
| `stages.jsonl` | 阶段日志（每阶段的耗时、SSIM、兜底页等，排障首选） |
| `progress.json` | 最新阶段（长任务可轮询） |

### 1.2 服务方式（FastAPI）

```powershell
$env:PDF_SERVICE_TOKEN = "test-token"
& .\.venv\Scripts\python.exe -m src.pdf_to_word_service
# POST /api/pdf-to-word/jobs (multipart: file)，GET /jobs/{id} 可查 quality
```

---

## 2. 架构与数据流

```
PDF
 └─ pdf_routing.analyze_pdf_text        逐页判定：text / page_image / ocr
     └─ pdf_layout.extract_pdf_layout(include_fidelity=True)
          · 文本行 + 字体 run（PdfTextSpan：字体/字号/颜色/粗斜体/z-order）
          · 表格（单元格 bbox/跨度/文本 bbox）、矢量对象、内嵌图片
          · 页眉页脚行、logo、多栏阅读顺序
         └─ ir_builder.build_document_ir        → Document IR（document_ir.py）
              └─ font_embedding.build_font_plan  → PDF 字体 → Word 字体 + ODTTF
                   └─ font_metrics.calibrate_font_metrics
                        → 每种字体的 framePr 垂直偏移 k（缓存）
                        └─ pdf_to_word_exporter.export_fidelity_docx
                             · 一页一 section（纸张=源 PDF，页边距 0）
                             · 文本逐 span w:framePr 绝对定位
                             · 图片/线条/表格/公式按 bbox 绝对定位
                             · 页眉页脚写入 Word header/footer
                             · 无法重建 → 贴区域图片
                              └─ docx_render_validation
                                   · Word COM 渲染回读 → page_delta / 空白页
                                   · 逐页 SSIM、字符 n-gram 覆盖率、行级 bbox/字号
                                    └─ 不达标页整页贴图后复检（只复检兜底页）
```

### 模块职责

| 文件 | 职责 |
|---|---|
| `src/pdf_routing.py` | 文本层完整性分析、逐页路由 |
| `src/pdf_layout.py` | 版面提取：文本行/字体 run/表格/矢量/图片/页眉页脚 |
| `src/font_resolver.py` | 字体识别：名称归一化、粗斜体判定、本机字体匹配与替换 |
| `src/font_embedding.py` | 抽取 PDF 内嵌字体 + ODTTF 混淆 + 写入 DOCX 包 |
| `src/font_metrics.py` | 用 Word 实测每种字体的 framePr 垂直偏移（带缓存） |
| `src/document_ir.py` | Document IR 数据结构与质量报告 |
| `src/ir_builder.py` | 版面/OCR → IR |
| `src/ooxml_positioning.py` | OOXML 绝对定位原语（framePr / 锚定图 / 形状 / 文本框 / 表格定位 / 页眉页脚） |
| `src/pdf_to_word_exporter.py` | `export_fidelity_docx()` 高保真导出（唯一导出路径） |
| `src/docx_render_validation.py` | Word 渲染回读、SSIM、覆盖率、bbox 对比 |
| `src/pdf_worker.py` | 任务编排（转换 + 校验 + 兜底） |
| `src/pdf_to_word_service.py` | FastAPI 服务与任务队列 |
| `tools/calibrate_font_offsets.py` | 独立字体度量标定工具（排查字体问题用） |

---

## 3. 当前转换模式（只有一个）

程序**只在转换时使用 `fidelity` 模式**，没有运行期模式选择：

- 旧流式导出（`export_ir_to_docx`、`export_text_pages_to_docx` 等）**已于 2026-09-12 整体移除**，`src/pdf_to_word_exporter.py` 现为 124 行的对外入口；
- Worker 不再读 `export_mode`，服务端不再有 `output_mode` / `PDF_SERVICE_EXPORT_MODE`；
- OCR 页 / 整页图片页统一按"整页图片 + 报告"处理。

---

## 4. 必须知道的坑（都是实际踩过并修复的）

1. **python-docx 的 `run.font.size` 会向下取整半磅**
   `Pt(10.45)` → `w:sz=20`（10.0pt，缩 4.3%）→ 整行变窄、bbox 误差暴涨。
   **必须**用 `ooxml_positioning.set_run_size_half_points(run, size)` 覆盖成 `round(size*2)`。

2. **ODTTF 混淆密钥是 GUID 十六进制字节「整体反转」**
   不是 `uuid.UUID(...).bytes_le`。参见 `font_embedding.guid_bytes`（测试里有 Word 反推的常量校验）。
   结构：`word/fonts/fontN.odttf` + `fontTable.xml` 的 `w:embedRegular`（`r:id` 来自
   `fontTable.xml.rels`，类型 `.../font`）+ `settings.xml` 的 `w:embedTrueTypeFonts`。

3. **Word 拒绝内嵌的符号字体子集**（Wingdings/Symbol/Webdings…）→ 字形消失。
   现在统一用系统字体（见 `font_embedding.SYMBOL_FONT_FAMILIES`）。

4. **Word 会自动在中文与西文之间加空格**（`w:autoSpaceDE/DN`），导致行宽和文本与源不一致。
   `add_absolute_text_paragraph` 里已显式关闭。

5. **源 PDF 的整页白色矩形会盖住 Word 页眉页脚**（页眉页脚绘制在正文层之下）。
   导出时跳过整页纯白填充；彩色背景则把页眉页脚留在正文层（见 `_page_background_blocks`）。

6. **每个 section 必须显式断开页眉页脚继承**，否则整页图片页会显示上一页的页码
   （`ooxml_positioning.detach_header_footer`）。

7. **绝对定位字号/坐标补偿**
   - 文本用 `w:framePr`（`hAnchor/vAnchor=page`，twips）；
   - 每个字体在 Word 里的"框顶→字形顶"偏移 `k` 不同（实测 SimSun≈-0.033、
     SimHei≈-0.052、系统 Times≈+0.104、Wingdings≈+0.062），必须用 `font_metrics` 标定；
   - `w:sz` 只能半磅 → 残余宽度差可用 `w:spacing` 补偿（`compensate_advance`，默认关闭，实测易过度补偿）。

8. **兜底不要滥用**：整页贴图在图片密集页并不比文本重建好（Word 导出会重编码图片，
   整页无损 PNG 兜底也只有 ~0.97）。现在只在
   `SSIM < 0.80` 或 `字符 n-gram 覆盖率 < 0.90` 时才贴图。

---

## 5. 质量报告怎么读

`result.json → quality`（也可从 `GET /api/pdf-to-word/jobs/{id}` 拿）：

| 字段 | 含义 |
|---|---|
| `route_summary` | 最终逐页路由统计（text / page_image / blank） |
| `page_delta` / `unexpected_blank_pages` | Word 渲染页数差、意外空白页 |
| `render_validation.ssim` | 逐页 SSIM：`min_ssim` / `mean_ssim` / `pages_below_threshold` |
| `render_validation.text_coverage` | 字符 n-gram 覆盖率：`min_coverage` / `pages_below_threshold` |
| `render_validation.text_layout` | 行级 bbox/字号误差：`median_bbox_error` / `p90_bbox_error` / `max_font_size_error` |
| `fonts.usage` | `pdf_font → word_font`（含粗斜体）使用计数 |
| `fonts.substituted` | 未安装/未嵌入而替换的字体及原因 |
| `fonts.embedding` | 内嵌字体数量、字节数、`entries`（哪些字体走内嵌、哪些走系统） |
| `fonts.metrics` | 各字体的标定偏移 `k` |
| `fidelity_auto_fallback_pages` | 触发整页贴图的页码 |
| `fidelity` | 逐页重建置信度、`fallback_regions`（含 bbox 与原因） |
| `fidelity_acceptance` | 验收结论：页数一致 / 无意外空白页 / SSIM / bbox / 字号 |

---

## 6. 当前验证状态（本机、CPU）

| 样本 | 页数 | 可编辑页 | 贴图页 | SSIM 均值/最低 | 说明 |
|---|---|---|---|---|---|
| `实验指导书.pdf`（STM32 教材前 20 页） | 20 | 17 | 3（本身是图片页） | 0.9603 / 0.8866 | 内嵌 SimSun/SimHei/Cambria |
| `F103电路板-用户手册.pdf` | 24 | 22 | 2 | 0.9605 / 0.9151 | 截图多，文本 bbox 中位数 ~1pt |
| `fixtures/*.pdf`（合成） | 1–2 | 1–2 | 0 | 0.98+ | 测试用，含表格/公式/图形 |

**尚未达到 0.98 验收目标的原因**：Word 导出 PDF 时会重编码图片（截图区域实测
0.95–0.99），这是管线天花板，与文本重建质量无关；文本位置中位数已在 1pt 量级。

---

## 7. 已知限制 / 待办

- [ ] 图片区域「局部贴图」替代整页贴图（图片密集页可减少不可编辑内容）；
- [ ] 提高截图类内嵌图片的有效分辨率，抵消 Word 的图片重编码；
- [ ] 逐行文本匹配对"一行拆成多个 span 帧"仍会报 `unmatched_line_count`（只影响报告，不影响渲染）；
- [ ] 更多泛化样本：扫描件（无文本层）、双栏论文、公式/复杂表格、旋转页面；
- [ ] 长文档性能：898 页全量约 15–20 分钟（含两次 Word 渲染），可考虑跳过兜底后复检。

---

## 8. 常用命令速查

```powershell
$py = ".\.venv\Scripts\python.exe"

# 全量测试
& $py -m unittest discover -s tests -v

# 单个模块测试
& $py -m unittest tests.test_font_embedding -v

# 字体度量标定（可选，排查字体偏移）
& $py tools\calibrate_font_offsets.py --output src\font_metrics_table.json

# 合成样本
& $py scripts\generate_fixture.py
& $py scripts\generate_table_fixture.py

# 单份 PDF 转换并计时（保真路径端到端，含渲染回读）
& $py tmp\run_and_time.py "C:\path\to\input.pdf" --run tmp\my_run
```

调试建议：

1. 先看 `stages.jsonl` 的 `font_plan_ready / font_metrics_ready / render_validation_completed / fidelity_auto_fallback_started`；
2. 用 `compare_pdf_pages` / `compare_text_coverage` / `compare_text_layout` 单独复算；
3. 怀疑字体问题：`extract_embedded_font_programs()` + `build_font_plan()` 打印计划，
   再用 `tools/calibrate_font_offsets.py` 看偏移是否异常。

---

## 9. 关键常量速查

| 常量 | 位置 | 值 |
|---|---|---|
| `DEFAULT_FIDELITY_FALLBACK_DPI` | `pdf_to_word_exporter.py` | 200 |
| `FIDELITY_TEXT_DX_FACTOR/BASE` | 同上 | 0.042 / 0.05 |
| `FIDELITY_TEXT_DY_FACTOR/BASE` | 同上（未标定字体的兜底） | 0.074 / 0.02 |
| `DEFAULT_OFFSET_FACTOR` | `font_metrics.py` | 0.074 |
| `OBFUSCATED_FONT_CONTENT_TYPE` | `font_embedding.py` | `application/vnd.openxmlformats-officedocument.obfuscatedFont` |
| 兜底阈值 | `pdf_worker.py` | SSIM < 0.80 或覆盖率 < 0.90 |
