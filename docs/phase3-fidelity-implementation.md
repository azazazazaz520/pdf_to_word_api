# 高保真绝对定位导出（阶段 3）

## 1. 目标

放弃"流式重排 + 压缩字号"的排版方式，改为：

- 一页一个 Word section（画布），页面尺寸/页边距严格等于源 PDF；
- 文本、图片、线条、矢量对象、表格全部按源 PDF 坐标绝对定位；
- 无法可靠重建的区域渲染为原区域图片兜底，并写入质量报告；
- 导出后用 Microsoft Word 渲染回读，逐页对比 SSIM、文本 bbox 和字号。

## 2. IR 扩展

`src/document_ir.py`：

- 新增 `IRTextLine`：行级 bbox、字体名、字号、颜色、粗体/斜体、对齐、行距、
  首行缩进、旋转角、z-order、置信度。
- `IRBlock` 新增高保真字段：`font_name`、`font_size`、`color`、`bold`、
  `italic`、`alignment`、`line_spacing`、`first_line_indent`、`z_order`、
  `rotation`、`layer`、`lines`、`vector`、`latex`、`fallback_image`、
  `fallback_reason`。
- `IRPage` 新增 `reconstruction_confidence`、`fidelity`、`header_footer_native`，
  以及 `header_blocks` / `footer_blocks` / `body_blocks` / `fallback_blocks`
  等分层访问器。
- `IRDocument.fidelity_report()`：汇总逐页重建置信度、兜底区域、bbox/字号误差；
  `quality_report()` 内嵌 `fidelity` 字段。

## 3. 版面提取扩展（`src/pdf_layout.py`）

`extract_pdf_layout(..., include_fidelity=True)` 时额外提取：

- 文本对象级字体（family/base name）、字号、颜色、粗体/斜体、旋转角；
- 通过文本矩阵行列式修正被缩放的字体（例如 PDF 以 0.75 倍缩放绘制 14.72pt 文本，
  实际字号为 11.04pt），这是真实中文报告能对齐的关键；
- 行级对齐方式（left/center/right/justify）、行距倍数、首行缩进；
- z-order（页面对象绘制顺序）；
- 矢量对象：线条 / 矩形 / 路径，包含描边颜色、填充颜色、线宽、虚线、绘制模式；
- 表格内部的边框线段在提取阶段剔除，避免与 Word 表格边框重叠；
- 跨页重复的小图片标记为 logo，并按位置归入 header/footer 层；
- 页眉页脚行保留 `is_header_footer` 与 `layer` 信息，供 Word header/footer 导出。

## 4. 字体识别（`src/font_resolver.py`）

- **名称归一化**：去掉 `ABCDEF+` 子集前缀、`-0` 序号，按别名表映射到字体族
  （`SimSun`/`SimHei`/`Microsoft YaHei`/`DengXian`/`Times New Roman`/`Arial`…）；
- **样式判定**：`/FontWeight≥600`、`/Flags`（ForceBold / Italic 位）、
  `/ItalicAngle`、`/StemV` 启发式，加上字体名中的 Bold/Oblique 等词；
- **描述符来源**：pypdf 扫描页面资源字典，兼容 Type0 的 `/DescendantFonts`；
  CID 子集（如 `CIDFont+F1`）用 pdfium 的 family name 兜底；
- **本机匹配**：读取 Windows 字体注册表（含 `Cambria & Cambria Math` 这类组合名），
  未安装时按中英文回退链替换，并记录 `font_not_installed` 原因；
- **font run 切分**：同一行按（字体族、字号、颜色、粗体、斜体、旋转）切分成
  `PdfTextSpan`，导出时每个 run 独立 `w:framePr`，混排（中文 + 英文粗体）不再
  被行级"主字体"覆盖；
- **报告**：`quality_report()["fonts"]` 输出 `usage`（PDF 字体→Word 字体计数）
  与 `substituted`（替换字体及原因）。
- `tools/calibrate_font_offsets.py`：用 Word 实测各字体 framePr 偏移的标定工具。

## 5. 字体嵌入与度量标定（`src/font_embedding.py` / `src/font_metrics.py`）

**字体嵌入（ODTTF）**

- 从 PDF 字体描述符的 `/FontFile2`（Type0 走 `/DescendantFonts`）抽取 TrueType 子集；
- 按 ECMA-376 §17.8.1 取 GUID 十六进制字节**整体反转**作为密钥，前 32 字节异或两遍
  （用 Word 自己生成的嵌入字体反推校验）；
- 写入 `word/fonts/fontN.odttf`（content type `obfuscatedFont`）、`fontTable.xml`
  的 `w:embedRegular w:fontKey="{GUID}"`、`fontTable.xml.rels` 的 `.../font` 关系，
  并在 `settings.xml` 打开 `w:embedTrueTypeFonts`；
- 每个内嵌字体使用独立字体名（如 `SimSun DE8B28`），避免 Word 优先用系统同名字体；
  字体本身已带粗斜体字形，因此不再叠加 `w:b`/`w:i`；
- PDF 未内嵌的字体（如标准 Times/Arial）继续使用系统字体 + 粗斜体标记，并写入报告。

**度量标定**

- Word 在 EXACT 行距下，`framePr` 原点到字形框顶部存在与字号成正比的偏移 `k`，
  且不同字体差异很大（实测内嵌 SimSun `k≈-0.04`、SimHei `k≈-0.05`、
  系统 Times `k≈+0.10`）；
- 导出前生成一页标定样本（每种字体 × 5 个字号，CJK/拉丁分别取样），用 Word 渲染
  后测量 `k = (字形顶 - 框顶) / 字号`，导出时按 `frame_y = 字形顶 - k*字号` 对齐；
- 结果按「字体数据哈希 + 脚本类别」缓存到 `model_cache/font_metrics.json`。

**兜底策略调整**

- `ssim_threshold`（默认 0.98）只作为验收目标记录，不再触发贴图；
- 只有 `ssim_fallback_threshold`（默认 0.80）以下、或文本 bbox 误差 > 3pt 的页面
  才整页贴图，避免"字体已正确嵌入却仍输出全图片"。

## 6. OOXML 定位模块（`src/ooxml_positioning.py`）

集中封装绝对定位原语：

| 内容 | 实现 |
|---|---|
| 文本 | `w:framePr`（hAnchor/vAnchor=page，twips 坐标） |
| 复杂文本（旋转等） | `mc:AlternateContent` + `wps:wsp` 文本框，VML 兜底 |
| 图片 | `wp:anchor` + `wp:positionH/V` 页面偏移 |
| 线条/矩形 | `wps:wsp` 形状（rect/line），支持 flipH/flipV，VML 兜底 |
| 表格 | `w:tblpPr` 浮动定位 + `w:tblLayout fixed` + 精确列宽行高 |
| 页眉页脚 | 独立 header/footer part，定位方式与正文一致 |
| 页面尺寸 | 严格等于源 PDF（仅超过 Word 22in 上限时等比缩小并记录） |

## 7. 高保真导出（`src/pdf_to_word_exporter.py`）

新增 `export_fidelity_docx(ir, ...)`：

1. 每个 PDF 页对应一个 section，`pgSz` / `pgMar` 与源页一致（页边距为 0）；
2. 页面内容按 z-order 排序后逐个绝对定位：
   - 文本按行输出 `w:framePr`，字号为 `w:sz` 精确值，不做压缩；
   - 图片/文本框/矢量形状使用页面锚定；
   - 表格使用浮动表格 + 固定布局，单元格字号取自源单元格；
   - 公式优先 OMML，失败时贴公式区域图片；
   - 低置信度块、复杂矢量、缺失几何的表格直接贴区域图片；
3. 页眉页脚写入 Word header/footer，失败时回落到正文绝对定位；
4. 每个块记录 `native` / `image_fallback` / `skipped` 状态，页级重建置信度按
   面积加权计算（原生 1.0 / 图片兜底 0.8 / 跳过 0.0）；
5. 源空白页保留为空白 section，保证页数与源 PDF 完全一致。

文本定位补偿：Word 的 framePr 原点到字形框存在与字号成正比的系统偏移，
导出时按 `0.042*size+0.05`（x）与 `0.074*size+0.02`（y）补偿；实测
8–28pt 字号下残差 < 0.3pt。

## 8. 单一转换模式与自动兜底（`src/pdf_worker.py`）

程序只保留一种转换模式：`fidelity`。Worker 不再接收 `export_mode`，服务端
也不再提供 `output_mode` 参数；旧的流式排版导出（`export_ir_to_docx` 等）
只保留在导出模块中供基准脚本和历史测试使用，不参与服务转换。

转换流程固定为：

1. 版面提取（`include_fidelity=True`，含字体 run）；
2. 组装 IR（`fidelity=True`、`keep_header_footer=True`）；
3. `export_fidelity_docx` 绝对定位导出；
4. Word 渲染回读 + SSIM/文本 bbox 校验；
5. SSIM 不达标页自动整页图片兜底并复检（只复检兜底页）。

渲染回读与自动兜底：

1. 高保真导出后调用 Word COM 渲染回读；
2. 逐页计算 SSIM（默认 110 DPI，阈值 0.98）、逐行比对 bbox（< 3pt）
   与字号（< 0.5pt）；
3. 若存在 SSIM 不达标页面，且 `fidelity_auto_fallback` 开启，则把这些页面
   渲染为整页 PNG（默认与 SSIM 校验同 DPI）并重新导出、重新校验；
4. 兜底页写入 `fidelity_auto_fallback_pages`、`fallback_regions` 和
   `fidelity_page_image_fallback` 告警，便于人工复核。

## 9. 表格与矢量定位细节

- 表格使用 `w:tblpPr` 浮动定位：`w:tblpX = 表格 bbox 左边 + 源单元格左内边距`，
  `w:tblpY = 表格 bbox 顶边`；Word 会把左右内边距算在表格框之外、把上内边距
  算进行高，因此导出时分别做了补偿。
- 单元格文本位置来自 `PdfTableCell.text_bbox`（源单元格内文本行的并集），
  用它推算左右/上内边距，并把行的 `w:trHeight` 调整为"源行高 − 上内边距"，
  实测边框误差 < 0.6pt、单元格文本 bbox 误差约 2–3pt。
- 矢量对象：线条/矩形直接输出 `wps:wsp` 形状；多段路径按线段拆分；
  复杂路径（>24 段）自动贴区域图片。矩形/线条的 bbox 会按描边宽度内缩，
  因为 pdfium 的 `get_bounds` 已包含描边宽度。

## 10. 验收结果（本机）

合成样本（拉丁文本 + 标题 + 彩色斜体 + 线条 + 矩形 + logo + 页眉页脚）：

```text
页数：源 2 页，Word 渲染 2 页，page_delta = 0
空白页：0（非源空白页为 0）
SSIM：0.9876（阈值 0.98）
文本 bbox 误差：max 0.594pt，mean 0.316pt（阈值 3pt）
字号误差：max 0.04pt（阈值 0.5pt）
兜底区域：0
```

真实中文报告（4 页，嵌入式 DengXian 子集字体）：

```text
重建后 SSIM：0.79（字体轮廓与系统字体不一致，属于不可可靠重建）
自动兜底页：1-4（整页 PNG）
兜底后 SSIM：min 0.9932，mean 0.9973
页数：page_delta = 0，空白页 0
报告：4 个 fallback_regions，needs_review_pages = [1,2,3,4]
```

自动化测试（`python -m unittest discover -s tests -v`）：81 个测试全部通过，
其中包含 Word 渲染回读的端到端验收测试（缺少 Microsoft Word 时自动跳过）。

## 11. 相关测试

- `tests/test_ooxml_positioning.py`：定位原语与 OOXML 结构。
- `tests/test_fidelity_export.py`：页数/页面尺寸、bbox/字号、兜底区域、
  页眉页脚、空白页保留、矢量形状、Word 渲染验收。
- `tests/test_page_similarity.py`：SSIM 与逐页对比工具。
- `tests/test_document_ir.py`：高保真 IR 字段与 `fidelity_report()`。
- `tests/test_pdf_worker_fidelity.py`：单一 fidelity 模式与页面兜底。
- `tests/test_font_resolver.py`：字体名归一化、样式判定、本机匹配、字体使用报告、font run 切分。
- `tests/test_font_embedding.py`：ODTTF 密钥/混淆、字体抽取与嵌入计划、DOCX 部件写入、标定文档与偏移测量。

## 12. 已知边界

- Word 不支持嵌入 PDF 源字体，若系统字体与 PDF 内嵌子集轮廓差异较大，
  只能通过整页图片兜底保证视觉一致；
- `fidelity` 模式以"行"为单位输出段落，便于精确对齐，但可编辑性弱于
  `flow` 模式；需要强可编辑性时使用 `flow` 或 `text`；
- 单页文档仍沿用"只有页码才判定为页眉页脚"的保守规则，页眉文本会作为
  正文绝对定位输出（视觉位置不变）；
- 旋转文本使用文本框实现，文本框内行距由 Word 决定，位置残差略大于 framePr；
- 表格单元格文本依赖 Word 的单元格排版，行内垂直残差约 2–3pt；表格边框与
  列位置可以对齐到 0.6pt 以内；
- 含表格、CJK 嵌入字体或复杂图形的页面通常无法达到 SSIM 0.98，会触发整页
  图片兜底；`fidelity` 模式保证"页数一致 + 视觉一致"，可编辑性以报告中的
  `fallback_regions` 与 `needs_review_pages` 为准。
- 字体识别 + 内嵌 + 标定后，文本位置误差降到 0.1–0.6pt（SimSun），但源 PDF 的
  光栅化与 Word 的 PDF 导出仍存在字形抗锯齿/量化差异，整页 SSIM 通常在
  0.85–0.96，未必达到 0.98 的验收目标；这类页面会被记为"未达标"但仍保留可编辑文本。
- 少数页面（文本层缺失、整页图片、极端复杂版面）仍会走整页图片兜底。
