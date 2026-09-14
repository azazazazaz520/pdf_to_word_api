# 外部 PDF 转换测试报告

编制日期：2026-09-13
适用范围：`D:\pdf_validation`（PDF 转 Word 独立验证项目）
被测版本：路由文本判定重构之后的工作区版本

> **本报告的转换质量结论已被后续修复推翻。**
>
> 报告第 3 节记录的两个缺陷（控制字符、APP14 JPEG）修复后，测试暴露出第三个缺陷：整页贴图兜底以图像相似度为依据，把文字完整的页面替换成截图，且兜底结果未写回质量报告。因此本报告第 2、5 节的"质量门禁"与"可编辑页"数据均不可用——那些数字取自引擎自报，而引擎当时系统性地虚报 9—17 页。修复过程与实际数据见 [整页贴图兜底重构](page-image-fallback-rework.md)。
>
> 修复后同一批样本的实际可编辑页由 3 页升至 40 页，引擎自报与实际一致。本报告的素材清单、素材截取方式、耗时构成测量（第 1、4 节）与遗留缺陷（第 6 节）仍然有效。

本文记录对 6 份外部真实 PDF 的转换测试，覆盖效果与耗时。素材为公开来源下载，与本项目既有样本不重复，均取 10—20 页规模。

---

## 1. 素材

| 样本 | 页数 | 来源 | 版式特征 |
|---|---|---|---|
| Attention Is All You Need | 15 | arxiv.org/pdf/1706.03762 | 英文双栏论文，公式与图表混排 |
| NIST SP 800-53r5 | 20 | nvlpubs.nist.gov | 英文标准文档，密排条目式排版 |
| ResNet | 12 | arxiv.org/pdf/1512.03385 | 英文双栏论文，图表密集 |
| GPT-3 | 20 | arxiv.org/pdf/2005.14165 | 英文长论文，取前 10 页与后 10 页 |
| NASA 系统工程手册 | 20 | nasa.gov | 英文手册，页眉页脚与配图 |
| BERT | 16 | arxiv.org/pdf/1810.04805 | 英文单双栏混排 |

NIST、NASA、GPT-3 原始文档分别为 492、297、75 页，按前 10 页与后 10 页截取，以同时覆盖封面/目录与正文版式。截取使用 `pypdfium2` 的 `import_pages`。

## 2. 转换结果

| 样本 | 页 | 耗时(秒) | 状态 | 质量门禁 | 路由 | 可编辑页 | 覆盖率均值 | 匹配行 | bbox 中位(pt) | SSIM 均值 |
|---|---|---|---|---|---|---|---|---|---|---|
| Attention Is All You Need | 15 | 120.56 | succeeded | error | text 15 | 15 | — | — | — | — |
| NIST SP 800-53r5 | 20 | 68.38 | succeeded | failed | text 20 | 20 | 0.6637 | 269 | 58.232 | 0.6729 |
| ResNet | 12 | 49.44 | succeeded | passed | text 9 + page_image 3 | 9 | 0.5303 | 305 | 1.681 | 0.6784 |
| GPT-3 | 20 | 47.65 | succeeded | passed | text 18 + page_image 2 | 18 | 0.8542 | 651 | 1.542 | 0.7897 |
| NASA 系统工程手册 | 20 | 35.50 | succeeded | failed | text 15 + page_image 5 | 15 | 0.6112 | 501 | 17.067 | 0.6256 |
| BERT | 16 | 18.80 | succeeded | error | text 6 + page_image 10 | 6 | — | — | — | — |

合计 103 页、340.3 秒，平均每页 3.30 秒。6 份全部转换成功，可编辑页合计 83 页。

「覆盖率」「匹配行」「bbox」「SSIM」为空表示该样本的渲染校验未完成，相关指标无从计算，详见第 3 节。

## 3. 测试中发现并修复的两个缺陷

两份样本在首次测试时无法完成转换。两者都在导出环节因文本层或图像层的异常输入而中断，与路由判定无关。

### 3.1 文本层控制字符导致导出中断

**症状**：ResNet 失败，报 `ValueError: All strings must be XML compatible: Unicode or ASCII, no NULL bytes or control characters`，耗时 4.52 秒。

**原因**：ResNet 第 5 页公式区域含 `U+0014`、`U+0015` 两个 C0 控制字符。`src/layout/text.py` 的 `_extract_text_characters` 已过滤空串、`\r`、`\n`、`\t` 与 `U+FFFE`，未覆盖其余 C0 控制字符；这些字符进入 DOCX 的 XML 即触发写入库校验失败。XML 1.0 只接受 TAB、换行与回车三类控制字符。

**修复**：`src/layout/text.py` 新增 `_is_xml_control_character`，在同一过滤点跳过其他 C0 控制字符。

**验证**：ResNet 由 failed 变为 succeeded，耗时 49.44 秒。

### 3.2 APP14 JPEG 无法被 DOCX 写入库识别

**症状**：NASA 手册失败，报 `UnrecognizedImageError`，耗时 1.04 秒。

**原因**：`src/layout/images.py` 的 `_image_bytes_from_object` 对 `DCTDecode` 图像有一条原始字节直通路径，仅校验 JPEG 首尾标记（`FFD8` / `FFD9`）。而 DOCX 写入库只按两种签名识别 JPEG：

```text
docx.image.jpeg.Jfif   offset=6  sig=b'JFIF'
docx.image.jpeg.Exif   offset=6  sig=b'Exif'
```

NASA 手册的配图为 `FFD8 FFEE`（APP14/Adobe）开头且无 JFIF、Exif 段，直通后写入库无法识别格式，抛出异常。该类 JPEG 常见于 Adobe 系列产品导出的文档。

**修复**：`src/layout/images.py` 新增 `_has_docx_readable_jpeg_header`，直通前校验第 7—10 字节是否为 `JFIF` 或 `Exif`；不满足时改走既有的位图重编码路径。

**验证**：NASA 手册由 failed 变为 succeeded，耗时 35.50 秒。

## 4. 耗时构成

自动化环节（版面提取、IR 构建、DOCX 写出、字体度量）稳定在数秒量级，DOCX 写出在 15 页样本上仅 4.51 秒。耗时集中在两处与 Word 相关的环节：

| 样本 | 最耗时环节 | 秒 | 占比 |
|---|---|---|---|
| Attention | render_validation_failed | 110.54 | 92% |
| NIST | fidelity_auto_fallback_started | 28.19 | 41% |
| ResNet | fidelity_auto_fallback_started | 30.47 | 62% |
| GPT-3 | fidelity_auto_fallback_started | 28.25 | 59% |
| NASA | fidelity_auto_fallback_started | 17.54 | 49% |
| BERT | render_validation_failed | 4.94 | 26% |

**整页兜底评估（`fidelity_auto_fallback_started`）**：在 ResNet、GPT-3、NIST、NASA 四个样本上占 41%—62% 的耗时，用于对 SSIM 或覆盖率不达标的页面重建整页图像。

**Word 渲染校验**：Attention 与 BERT 的校验失败分别耗时 110.54 秒与 4.94 秒。Attention 的失败信息为 `returncode=0, 未能生成 PDF`，即 Word 进程正常退出但未产出文件，属隐性失败；该样本的可编辑页为 15/15、DOCX 写出仅 4.51 秒，转换本身正常完成。

字体度量校准在缓存命中后仍需 4—6 秒，是第三项固定开销。

## 5. 质量门禁判定

6 份样本中门禁通过 2 份，其余 4 份的判定与转换质量不一致：

| 样本 | 门禁 | 转换状态 | 可编辑页 | 说明 |
|---|---|---|---|---|
| GPT-3 | passed | succeeded | 18 / 20 | 覆盖率 0.8542 |
| ResNet | passed | succeeded | 9 / 12 | 覆盖率 0.5303 仍判通过 |
| NIST | failed | succeeded | 20 / 20 | 全覆盖可编辑却判失败 |
| NASA | failed | succeeded | 15 / 20 | 覆盖率 0.6112 |
| Attention | error | succeeded | 15 / 15 | 渲染校验失败 |
| BERT | error | succeeded | 6 / 16 | 渲染校验失败 |

以可编辑页与字符覆盖率衡量，NIST 的 20/20 页全可编辑、覆盖率 0.6637 高于通过门禁的 ResNet（0.5303），却判为失败。这与 [下一阶段规划](next-phase-plan.md) 第二节记录的缺陷一致：门禁目前只检查渲染回读成功、`page_delta` 与空白页，不覆盖 SSIM、覆盖率与 bbox 误差，因此判定结果与实际转换质量并不同向。

## 6. 遗留缺陷

### 6.1 BERT 的导出物含超限形状尺寸

BERT 首次测试即报 `WORD_OPEN_FAILED`。排查确认其 DOCX 包结构完整：ZIP 无损坏、`word/document.xml`（6.79 MB）可解析、无 XML 非法字符、`sectPr` 层级合法（body 1 个 + 段落属性 15 个）。异常在于形状尺寸：

| 文档 | `<wp:extent>` 数量 | cx 取值范围（EMU） |
|---|---|---|
| BERT | 3596 | 5080 — 1 796 527 316 |
| Attention | 13 | 2 779 776 — 7 772 400 |
| 计算机组成原理实验 | 26 | 7 560 056（恒定） |

A4 页宽约 7.56 × 10⁶ EMU，BERT 的形状宽高最大被写成 1.80 × 10⁹ EMU（约 47 米）。该值超出 Word 可接受范围，Word 因此拒绝打开文档。

同一文档在后续测试中门禁报 `error`、可编辑页仅 6/16，路由将 10 页判为 `page_image`。版面提取在该文档上产出退化包围盒，是本项缺陷的根因方向，尚未定位到具体代码位置。

### 6.2 渲染校验对部分文档失败

Attention 与 BERT 的 Word 渲染校验分别报 `returncode=0, 未能生成 PDF` 与 `WORD_OPEN_FAILED`。两次失败都发生在渲染环节而非转换环节，且 Attention 的失败耗时 110.54 秒。以隐藏方式运行的 Word 在打开或导出时可能存在无提示的失败路径，需要单独排查。
