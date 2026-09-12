# 阶段 2 第二批实施记录

实施日期：2026-09-11

## 1. 本批目标

- 增加无边框表格识别，并输出真实 Word 表格。
- 增加复杂表格能力：重复列锚点、跨页续表判定继续复用现有逻辑。
- 增加公式 OMML 输出，失败时保留公式区域图片。
- 增加 PDF outline 提取，写入 Word 书签，并支持 Word TOC 域。

## 2. 无边框表格识别

实现位置：`src/pdf_layout.py`。

- 对不在矢量表格内的文本行按 y 坐标分组为行。
- 统计重复出现的 x0 列锚点，要求至少 3 列、至少 3 行。
- 每行至少匹配 2 个列锚点，并且连续行之间不能有明显的大间距。
- 根据列的实际 x0/x1 生成列边界，根据行 top/bottom 生成行边界。
- 生成 `PdfTable`，包含 `rows`、`cells` 和 `header_row_count`，导出层会写成真实 `w:tbl`。
- 普通双栏正文只有 2 个锚点，不会进入无边框表格，避免误判。

真实验证：

- `novus_tables.pdf` 之前输出为散乱项目符号文本，表格数为 0。
- 现在识别出 1 个真实 Word 表格。

## 3. 公式 OMML

新增 `src/formula_omml.py`。

实现链路：

1. 公式文本先归一化：Unicode 希腊字母、求和、积分、根号、上下标转 LaTeX。
2. `latex2mathml` 转 MathML。
3. 使用 Microsoft Word 自带的 `MML2OMML.XSL` 转 OMML。
4. 导出层把 OMML 写入 Word 段落。

覆盖的公式结构：

- 求和：`∑ᵢ xᵢ = 1` 生成 `m:nary`。
- 根式：`√x = 2` 生成 `m:rad`。
- 分式：`x = \frac{a}{b} + 1` 生成 `m:f`。

兜底策略：

- 如果 `MML2OMML.XSL` 不存在或转换失败，导出层使用公式区域图片。
- 图片区域通过 `pypdfium2` 从源 PDF 按公式 bbox 渲染。
- 如果没有源 PDF 或 bbox，最后退回普通斜体公式文本。
## 4. 书签与目录

新增 `src/pdf_outline.py`：

- 使用 `pypdf` 读取 PDF outline。
- 递归处理层级，返回 `title`、`page`、`level`。
- Worker 将 outline 写入 `IRDocument.metadata["outline"]`。
- 导出时：
  - 按 outline 页码在对应 Word 页面开头插入 `w:bookmarkStart` / `w:bookmarkEnd`。
  - 如果启用 TOC，则在文档开头插入 `目录` 标题和 Word `TOC` 域。
  - TOC 域使用 `TOC \o "1-3" \h \z \u`，用户在 Word 中更新域即可生成目录。

真实验证：

- `novus_bookmarked_toc_10_page.pdf` 输出 DOCX 中包含 `w:bookmarkStart` 和 `TOC` 域。

服务端配置：

- `PDF_SERVICE_INCLUDE_TOC`：默认 `false`。
- `PDF_SERVICE_INCLUDE_BOOKMARKS`：默认 `true`。

## 5. 测试

当前自动化测试：

```text
Ran 49 tests in 3.660s
OK
```

新增覆盖：

- LaTeX/Unicode 公式归一化。
- 公式 OMML 的 `m:f`、`m:nary`、`m:rad` 结构。
- OMML 写入 DOCX。
- 无边框表格识别，以及双栏正文不误判为表格。
- PDF outline 标题、页码和层级提取。
- Word 书签和 TOC 域写入。
- Worker 端到端写入 PDF outline 书签和 TOC 域。

## 6. 当前边界

- 无边框表格目前要求至少 3 列和 3 行；2 列无边框表格暂不覆盖。
- 无边框表格暂不推断合并单元格，只输出规则网格。
- 公式归一化主要覆盖常见数学符号和简单 LaTeX；复杂公式依赖 `latex2mathml` 的可解析范围。
- OMML 依赖本机 Microsoft Word 的 `MML2OMML.XSL`；没有 Word 时回退公式图片或普通文本。
- PDF outline 只能映射为书签和 TOC 域；还没有自动把 outline 标题匹配到正文标题。
- 跨页无边框表格的续表判定复用现有逻辑，尚未用真实样本专项验收。
- 完整 898 页实验指导书尚未在本批改动后重新验证无边框表格和公式 OMML 的数量与质量。
## 7. 898 页全量回归

`实验指导书.pdf` 在本批改动后的结果：

| 指标 | 结果 |
|---|---:|
| 总耗时 | 约 160.1 秒 |
| 输出 DOCX | 23,743,458 bytes |
| result.json | 约 21.9 KB |
| 路由 | text 891 / page_image 7 |
| Word 表格 | 316 |
| 媒体对象 | 711 |
| 公式块 | 130 |
| 代码块 | 1333 |
| 列表 | 1380 |
| 标题 | 616 |
| needs_review | 7 页 |
| OMML 段落 | 130 |
| Word 书签 | 1330 |

表格数从 143 增加到 316，主要来自无边框表格识别；抽样表格包括 BOOT0、资源、CPU 位数、引脚分类、总线名称和外设名称等，结构上符合原书内容。

本批公式样本以简单寄存器/数值表达式为主，转换后主要是 OMML 普通运行，没有出现 `m:nary`、`m:rad`、`m:f` 结构；复杂公式结构已由单元测试覆盖。

当前完整 898 页输出没有执行 Word 渲染回读，因此本批的 `page_delta` 仍为 0，需要后续单独跑 Word 渲染验证。
## 8. 898 页完整 Word 渲染回读

使用 Word COM 渲染完整 898 页输出并回读：

| 指标 | 结果 |
|---|---:|
| 源页数 | 898 |
| Word 实际渲染页数 | 1204 |
| page_delta | +306 |
| 空白页 | 6 页 |
| 空白页页码 | 236, 264, 618, 784, 925, 1038 |
| 渲染文件大小 | 约 35 MB |
| 回读耗时 | 约 69 秒 |

结论：

- 渲染回读链路可以处理完整 898 页文档，但 Word 真实排版与源 PDF 页数差异很大。
- 当前输出不能以页数不变作为质量合格标准。
- 后续需要继续优化代码块、列表、表格和图片的 Word 分页策略，并把 page_delta 作为质量门禁指标。
- 空白页需要通过分页控制和段落前后间距进一步排查。