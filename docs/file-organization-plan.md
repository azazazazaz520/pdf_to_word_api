# PDF 转 Word 项目文件整理方案

编制日期：2026-09-12
适用范围：`D:\pdf_validation`（PDF 转 Word 独立验证项目）
排序原则：先做不改变行为的清理，再做模块边界调整；每一步都可以单独回滚。

---

## 1. 现状规模与问题分级

| 文件 | 行数 | 问题 |
|---|---:|---|
| `src/pdf_to_word_exporter.py` | 3237 | 原 3579 行；已拆出 `src/export/text_utils.py`（89 行）与 `src/export/table_render.py`（274 行），仍包含三条导出路径，见第 1.3 节 |
| `src/pdf_layout.py` | 2468 | 文本、表格、矢量、图片、阅读顺序五个关注点混在一个文件 |
| `src/pdf_worker.py` | 1324 | 编排、OCR、导出、渲染回读、门禁五段逻辑混排 |
| `tests/test_flow_export.py` | 482 | 16 个用例混合覆盖旧流式导出与内容组装逻辑（`_group_text_lines` 等函数保真路径也在调用） |
| ~~`src/fidelity_writer.py`~~ | ~~181~~ | 死代码，已于 2026-09-12 删除（见第 1.1 节） |
| ~~`tmp/` 200 余文件~~ | —— | 一次性 `patch_*.py` / `fix_*.py` / 诊断脚本，已于 2026-09-12 清理（见第 1.1 节） |

按处理代价从低到高分成四类。

### 1.1 零风险删除 —— 已完成（2026-09-12）

| 项目 | 结果 |
|---|---|
| 删除 `src/fidelity_writer.py` | 已删除。删除前确认全仓库（`src`、`tests`、`scripts`、`tools`）零引用，仅自身第 73 行有 `extract_fidelity_pages` 定义；其职责（按区域裁掉源页文字再贴图）由 `pdf_to_word_exporter._region_image_bytes()` 承担 |
| 清理 `tmp/` 一次性脚本 | 已清理。710 个文件 / 737.1 MB → 32 个文件 / 213.9 MB，释放 523.2 MB；删除 678 个文件（294 个根目录散落脚本与输出、17 个 Word COM 渲染探针 `.ps1`、各阶段性探针目录） |
| 验证 | 96 个测试全部通过（19.1 秒）；`import src.pdf_to_word_service` 正常；`src/` 由 19 个模块减为 18 个 |

**归档路径**（均在仓库之外，不占用 `tmp/`、不会被提交，删除内容可从此恢复）：

| 归档文件 | 内容 | 校验 |
|---|---|---|
| `D:\pdf_validation_tmp_archive\tmp-20260912-pre-cleanup.tar` | 清理前的 `tmp/` 全量快照，737.7 MB，802 个条目 | 已用 `tar -tf` 验证可读 |
| `D:\pdf_validation_tmp_archive\src\fidelity_writer.py` | 已删除的死代码模块（181 行） | SHA256 `B1290926A9F19A3C523EAD6EF03F63B1A68A911DECDE4215799BC8EFDA816F48` |
| `D:\pdf_validation_tmp_archive\section140-sept-2026-09-12\` | 第 1.3 节分区整理的证据：整理前工作区备份 `pdf_to_word_exporter.py.bak`、归位脚本、AST 等价性校验脚本与其输出 | 校验输出记录"121 个定义无新增、无删除、无实现变化" |

清理时保留的三类内容：

1. `experiment_guide_test\实验指导书.pdf`、`f103_test\F103电路板-用户手册.pdf` —— 仅有的两个外部真实样本，无其他副本；
2. `experiment_guide_test\fidelity_first20_v5`、`f103_test\fidelity_run_v4` —— 最新一次运行的完整产物，即 S0 基线来源；
3. `experiment_guide_test\acceptance_final`、`experiment_guide_test\acceptance_bookmark_fix` —— 898 页最终验收产物，对应 `acceptance-898-report.md` 第 10 节结论，重跑需约 8 分钟且依赖 Word COM。

### 1.2 重复与命名不一致 —— 第 1、2 项已完成（2026-09-12）

1. **常量改名（已完成）。** 原 `MAX_WORD_PAGE_INCHES = 22.0`（`ooxml_positioning.py:29`）与 `MAX_PAGE_DIMENSION_INCHES = 11.0`（`pdf_to_word_exporter.py:148`）名字相近、含义不同，已分别改为：

   | 新名字 | 位置 | 含义 |
   |---|---|---|
   | `WORD_MAX_PAGE_INCHES` | `src/ooxml_positioning.py:29` | Word 允许的最大页面尺寸（22 英寸），超出时 `set_exact_page` 等比缩小并记录 `scaled` |
   | `PAGE_IMAGE_MAX_DIMENSION_INCHES` | `src/pdf_to_word_exporter.py:148` | `_scaled_page_size()` 允许的最大边长（11 英寸），超出后等比缩小 |

   两处均补充了说明用途的注释；旧名字已无残留引用。

2. **测试文件名对齐（已完成，名字与初版方案不同）。** `tests/test_pdf_to_word_exporter.py` 已改名为 `tests/test_flow_export.py`。初版方案建议的 `test_legacy_export.py` 与实际情况不符：该文件 16 个用例中只有一部分测旧流式导出函数，其余测的是保真路径同样依赖的内容组装逻辑——`_group_text_lines` 与 `_layout_lines_to_text` 虽然定义在导出模块里，但 `src/ir_builder.py` 与 `src/pdf_worker.py` 都在调用。命名为 `legacy_export` 会让人误判整个文件都是待删死代码。`test_flow_export.py` 只陈述实际被测试的导出方式，不隐含删除结论。

   后续若要进一步对齐，可把内容组装类用例拆到 `test_layout_content.py`，但需要连同 `_group_text_lines`、`_layout_lines_to_text` 等函数从导出模块迁出（属第 1.3 节导出模块拆分范围），届时一并处理。

3. **统一 IR 模块命名（未执行）。** `ir_builder` 与 `document_ir` 命名不成对，处理方式归入 S4 的 `ir` 包重构，本次不做。

### 1.3 超大文件拆分 —— 第 1 项的分区已落定（2026-09-12）

拆分需要先建立安全网（见第 4 节）。按"先文件后目录"原则，本轮先在原文件内落定分区与定义顺序，不做物理搬迁。

#### 1.3.1 `src/pdf_to_word_exporter.py` 分区结果

整理后文件 3579 行，分区如下：

| 分区 | 行范围 | 内容 |
|---|---|---|
| 文件头 | 1—79 | 模块 docstring（三条导出路径与分区说明）、import、模块级正则与常量 |
| 共享工具 | 80—566 | 段落/公式/代码识别、分页设置、文档样式、阶段回调、整页图片渲染 |
| 表格渲染 | 567—1435 | `_add_pdf_table` 等 PdfTable → Word 表格的构建，三条路径共用 |
| 流式导出 | 1436—1635 | 内容识别与分组工具 + 三个流式导出入口 |
| IR 流式导出 | 1636—2006 | 图片对齐、公式 OMML、书签、目录域 + `export_ir_to_docx` |
| 保真导出 | 2007—3292 | 专用常量、字体与坐标工具、逐块放置、页眉页脚、重建置信度 |
| 保真导出入口 | 3293—3579 | `export_fidelity_docx` 与 `_CanvasContext` |

#### 1.3.2 落定分区时修正的边界违规

`FIDELITY_MODES` 等 10 个保真专用常量与 `_fidelity_text_origin()` 原先位于文件头（旧 107—145 行），被流式共享区包围，但它们只被保真路径使用（调用点 2356、2454、2788、3174、3210 行均属保真分区）。已整体归位到保真分区，共移动 39 行。

顺带修正：`_looks_like_code_line` 前缺失的函数间空行。

#### 1.3.3 尚未解决的归属问题（物理搬迁前必须处理）

`_group_text_lines`、`_layout_lines_to_text`、`_layout_line_role`、`_text_line_role`、`_is_formula_line` 等内容识别与分组函数定义在流式导出分区内，但被 `src/ir_builder.py` 与 `src/pdf_worker.py` 调用，实测被"流式导出 + IR 流式导出 + 保真导出"三条路径共同使用。它们就是共享层的核心成员，物理搬迁时应迁入 `export/shared.py`。

搬迁前必须先解决顺序约束：保真分区依赖共享分区的工具函数，因此 `shared.py` 不得反向导入 `fidelity.py`。

#### 1.3.4 本轮的等价性验证方式

除 96 个测试外，新增 AST 级校验：对比整理前后每个顶层定义的实现体（忽略 docstring），确认 121 个定义中无新增、无删除、无实现变化，差异仅限定义顺序与注释。该方式可复用于后续任何搬迁。脚本、输出与整理前备份见第 1.1 节归档目录下的 `section140-sept-2026-09-12`。

#### 1.3.6 物理搬迁（已完成，2026-09-12）

搬迁后 `src/` 结构与规模：

| 模块 | 行数 | 内容 |
|---|---:|---|
| `src/export/text_utils.py` | 111 | 表格 HTML 解析、代码行识别、文本压缩与字段取值，仅依赖标准库 |
| `src/export/table_render.py` | 274 | PdfTable → Word 固定布局表格，含单元格取值与内边距 |
| `src/export/content.py` | 642 | 内容角色识别、行分组、列表与代码块、可编辑块写入 |
| `src/export/document_setup.py` | 345 | 文档样式、分页、整页与区域图片渲染、阶段回调 |
| `src/export/fidelity.py` | 1646 | 保真导出全部实现（专用常量、放置函数、页眉页脚、重建置信度） |
| `src/pdf_to_word_exporter.py` | 827（原 3579） | 对外入口：三条流式导出入口 + 从子模块精确转出全部符号 |
| `src/layout/models.py` | 360 | 版面数据契约：页面、文本行、表格、矢量、图片的表示 |
| `src/layout/text.py` | 525 | 字符与字体样式提取、文本行组装、font run 切分 |
| `src/layout/vectors.py` | 345 | 线条聚类、覆盖率、表格边界线、矩形与多段路径 |
| `src/layout/images.py` | 182 | 内嵌图片提取与编码（JPEG/PNG 选择、像素上限） |
| `src/layout/tables.py` | 717 | 有线表格、无边框表格、跨页续表与表头 |
| `src/layout/reading_order.py` | 298 | 多栏阅读顺序、页眉页脚、logo 判定 |
| `src/layout/layout.py` | 213 | `extract_pdf_layout` 入口与边框矢量过滤 |
| `src/pdf_layout.py` | 123（原 2468） | 对外入口：精确转出 95 个符号 |
| `src/worker/progress.py` | 99 | 取消与超时异常、阶段日志写入、OCR 流水线缓存 |
| `src/worker/quality.py` | 307 | 质量告警、门禁、渲染结果合并、保真验收、整页兜底 |
| `src/pdf_worker.py` | 1034（原 1324） | 任务编排入口 + `_prepare_fonts`、`_run_render_validation` |

依赖方向严格单向：

```text
export:  text_utils → table_render → content → document_setup → fidelity
layout:  models → text → vectors → images → tables → reading_order → layout
worker:  progress → quality
```

要点：

1. **对外接口零变化。** `src.pdf_to_word_exporter`、`src.pdf_layout`、`src.pdf_worker` 均可导入全部原有符号（含 `_plain_text`、`_add_pdf_table`、`_ProgressWriter`、`PdfPageLayout` 等下划线名字），`tests/`、`scripts/`、`tools/` 无需改动导入路径。
2. **依赖环以两种方式打破。** `render_pdf_region` 从导出主模块下沉到 `document_setup`；目录域、书签与公式 OMML 三个与流式导出共用的函数保留在主模块，由保真模块在函数体内延迟导入。
3. **顺带清理无用导入。** 导出主模块删除 `HTMLParser`、`WD_CELL_VERTICAL_ALIGNMENT`、`WD_ROW_HEIGHT_RULE`；`table_render` 删除 `WD_LINE_SPACING`。
4. **搬迁方式。** 由脚本按 AST 定义区间切分正文生成新模块，再删除原文件中同名定义并补写导入。切分必须以定义区间为准并保留 `@dataclass` 等装饰器，同时保证块内缩进按"原列号 − 父级列号"平移——本次过程中因手工改缩进与漏掉装饰器各返工一次，相关脚本与备份均已归档。

#### 1.3.7 搬迁后的验证结果

| 验证项 | 结果 |
|---|---|
| 96 个自动化测试 | 全部通过（18.7 秒） |
| 实验指导书前 20 页样本 | 与基线快照逐项一致：`page_delta=0`、`mean_ssim=0.9603`、`min_ssim=0.8866`、`matched_line_count=275`、`median_bbox_error=0.37` |
| F103 样本 | 与交接文档第 6 节记录吻合：`min_ssim=0.9151`（文档同为 0.9151）、`mean_ssim=0.9599`（文档 0.9605）、`page_delta=0`、`media_count=78` |
| 外部可导入符号 | 无缺失；`scripts/run_external_pdf_benchmark.py --help` 正常 |
| 全部模块导入 | 17 个 `src` 模块 + `tools/calibrate_font_offsets.py` 全部导入成功 |

**F103 的 `text_layout` 盲区已消失。** 搬迁后该样本 24 页全部参与布局比对（`matched_line_count=319`，中位数 0.983pt），而 1.3.3 节记录的旧行为是 0 行匹配、判定为"无有效数据"。这说明比对本身可用，此前的零匹配来自当时那次运行所用源文件的行切分方式，需在阶段 A 按第 1.3.3 节继续处理。

**关于 `tmp/baseline/f103-summary.json` 的可比性。** 该快照取自一次带整页兜底的早期运行（`route_summary` 为 text 22 / page_image 2、只比 2 页、`mean_ssim=0.9722`），与本次参数不同，因此不构成可比基线；F103 的可比基准以上表交接文档记录为准。实验指导书样本的快照参数一致，仍作为主要基线。
1. `quality.unexpected_blank_pages`：合并函数只写入 `quality.blank_pages`，该键在顶层报告里本就不存在（`render_validation.unexpected_blank_pages` 为 `[]`，`fidelity_acceptance.no_unexpected_blank_pages` 为 `true`）。基线值来自当时基线脚本自建的汇总字段，属报告 schema 的既有不一致，不是本次搬迁引入。
2. `table_count`、`media_count`：first20 基线汇总缺少这两个字段，本次脚本补齐，实测值 1 与 59。
3. `Arial` 标定偏移：本次首轮运行中 Arial 未出现在任何 span 中，故未纳入标定；`fonts.usage` 仍包含 Arial，属缓存内容差异，不影响渲染。
4. 字体标定来源由 `cache` 变为 `measured:5`：本次使用独立缓存路径，首轮重新实测；偏移数值与基线逐一相同（Wingdings 差 0.003，样本仅 42 处符号且不进正文可编辑文本）。

#### 1.3.8 第 1.3 节完成情况

| 原计划项 | 状态 |
|---|---|
| `src/pdf_to_word_exporter.py` 拆分 | 已完成，拆为 `src/export/` 五个模块，主模块由 3579 行降至 827 行 |
| `src/pdf_layout.py` 拆包 | 已完成，拆为 `src/layout/` 七个模块，主模块由 2468 行降至 123 行 |
| `src/pdf_worker.py` 拆分 | 已完成，抽入 `src/worker/` 两个模块并提取 `_prepare_fonts`、`_run_render_validation`，主模块由 1324 行降至 1034 行 |
| `process_job` 进一步缩小 | 未完成：仍为 636 行。剩余的 OCR 执行块与版面提取块已定位（见下），但连续两次提取都因手工调整缩进而损坏文件，最终按仓库"优先可恢复"的规则回退到已验证状态 |

后续若要继续缩小 `process_job`，做法已明确：把 `if ocr_indices:` 至 `ocr_budget_active = False`（约 119 行）抽为 `_run_ocr_pages`，把 `layout = None` 至版面提取结束（约 87 行）抽为 `_extract_layout_pages`。关键是**不要手工改缩进**：块的基准缩进应取块内非空行的最小缩进，函数体统一为列 4，嵌套层按原相对缩进平移；每次替换后立即用 `compile()` 校验，并确认新函数出现在 AST 顶层。

### 1.4 死代码与可提升资产 —— 已完成（2026-09-12）

1. **`src/ooxml_positioning.py` 的公开面与死代码。** 初版方案把 7 个"模块外零引用"的函数整体列为可疑对象，逐一核查后只有 1 个是真正的死代码：

   | 函数 | 定义行（核查时） | 模块内调用 | 其他 src | 测试/脚本 | 结论 |
   |---|---:|---:|---:|---:|---|
   | `insert_ordered` | 158 | 7 | 0 | 0 | 活跃，模块内部插入顺序原语 |
   | `clear_paragraph` | 243 | 2 | 0 | 0 | 活跃 |
   | `apply_absolute_frame` | 282 | 1 | 0 | 0 | 活跃，由 `add_absolute_text_paragraph` 调用 |
   | `set_run_size_half_points` | 303 | 1 | 0 | 0 | 活跃，交接文档第 4.1 节的关键修复点 |
   | `style_run` | 321 | 1 | 0 | 0 | 活跃 |
   | `apply_character_spacing` | 423 | 1 | 0 | 0 | 活跃 |
   | `set_table_grid` | 808 | **0** | **0** | **0** | **死代码，已删除** |

   已执行的动作：

   - **删除 `set_table_grid`**（原第 808—834 行，27 行）。全仓库 `git grep` 复核无残留；删除后模块由 879 行降至 864 行，25 个顶层定义。原功能由 `export/table_render.py` 的 `_set_fixed_table_layout` 与 `position_table` 共同承担。
   - **补写模块 docstring 的公开面说明**：列出导出方使用的 14 个公开原语（按尺寸颜色、文本、图形、表格、页眉页脚、页面承载分组），并写明其余下划线名字是模块内部工具，以及"判断函数能否删除时必须同时统计模块内调用，不能只看模块外引用"。

   同时更正初版方案的一处错误依据：`insert_ordered` 的 7 处调用**全部在 `ooxml_positioning` 模块内部**，并非"被 `pdf_to_word_exporter` 广泛使用"（该模块对它的引用数为 0）。

2. **探针脚本按需改写，不提升为正式资产。** 第 1.1 节归档的 5 个探针已取出复核：`font_metrics_probe.py`（解析 TTF 的 head/hhea/OS2 表）、`matrix_probe.py`、`drift_probe.py` 与两份 `offset_stats.py` 都是路径硬编码、绑定特定运行目录的一次性诊断脚本；`region_ssim.py` 的整页与文本行区域 SSIM 复算已被 `docx_render_validation.compare_pdf_pages()` 覆盖。

   结论：目前没有活跃的重复需求，不把它们改写成 `tools/` 命令或 `tests/` 断言——按"不因更完整而主动增加资产"的约束，缺少需求支撑的提升属于额外增加。其中唯一可能复用的是 `region_ssim.py` 的**逐文本行区域 SSIM + 最佳位移搜索**，记为阶段 A（分区域误差统计）的备选工具：届时若 `compare_pdf_pages()` 不足以定位区域级偏差，再正式落地到 `tools/`。提取命令见第 1.1 节归档路径。

3. **验证**：96 个测试全部通过（15.3 秒）；缺名检查 0 处；动态导入检查 44 个模块 0 问题；实验指导书前 20 页与 F103 的基线指标在删除前后一致。

---

### 1.5 S4 实施记录：建立 `src/ir/` 包（已完成，2026-09-12）

三个 IR 模块整文件移入子包，顶层保留同名门面文件精确转出公开符号：

| 原文件 | 行数 | 现位置 | 公开符号 |
|---|---:|---|---:|
| `src/document_ir.py` | 749 | `src/ir/model.py` | 7 |
| `src/ir_builder.py` | 738 | `src/ir/builder.py` | 4 |
| `src/pdf_outline.py` | 91 | `src/ir/outline.py` | 2 |
| 门面合计 | 13 + 11 + 9 = 33 | `src/document_ir.py`、`src/ir_builder.py`、`src/pdf_outline.py` | 13 |

搬迁中修正的两处导入关系：

1. **`ir_builder` 不再经由导出主模块取文本工具。** 它原先从 `pdf_to_word_exporter` 导入 `_field`、`_group_text_lines`、`_layout_lines_to_text`、`_line_is_in_pdf_table`、`_ordered_blocks`、`_plain_text`、`_typical_body_left`；这些名字定义在 `export/content.py` 与 `export/text_utils.py`，主模块只是转出。现改为直接导入定义处，依赖方向变为 `ir → export`，去掉一层中转。
2. **顶层私有名零外部引用的确认。** 三个模块共 27 个下划线开头的定义，全仓库无一处外部引用，因此门面文件只转出公开符号即可，不会破坏任何调用方。

#### 1.5.1 顺带消除两处导入期环

S4 期间用模块级导入图核查（`tmp/check_cycles3.py`），发现 `export/` 与导出主模块之间存在**三个导入期环**，全部由函数体内的延迟导入打破。其中两个的成因是 `_add_table` 放在了主模块：

| 环 | 处理 |
|---|---|
| `export.content → pdf_to_word_exporter → export.content` | **已消除**：`_add_table`（13 行）下沉到 `export/table_render.py`（它只依赖 `_table_rows` 与 `_set_table_cell_margins`），`content` 与主模块都改为模块级导入 |
| `export.content → … → export.fidelity → export.content` | 同上，随之消除 |
| `export.fidelity → pdf_to_word_exporter → export.fidelity` | **保留**：`_add_formula_omml_element`、`_add_toc_field`、`_add_bookmark`、`_bookmark_name` 四个工具同时被主模块的流式导出与保真导出使用，任一侧持有都会产生方向相反的边；现由保真模块在函数体内延迟导入，属有意为之，代码中有注释说明 |

`ir` 包在所有环之外，导入期无环。

#### 1.5.2 验证结果

| 验证项 | 结果 |
|---|---|
| 96 个自动化测试 | 全部通过（13.6 秒） |
| 缺名检查（28 个模块） | 0 处可疑名字 |
| 导入检查 | 10 个核心模块与三个门面全部导入成功，公开符号无缺失 |
| 实验指导书前 20 页 | 与基线逐项一致：`page_delta=0`、`mean_ssim=0.9603`、`min_ssim=0.8866`、匹配 275 行、bbox 中位数 0.37pt |
| F103 样本 | 与交接文档记录吻合：`min_ssim=0.9151`（文档同为 0.9151）、`mean_ssim=0.9599`（文档 0.9605）、`media_count=78` |

F103 与 `tmp/baseline/f103-summary.json` 的 71 处差异**全部源于该快照不可比**（其产生于带整页兜底、只比 2 页的早期运行）。差异中最关键的一项是度量口径的恢复：`fidelity_acceptance.bbox_ok` 由 `true` 变为 `false`、`max_bbox_error` 由 `0` 变为 `296.6pt`、`text_layout_available` 由 `false` 变为 `true`。旧记录里的"0 误差"是零匹配产生的假值，本次 319 行匹配才是真实测量——这正是第 1.3.3 节记录、阶段 A 需要处理的度量问题，属报告恢复真实结论，不是本次搬迁引入的回归。

---

### 1.6 S5 实施记录：建立 `src/validate/` 包（已完成，2026-09-12）

`docx_render_validation.py` 按职责拆为四个模块，顶层保留门面：

| 新模块 | 行数 | 内容 | 依赖 |
|---|---:|---|---|
| `src/validate/render.py` | 177 | Word COM 渲染、页数与空白页检查、PDF 页面渲染 | 仅标准库与 pypdf |
| `src/validate/similarity.py` | 140 | SSIM 计算与逐页视觉对比 | `render` |
| `src/validate/text_compare.py` | 282 | 字符覆盖率、行级 bbox 与字号比对 | 延迟依赖 `..pdf_layout` |
| `src/validate/report.py` | 94 | `validate_docx_rendering` 总入口与结果组装 | `render`、`similarity`、`text_compare` |
| `src/docx_render_validation.py` | 25（原 721） | 对外门面，转出 9 个公开符号 | 上述四个模块 |

外部导入面共 5 处（`font_metrics.py`、`pdf_worker.py` 与三个测试文件），全部无需改动。

#### 1.6.1 搬迁中修正的两处问题

1. **函数体内的相对导入层级未修正。** `_page_text_lines` 内有一处 `from .pdf_layout import ...`，拆包后应为 `from ..pdf_layout import ...`。顶层导入的层级修正并未覆盖函数体内的延迟导入，而它只在**运行期**才抛 `ModuleNotFoundError`，测试以 `text_layout.status == "failed"` 的形式暴露。已修正为 `..pdf_layout`。
2. **`report.py` 直接绑定函数引用，导致测试的补丁失效。** 测试以门面为补丁目标（`patch("src.docx_render_validation.render_docx_to_pdf")`），而 `report.py` 在模块级捕获了函数对象。现改为经 `_resolve(name)` 解析：先取门面上的同名属性（使补丁生效），未被打补丁时回落子模块。

#### 1.6.2 由此确立的两条硬性规则

后续每次搬迁都必须遵守，前两次（S4、S5）各违反一次：

1. **相对导入的层级修正必须覆盖函数体内的延迟导入。** 做法：搬迁后对每个新模块列出**全部** `ImportFrom`（含 `ast.walk` 到函数体内），确认层级正确；再运行 `tmp/check_imports_dynamic.py`（导入全部 44 个模块并逐个检查相对导入符号是否存在）作为门禁。只查顶层导入的静态检查会漏掉这一类。
2. **被测试打补丁的函数不要在新模块里提前绑定。** 若某模块被测试以门面路径打补丁，实现侧应通过 `_resolve(name)` 或模块属性访问取值，而不是 `from x import f` 后直接调用 `f`。

#### 1.6.3 验证结果

| 验证项 | 结果 |
|---|---|
| 96 个自动化测试 | 全部通过（16.0 秒） |
| 缺名检查（31 个模块） | 0 处可疑名字 |
| 动态导入检查（44 个模块） | 0 个问题：每个模块可导入，且其相对导入的符号都存在 |
| 导入期环 | 未新增：仍只有 `export.fidelity ↔ pdf_to_word_exporter` 一处（延迟导入打破） |
| 实验指导书前 20 页 | 与基线逐项一致：`page_delta=0`、`mean_ssim=0.9603`、`min_ssim=0.8866`、匹配 275 行、bbox 中位数 0.37pt |
| F103 样本 | 与交接文档吻合：`min_ssim=0.9151`（文档同为 0.9151）、`mean_ssim=0.9599`（文档 0.9605）、`media_count=78` |

### 1.7 S7 实施记录：建立 `src/fonts/` 与 `src/service/` 包（已完成，2026-09-12）

S7 覆盖三个目标：`worker/`（S7 早期已完成）、`fonts/`、`service/`。本次补齐后两者，阶段表所列的包已全部落地。

#### 1.7.1 `src/fonts/` 包

| 原文件 | 行数 | 现位置 | 门面 |
|---|---:|---|---:|
| `src/font_resolver.py` | 564 | `src/fonts/resolver.py` | 20 行 |
| `src/font_embedding.py` | 536 | `src/fonts/embedding.py` | 26 行 |
| `src/font_metrics.py` | 335 | `src/fonts/metrics.py` | 22 行 |

外部导入方共 6 处（`fonts/embedding`、`layout/models`、`layout/text`、`layout/layout`、`ir/model`、`export/fidelity`、`export/pdf_to_word_exporter`、`pdf_worker` 与两个测试文件），导入路径全部保留。

**关键约束已保持**：`fonts/metrics.py` 在函数体内延迟导入 `render_docx_to_pdf`（用 Word 实测标定）。这是 `fonts → validate` 的唯一入口，若改为顶层导入会形成 `validate → pdf_layout → layout → fonts` 的环。

#### 1.7.2 `src/service/` 包

| 原文件 | 行数 | 现位置 | 门面 |
|---|---:|---|---:|
| `src/pdf_to_word_service.py` | 1142 | `src/service/app.py` | 30 行（转出 21 个符号） |
| `src/job_store.py` | 158 | `src/service/jobs.py` | 11 行 |
| `src/ocr_quality.py` | 99 | `src/service/ocr_quality.py` | 14 行 |

顺带消除一处同名模块重复加载：`pdf_worker.py` 原先 `from .ocr_quality import ...`，与 `service/app.py` 使用的 `..service.ocr_quality` 会加载成两个模块对象（若其中一处含可变状态将产生隐蔽不一致）。现统一为 `from .service.ocr_quality import summarize_ocr_page`。

#### 1.7.3 本次暴露并修正的两类问题

1. **相对导入层级修正再次不完整。** `service/app.py` 与 `service/ocr_quality.py` 中还有 `from .pdf_worker`、`from .pdf_to_word_exporter`、`from .pdf_routing` 等根级导入需要上移一层。**这次是动态导入门禁（`tmp/check_imports_dynamic.py`）直接查出的**，5 个问题一次列全，未再依赖测试失败来发现。
2. **门面生成漏掉 `async def`。** 公共名提取只识别 `ast.FunctionDef`，导致 `create_job`、`lifespan` 两个异步接口未转出。已补入门面，并新增门面完整性核对（把每个门面转出的名字与其子模块的全部顶层定义逐一比对，10 组门面全部通过）。

#### 1.7.4 验证结果

| 验证项 | 结果 |
|---|---|
| 96 个自动化测试 | 全部通过（16.0 秒） |
| 动态导入检查 | **52 个模块，0 问题** |
| 缺名检查 | 0 处可疑名字 |
| 门面完整性核对 | 10 组门面全部覆盖其子模块的公开定义 |
| 导入期环 | 未新增：仍只有 `export.fidelity ↔ pdf_to_word_exporter` 一处 |
| 实验指导书前 20 页 | 与基线逐项一致：`page_delta=0`、`mean_ssim=0.9603`、`min_ssim=0.8866`、匹配 275 行、中位数 0.37pt |
| F103 样本 | 与交接文档吻合：`min_ssim=0.9151`、`mean_ssim=0.9599`、`media_count=78` |

### 1.8 S8 实施记录：删除旧流式导出路径（已完成，2026-09-12）

按选项 2 执行：删除旧流式导出实现及其专属测试。

#### 1.8.1 删除边界与删除量

| 文件 | 删除定义 | 行数变化 |
|---|---:|---|
| `src/pdf_to_word_exporter.py` | 16 | 816 → 124 |
| `src/export/content.py` | 7 | 640 → 365 |
| `src/export/document_setup.py` | 3 | 345 → 261 |
| `src/export/table_render.py` | 1 | 291 → 276 |
| `src/export/text_utils.py` | 2 | 111 → 77 |
| `tests/test_flow_export.py` → `test_content_grouping.py` | 12 个用例 | 482 → 66 |
| `tests/test_formula_omml.py` | 1 个用例 | 63 → 33 |
| `tests/test_pdf_outline.py` | 2 个用例 | 161 → 78 |
| `scripts/run_external_pdf_benchmark.py` | 整个文件 | 323 → 0（已归档） |

删除的 29 个定义：四个旧导出入口（`export_text_pages_to_docx`、`export_source_pages_to_docx`、`export_results_to_docx`、`export_ir_to_docx`）及其专属辅助——页面组装（`_add_layout_pages`、`_add_text_pages_with_sizes`、`_add_ocr_pages`、`_add_source_pages`、`_add_page_image`、`_iter_page_images`、`_pdf_page_sizes`、`EXPORT_MODES`）、IR 流式逐块写入（`_write_ir_block`、`_add_ir_image`、`_image_alignment`、`_add_formula_omml`、`_add_formula_image`、`_add_zero_height_bookmark_paragraph`、`_reset_list_state`）、可编辑块与编号（`_add_block`、`_add_content_lines`、`_add_layout_page_content`、`_add_editable_text_block`、`_add_code_block`、`_create_numbering_instance`、`_set_numbering_instance`）与 HTML 表格渲染（`_add_table`、`_table_rows`、`_TableParser`）。

`pdf_to_word_exporter.py` 现为 124 行，只保留：`render_page_image`（服务用于页面渲染）、`_add_formula_omml_element`、`_bookmark_name`、`_add_bookmark`、`_add_toc_field`（后四者由保真模块在运行期延迟导入），以及对外转出的 10 个符号。

#### 1.8.2 删除边界的两处误判与最终判据

删除集最初由静态可达性分析给出，**两次误判**，均由动态导入检查发现：

1. `_scaled_page_size`、`_set_section_page` 被判为死代码。实际 `export/fidelity.py` 从 `export/document_setup` 直接导入并使用它们。静态分析因为 `fidelity` 经由门面取用、名字未直接出现而漏判。
2. `PAGE_IMAGE_MAX_DIMENSION_INCHES` 被判为死代码。它是 `_scaled_page_size`（存活）使用的常量，随上一条一并误判。

**结论：删除边界的最终判据是动态导入检查（`tmp/check_imports_dynamic.py`：导入全部模块并逐个核对相对导入符号），静态可达性分析只能作为候选集生成的起点。** 这一条已与第 1.6.2 节的两条规则并列，作为后续删除类改动的门禁。

#### 1.8.3 保留与删除的取舍

- **保留 4 个内容组装用例**（`test_content_grouping.py`）：它们测的 `_group_text_lines`、`_layout_lines_to_text` 等函数由 `ir_builder` 在活路径上调用。判据是"用例触及的符号是否仍有存活使用者"，不是文件归属。
- **删除整个基准脚本**：`_docx_metrics` 与 `_source_metrics` 只依赖存活 API，但三个 `_convert_*` 全部基于被删的旧导出；保留骨架需重写为保真路径基准，而服务的 `process_job` + 质量报告已覆盖同一需求。原文件归档在 `D:\pdf_validation_tmp_archive\section140-sept-2026-09-12\run_external_pdf_benchmark.py.removed_in_s8`。

#### 1.8.4 验证结果

| 验证项 | 结果 |
|---|---|
| 自动化测试 | **81 个全部通过**（原 96 个，删除 15 个仅测旧路径的用例） |
| 动态导入检查 | 52 个模块，0 问题 |
| 缺名检查 | 0 处可疑名字 |
| 导入期环 | 未新增：仍只有 `export.fidelity ↔ pdf_to_word_exporter` 一处 |
| 实验指导书前 20 页 | 与基线逐项一致：`page_delta=0`、`mean_ssim=0.9603`、`min_ssim=0.8866`、匹配 275 行、中位数 0.37pt |
| F103 样本 | 与交接文档吻合：`min_ssim=0.9151`、`mean_ssim=0.9599`、`media_count=78` |

---

## 2. 目标目录结构

标注含义：**已完成**表示目录已按此落地；**待做**表示仍为实现目标。

```text
D:\pdf_validation
├─ src/
│  ├─ __init__.py
│  ├─ pdf_routing.py              页面路由（保留为单模块，未拆）
│  ├─ layout/                     版面提取【已完成】
│  │  ├─ __init__.py
│  │  ├─ models.py                数据契约：PdfTextSpan / PdfTextLine / PdfImageBlock
│  │  │                           / PdfVectorObject / PdfTable / PdfPageLayout
│  │  ├─ text.py                  字符提取 + font run 切分 + 行组装
│  │  ├─ tables.py                有线表格 + 无边框表格 + 续表恢复
│  │  ├─ vectors.py               矢量对象与表格边界线
│  │  ├─ images.py                内嵌图片编码与提取
│  │  ├─ reading_order.py         页眉页脚 / logo / 多栏阅读顺序
│  │  └─ layout.py                extract_pdf_layout 入口与编排
│  ├─ ir/                         中间表示【已完成，阶段 S4】
│  │  ├─ model.py                 ← document_ir.py
│  │  ├─ builder.py               ← ir_builder.py
│  │  └─ outline.py               ← pdf_outline.py
│  ├─ fonts/                      字体【已完成，阶段 S7】
│  │  ├─ resolver.py              ← font_resolver.py
│  │  ├─ embedding.py             ← font_embedding.py
│  │  └─ metrics.py               ← font_metrics.py
│  ├─ export/                     导出【已完成，阶段 S3】
│  │  ├─ __init__.py
│  │  ├─ text_utils.py            纯文本工具与字段取值
│  │  ├─ table_render.py          PdfTable → Word 表格
│  │  ├─ content.py               内容识别、行分组、可编辑块
│  │  ├─ document_setup.py        文档样式、分页、整页与区域图片渲染
│  │  └─ fidelity.py              export_fidelity_docx 及其放置逻辑
│  ├─ worker/                     worker 内部模块【已完成】
│  │  ├─ progress.py              取消与超时、阶段日志、流水线缓存
│  │  └─ quality.py               告警、门禁、渲染合并、保真验收、兜底
│  ├─ validate/                   回读校验【已完成，阶段 S5】
│  │  ├─ render.py                Word COM 渲染与页数检查
│  │  ├─ similarity.py            SSIM 与逐页对比
│  │  ├─ text_compare.py          覆盖率与行级 bbox 比对
│  │  └─ report.py                validate_docx_rendering 总入口
│  └─ service/                    服务入口【已完成，阶段 S7】
│     ├─ app.py                   ← pdf_to_word_service.py
│     ├─ jobs.py                  ← job_store.py
│     └─ ocr_quality.py
├─ tests/                         按模块对应命名
├─ tools/                         标定与诊断工具
├─ scripts/                       样本生成与基准
├─ fixtures/                      合成输入
└─ docs/                          文档
```

与初版方案的差异（搬迁后据实调整）：

- 版面提取用 `src/layout/`（而非 `extract/`），并保留 `src/pdf_layout.py` 作为对外入口，避免改动 `ir_builder`、`pdf_worker`、`fidelity` 等十余个调用方的导入路径。
- 导出包新增 `text_utils.py` 与 `content.py`、`document_setup.py`（初版只规划了 `shared.py`）；流式导出仍留在 `pdf_to_word_exporter.py`，未单独拆出 `legacy.py`。
- 新增 `src/worker/` 用于 worker 内部模块。

### 2.1 剩余目标目录与阶段对应（四处去留已定，2026-09-12）

`【】` 内为执行阶段；判断依据是"文件是否已大到影响定位与修改"，而不是目录树是否整齐。

| 目标包 | 阶段 | 现状与迁移来源 |
|---|---|---|
| `layout/` | S6 | **已完成**，来源 `pdf_layout.py` |
| `export/` | S3 | **已完成**，来源 `pdf_to_word_exporter.py` |
| `worker/` | S7 | **已完成**（worker 部分），来源 `pdf_worker.py` |
| `fonts/` | **S7（已决定拆）** | 待做：`font_resolver.py` → `resolver.py`、`font_embedding.py` → `embedding.py`、`font_metrics.py` → `metrics.py` |
| `ir/` | **S4** | 待做：`document_ir.py` → `model.py`、`ir_builder.py` → `builder.py`、`pdf_outline.py` → `outline.py` |
| `validate/` | **S5** | 待做：`docx_render_validation.py` 拆为 `render.py`、`similarity.py`、`text_compare.py`、`report.py` |
| `service/` | **S7 剩余部分** | 待做：`pdf_to_word_service.py` → `app.py`、`job_store.py` → `jobs.py`、`ocr_quality.py` |

四处原本未分配的位置，决定如下：

| 位置 | 决定 | 依据 |
|---|---|---|
| `fonts/` | **拆** | 三个模块（564 + 536 + 335 = 1435 行）构成一条单向流水线 `resolver → embedding → metrics`，命名已为同一族；按 `layout/`、`export/` 的既有做法分组与全仓一致，且组名不影响对外导入路径（保留三个门面文件） |
| `routing/` | **不拆** | `pdf_routing.py` 仅 281 行、9 个定义、无任何包内依赖，单文件本身已是恰当的模块粒度 |
| `export/formula.py` | **不拆** | `formula_omml.py` 仅 127 行、4 个定义，且只依赖 `latex2mathml` 与 `lxml`；单独成目录没有收益 |
| `export/primitives.py` | **不拆** | `ooxml_positioning.py` 879 行、27 个定义，内部已按"单位与页面原语 → 文本样式与绝对定位 → 图形与文本框 → 页眉页脚"三层排列；它的导入方只有 3 个（`pdf_to_word_exporter.py`、`export/fidelity.py`、`tests/test_ooxml_positioning.py`），且这 3 处各导入 11—13 个横跨四组的符号，拆开后导入复杂度不降反升。待办改为：删除死代码 `set_table_grid`、补公开面说明（见第 1.4 节） |

`fonts/` 与 `validate/` 之间存在一条跨包依赖需要记住：`font_metrics.py` 依赖 `docx_render_validation`（标定要用 Word 渲染回读实测偏移）。包级依赖方向为 `layout → fonts → validate`，无环；核查脚本与输出已归档。若 S5 之后该依赖造成困扰，可把"用 Word 实测偏移"这一步的调用方改为由 worker 注入，而不是让 `fonts/` 反向依赖 `validate/`。

三条设计原则：

1. **先文件后目录。** 每个阶段先在原文件内用分区注释与区块顺序稳定边界，等测试能证明行为不变，再物理搬迁。一次搬到位难以验证，也难以回滚。
2. **对外入口保留。** 被搬迁模块的既有导入路径保持可用：入口文件精确转出子模块定义的全部符号，使 `tests/`、`scripts/`、`tools/` 与其余 `src` 模块无需同步改导入。这比初版设想的"在 `__init__.py` 里 re-export"更稳，因为它同时保住了模块属性访问（如 `pdf_worker.process_job`）。
3. **顶级模块保持扁平可达。** 目录深度最多两层，避免过长的相对导入路径。

---

## 3. 执行步骤

| 步骤 | 内容 | 前置条件 | 验证方式 |
|---|---|---|---|
| S0 | 建立基线快照 **（已完成）** | 无 | 基线已固化到 `tmp\baseline\`：`first20-result.json`、`first20-summary.json`、`f103-result.json`、`f103-summary.json`，取自 `fidelity_first20_v5` 与 `f103_test\fidelity_run_v4` 两次既有运行 |
| S1 | 删除死代码与清理临时文件 **（已完成）** | S0 | 96 个测试通过；`src/fidelity_writer.py` 已移除；详见第 1.1 节 |
| S2 | 常量与测试文件改名 **（第 1、2 项已完成，第 3 项归入 S4）** | S1 | 96 个测试通过；旧常量名无残留引用；详见第 1.2 节 |
| S3 | 抽出 `export/shared.py`，落定导出模块边界 **（已完成，方案调整为 `export/` 五模块）** | S2 | 见第 1.3.6—1.3.7 节 |
| S4 | 建 `ir` 包（`document_ir` → `model.py`，`ir_builder` → `builder.py`，`pdf_outline` → `outline.py`） **（已完成；顺带消除两处导入期环）** | S3 | 见第 1.5 节 |
| S5 | 抽 `validate/` 包 **（已完成）** | S4 | 见第 1.6 节；重点确认 `fidelity_acceptance` 字段不变 |
| S6 | 拆 `extract/` 与 `routing/` 包 **（已完成：版面提取落地为 `src/layout/`；`routing/` 经评估不拆）** | S5 | 见第 1.3.6—1.3.7 节与第 2.1 节 |
| S7 | 拆 `service/` 包（worker、字体与服务） **（已完成：`src/worker/`、`src/fonts/`、`src/service/`）** | S6 | 见第 1.7 节 |
| S8 | 旧导出路径降级为 `export/legacy.py`，评估是否整体移除 **（已完成：按"整体移除"执行）** | S7 | 见第 1.8 节 |

关于 S8 的决策依据（决策前记录，保留备查）：旧路径的引用分布是 `export_ir_to_docx` 8 处、`export_text_pages_to_docx` 10 处、`export_source_pages_to_docx` 7 处、`export_results_to_docx` 8 处，除定义与相互调用外，全部来自 `tests/test_flow_export.py`、`tests/test_formula_omml.py`、`tests/test_pdf_outline.py` 与 `scripts/run_external_pdf_benchmark.py`。当时的三个选项：

1. 迁入 `src/export/legacy.py`，保留对外符号 —— 纯搬迁，可逆；
2. **整体删除**（连同仅测旧路径的用例与基准脚本）—— 输出最小，代价是失去外部 PDF 基准的对比能力；**已按此执行，见第 1.8 节**；
3. 保持现状 —— 零风险，但 `pdf_to_word_exporter.py` 会一直带着约 530 行旧实现。

删除时注意：`test_flow_export.py` 中另有一部分用例测的是内容组装逻辑（`_group_text_lines`、`_layout_lines_to_text` 等），那些函数由 `ir_builder` 在活路径上调用，不能随旧路径一起删——它们已迁入 `tests/test_content_grouping.py`。

---

## 4. 安全网

整理过程的唯一验收标准是行为不变，每一步都要跑同一组对照：

1. **测试底线**：`& .\.venv\Scripts\python.exe -m unittest discover -s tests`，96 个全部通过；用例数不得无故减少（删除死代码导致减少时需逐条说明原因）。
2. **样本对照**：基线已固化在 `tmp\baseline\`，逐项比对 `page_delta`、`blank_pages`、`ssim.mean_ssim` / `min_ssim`、`text_layout.matched_line_count`、`fonts.embedding.entries`、`fidelity_auto_fallback_pages`。任何一项变化都要查明原因。

   当前基线值：

   | 指标 | `first20-summary.json` | `f103-summary.json` |
   |---|---:|---:|
   | `page_delta` | 0 | 0 |
   | `mean_ssim` | 0.9603 | 0.9722 |
   | `min_ssim` | 0.8866 | 0.9699 |
   | `matched_line_count` | 275 | 0 |
   | `median_bbox_error` | 0.37 | 0.0 |

   基线取自既有的一次运行产物，未按 S0 原计划重跑；如需重新生成，命令与参数见 `docs/agent-handoff.md` 第 1.1 节。
3. **导入面检查**：搬迁后执行一次 `python -c "import src.pdf_to_word_service"`，确认服务入口仍可导入。
4. **版本控制节奏**：`main` 分支目前停留在 `5a19209`，未提交改动量极大。建议先按功能分批提交现状代码，再开始整理；否则无法用 `git diff` 区分整理引入的改动与原有未提交改动。提交需经确认后执行。

---

## 5. 不建议做的事

- **不拆分旧流式导出部分。** 它虽然约占 1900 行，但已处于"将删未删"的隔离状态，为它做精细拆分是无效投入。要么整体留在 `legacy.py`，要么整体移除。
- **不重命名公开函数名。** `export_fidelity_docx`、`build_document_ir`、`extract_pdf_layout` 等名字在文档、测试、脚本与外部基准中反复出现，改名收益低、辐射面大。
- **不引入新的抽象层。** 本次整理的目标是让文件边界对应流水线的七个阶段，便于阅读与定位，不增加间接层。
- **不逐个归档 `tmp/` 中的历史产物。** 直接清空即可。

---

## 6. 完成标准

1. `src/` 下每个文件不超过约 900 行，且文件名能直接对应流水线的一个阶段；
2. 无死代码、无重复含义的常量、无命名不成对的模块；
3. `tests/` 中每个测试文件对应一个源模块；
4. 任一阶段完成后，S0 基线样本的指标逐项一致；
5. 后续接手者能按"路由 → 提取 → IR → 字体 → 导出 → 校验"的顺序，在目录树上直接找到对应文件。
