# src 结构收敛方案

编制日期：2026-09-12
适用范围：`D:\pdf_validation`
前置文档：`docs/file-organization-plan.md`（S0—S8 已执行完毕）

---

## 1. 现状量化

`src/` 下 52 个 `.py`、13886 行。文件数本身不算失控，问题在于三类冗余。

### 1.1 冗余一：12 个门面文件与重复的导入路径

| 门面文件 | 行数 | 实现在 |
|---|---:|---|
| `pdf_layout.py` | 123 | `layout/` 七个模块 |
| `pdf_to_word_exporter.py` | 124（含 5 个自有函数） | `export/` 五个模块 |
| `pdf_outline.py` / `ir_builder.py` / `document_ir.py` | 12 / 14 / 17 | `ir/` 三个模块 |
| `font_resolver.py` / `font_metrics.py` / `font_embedding.py` | 20 / 22 / 26 | `fonts/` 三个模块 |
| `docx_render_validation.py` | 25 | `validate/` 四个模块 |
| `pdf_to_word_service.py` | 30 | `service/app.py` |
| `ocr_quality.py` / `job_store.py` | 14 / 11 | `service/ocr_quality.py`、`service/jobs.py` |
| **合计** | **438 行、12 个文件** | 全部是转出语句，无任何逻辑 |

经门面路径导入的调用点共 **46 处**（`pdf_layout` 10、`pdf_to_word_exporter` 8、`document_ir` 6、`docx_render_validation` 5、`font_resolver` 5、`font_embedding` 3、`font_metrics` 2、`ir_builder` 2、`pdf_outline` 2、`ocr_quality` 2、`pdf_to_word_service` 1）。

后果：同一个模块存在两条导入路径（如 `from src.pdf_layout import X` 与 `from src.layout.models import X`），阅读时无法判断哪条是"正路"。

### 1.2 冗余二：`src/` 下混进运行时目录

| 目录 | 内容 | 说明 |
|---|---|---|
| `src/model_cache/` | `font_metrics.json` | 字体标定缓存；仓库根目录另有一份 `model_cache/` |
| `src/service_data/` | 102 个文件（`jobs/` 与任务目录） | 服务运行状态；根目录另有一份 `service_data/` |

成因：以 `-m src.pdf_to_word_service` 之外的方式从 `src` 目录内启动服务时，`PDF_SERVICE_DATA_ROOT` 之类的相对路径落在 `src/` 下。它们被 `.gitignore` 忽略，但会让 `src/` 显得杂乱，且两份数据容易混淆。

### 1.3 冗余三：编排层未拆完

`pdf_worker.py` 1034 行，其中 `process_job()` 单个函数仍约 636 行。`docs/file-organization-plan.md` 第 1.3.8 节已定位两块可提取的内联逻辑（OCR 执行块约 119 行、版面提取块约 87 行），但当时因连续两次手工改缩进损坏文件而回退。

### 1.4 次要问题

| 位置 | 行数 | 问题 |
|---|---:|---|
| `worker/` | 406（2 个模块） | 包内只有 `progress.py`(99) 与 `quality.py`(307)，包级目录收益偏低 |
| `run_validation.py` | 238 | 混合两种职责：`build_pipeline`（PaddleOCR 流水线构建，被 worker 调用）与 CLI 入口 |
| `pdf_routing.py` | 281 | 页面路由，与 `layout/` 同属"分析 PDF 页面"，却留在顶层 |
| `ooxml_positioning.py` | 864 | 26 个定义、3 个导入方，是最大的未拆模块 |
| `formula_omml.py` | 127 | 外部依赖 `latex2mathml` + `lxml`，与 OOXML 相关但独立 |

---

## 2. 三条可选路线

### 路线 A：只清门面与运行时目录（文件数 52 → 40）

删除 12 个门面文件并改写 46 处导入；清理 `src/model_cache/`、`src/service_data/`。

- 优点：删除的是纯转出代码，风险最低；同时消除双导入路径这一最大困惑源。
- 缺点：包内碎片（`worker/` 只有两个模块、`run_validation.py` 混职责）保持原样。

### 路线 B：A + 关键合并与拆分（文件数 52 → 约 43）

在 A 之上：

- `worker/` 并入 `service/`：`progress.py` → `service/worker_progress.py`，`quality.py` → `service/quality.py`（服务层原本就包含 `app.py`、`jobs.py`，把 worker 编排的辅助模块放在同一包内更连续）；
- `run_validation.py` 拆分：`build_pipeline` 移入 `service/pipeline.py`，`main()` 留在 `run_validation.py`（CLI 入口）；
- `pdf_routing.py` 移入 `layout/routing.py`（与 `layout/` 同属页面分析，且 `layout/models.py` 已在依赖它）；
- 完成 `process_job()` 的两块提取（OCR 执行、版面提取），使 `pdf_worker.py` 落到约 800 行。

- 优点：消除全部三类冗余，且不破坏已按流水线阶段划分的 `layout/`、`export/`、`ir/`、`fonts/`、`validate/`。
- 缺点：涉及移动文件与模块名变更，需要重跑基线对照。

### 路线 C：按流水线阶段彻底重组（文件数 52 → 约 48）

把 7 个包重组为 5 个阶段包（如 `extract/` = 现 `layout/` + `pdf_routing.py`；`emit/` = 现 `export/` + `ooxml_positioning.py` + `formula_omml.py` 等）。

- 优点：目录名与流水线阶段一一对应。
- 缺点：`layout/`、`export/`、`validate/` 的内部划分是 S3—S8 逐个验证过的结果，重新组合会破坏已确认的依赖方向（如 `layout → fonts → validate`），收益仅是命名更整齐。

**建议：路线 B。** 路线 C 的改动面远大于收益；路线 A 则留下 `worker/` 与 `run_validation.py` 两处已被识别的问题。

---

## 3. 任务清单（按性价比排序）

| 任务 | 内容 | 影响面 | 验证方式 |
|---|---|---|---|
| **T1** | 删除 12 个门面文件，改写 46 处导入指向子包 | 12 个文件删除 + 46 处导入 | 动态导入检查（52 个模块）+ 81 个测试 + 基线指标 |
| **T2** | 清理 `src/model_cache/`、`src/service_data/`；确认服务数据根路径不落在 `src/` 下 | 2 个目录 | 启动服务一次，确认数据写在仓库根目录 |
| **T3** | 拆分 `run_validation.py`：`build_pipeline` → `service/pipeline.py` | 1 个文件拆分 + 2 处导入 | 同上 |
| **T4** | `pdf_routing.py` → `layout/routing.py` | 1 个文件移动 + 约 5 处导入 | 同上 |
| **T5** | `worker/` 并入 `service/`（包级目录数 7 → 6） | 2 个文件移动 + 约 4 处导入 | 同上 |
| **T6** | 完成 `process_job()` 两块提取，`pdf_worker.py` 约 1034 → 800 行 | 1 个文件 | 同上，另需核对 `stages.jsonl` 阶段序列不变 |
| **T7**（可选） | `ooxml_positioning.py` + `formula_omml.py` → `ooxml/` 包 | 2 个文件移动 + 3 处导入 | 同上 |

T1 与 T2 不改变任何行为，可立即执行。T3—T5 是文件移动，与 S4—S7 同型、已有成熟做法与门禁。T6 需要遵守第 1.3.8 节记录的缩进规则。

T7 的判据仍在"是否影响定位与修改"：`ooxml_positioning.py` 的 864 行偏长，但只有 3 个导入方、各导入 11—13 个横跨多组的符号（页面／文本／图形／页眉页脚），拆开会使导入复杂度上升；建议仅在确实出现定位困难时再处理。

---

## 4. 执行门禁

每个任务完成后必须全部通过：

1. `python -m unittest discover -s tests`：81 个测试全绿；
2. `python tmp/check_imports_dynamic.py`：52 个模块（T5 后为相应数量）导入成功、相对导入符号无缺失；
3. `python tmp/check_names.py`：0 处可疑名字；
4. `python tmp/check_cycles3.py`：导入期无新增环（当前仅 `export.fidelity ↔ pdf_to_word_exporter` 一处，由延迟导入打破）；
5. `python tmp/verify_against_baseline.py`：实验指导书前 20 页与 `tmp/baseline/` 逐项一致（`page_delta=0`、`mean_ssim=0.9603`、`min_ssim=0.8866`、匹配 275 行、中位数 0.37pt）；F103 与交接文档记录吻合（`min_ssim=0.9151`、`mean_ssim=0.9599`）。

**删除类改动（T1、T2）的边界判据以动态导入检查为准**，不以静态可达性分析为准——后者在 S8 中两次误判（详见 `docs/file-organization-plan.md` 第 1.8.2 节）。

---

## 5. 完成标准

1. 同一模块只有一条导入路径；
2. `src/` 下只有代码与 `__init__.py`，无运行时数据目录；
3. `src/` 顶层 `.py` 由 17 个降到 5—6 个。删除 12 个门面后应剩余：

   | 顶层模块 | 行数 | 处置 |
   |---|---:|---|
   | `pdf_worker.py` | 1034 | T6 提取后约 800 行 |
   | `ooxml_positioning.py` | 864 | T7（可选） |
   | `pdf_routing.py` | 281 | T4 移入 `layout/` |
   | `run_validation.py` | 238 | T3 拆分 |
   | `formula_omml.py` | 127 | T7（可选） |

   执行 T3、T4 后顶层剩 3 个（`pdf_worker.py`、`ooxml_positioning.py`、`formula_omml.py`）；T7 若执行，最终剩 1 个（`pdf_worker.py`）。
4. 无单文件超过约 900 行（`pdf_worker.py` 与 `ooxml_positioning.py` 需 T6、T7 才能满足）；
5. 任一步骤完成后，基线与测试门禁全部通过。
