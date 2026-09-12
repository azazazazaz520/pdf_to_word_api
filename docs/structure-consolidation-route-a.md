# 路线 A 执行记录：删除门面层与运行时目录

执行日期：2026-09-12
依据文档：`docs/structure-consolidation-plan.md`（路线 A）
前置状态：`docs/file-organization-plan.md` 的 S0—S8 已执行完毕

---

## 1. 执行内容

### 1.1 T1：删除门面层

改写 24 个文件、45 条导入语句（拆分为 52 条），把每个符号的导入从门面路径指向其真实定义模块。改写规则：

- src 内部保持相对导入风格，层级按文件所在深度换算（如 `src/export/content.py` 用 `..layout.models`）；
- tests、scripts、tools 保持绝对风格（`src.layout.models`）；
- 随导入路径一并修正两处测试的补丁目标：`patch("src.docx_render_validation.render_docx_to_pdf")` → `patch("src.validate.render.render_docx_to_pdf")`，`import src.pdf_to_word_service` → `import src.service.app`。

删除 11 个纯门面文件：`pdf_layout.py`、`document_ir.py`、`ir_builder.py`、`pdf_outline.py`、`docx_render_validation.py`、`font_resolver.py`、`font_embedding.py`、`font_metrics.py`、`pdf_to_word_service.py`、`job_store.py`、`ocr_quality.py`。

第 12 个门面 `pdf_to_word_exporter.py` **不是纯门面**——它含 5 个自有函数：

| 函数 | 用途 | 使用方 |
|---|---|---|
| `render_page_image` | 公开的单页渲染接口 | `pdf_worker.py`、`worker/quality.py` |
| `_add_formula_omml_element`、`_bookmark_name`、`_add_bookmark`、`_add_toc_field` | OOXML 写入工具 | `export/fidelity.py`（运行期延迟导入） |

因此保留为真实模块，去掉不再被引用的 4 条转出语句，并改名为语义相符的 **`page_render.py`**（108 行）。

### 1.2 T2：清理运行时目录

`src/model_cache/`（1 个文件）与 `src/service_data/`（102 个文件）是历史运行产物：默认数据根锚定仓库根（`ROOT = parents[2]`），这两处是早期以 `src` 为工作目录运行留下的。删除前已确认仓库根存在对应数据，并整体归档到 `D:\pdf_validation_tmp_archive\src_runtime_dirs_sept12`（103 个文件）。

### 1.3 顺带修复的两处路径缺陷

删除 `src/service_data` 后核对数据根，发现**两处 `parents[N]` 计算在 S7 搬迁后少算了一层**——此前所有测试与基线对照都未覆盖到：

| 位置 | 搬迁后错误指向 | 修正 |
|---|---|---|
| `src/service/app.py:31` | `D:\pdf_validation\src`（`parents[1]`） | `parents[2]` → 仓库根 |
| `src/fonts/metrics.py:84` | `D:\pdf_validation\src\model_cache`（`parent.parent`） | `parents[2]` → 仓库根 |

修正后实测：`CONFIG.data_root = D:\pdf_validation\service_data`、`JOB_ROOT = D:\pdf_validation\service_data\jobs`、字体标定默认缓存 = `D:\pdf_validation\model_cache\font_metrics.json`（文件存在）。

---

## 2. 执行前后对比

| 指标 | 执行前 | 执行后 |
|---|---:|---:|
| `src/` 下 `.py` 文件 | 52 | **41** |
| 顶层 `.py`（除 `__init__`） | 17 | **6** |
| 其中门面文件 | 12 | **0** |
| 门面转出行数 | 438 | 0 |
| `src/` 下运行时目录 | 2（103 个文件） | **0** |
| 同一模块的导入路径数 | 2 | **1** |
| 导入期环 | 1 处（`export.fidelity ↔ pdf_to_word_exporter`） | **0** |
| 自动化测试 | 81 | 81（全部通过） |

顶层剩余的 6 个模块：

| 模块 | 行数 | 说明 |
|---|---:|---|
| `pdf_worker.py` | 1026 | 任务编排入口 |
| `ooxml_positioning.py` | 864 | 绝对定位原语 |
| `pdf_routing.py` | 281 | 页面路由 |
| `run_validation.py` | 238 | 流水线构建与 CLI |
| `formula_omml.py` | 127 | 公式 → OMML |
| `page_render.py` | 108 | 单页渲染与 OOXML 工具 |

**附带收益**：`page_render.py` 不再反向依赖门面后，原先唯一的导入期环消失，`src` 现已无环。

---

## 3. 验证结果

| 门禁 | 结果 |
|---|---|
| 动态导入检查 | 41 个模块，0 问题 |
| 缺名检查 | 0 处可疑名字 |
| 导入期环检查 | 无环 |
| 自动化测试 | 81 个全部通过 |
| 实验指导书前 20 页 | 与基线逐项一致：`page_delta=0`、`mean_ssim=0.9603`、`min_ssim=0.8866`、匹配 275 行、中位数 0.37pt |
| F103 样本 | 与交接文档吻合：`min_ssim=0.9151`、`mean_ssim=0.9599`、`media_count=78` |
| 服务配置路径 | `data_root` 与字体缓存均落在仓库根，且目标文件存在 |

---

## 4. 未执行项

`docs/structure-consolidation-plan.md` 中路线 A 范围之外的任务（T3—T7）保持原状：

- `run_validation.py` 仍混合流水线构建与 CLI（T3）；
- `pdf_routing.py` 仍在顶层（T4）；
- `worker/` 仍为两个模块的独立包（T5）；
- `process_job()` 仍约 630 行（T6）；
- `ooxml_positioning.py`（864 行）与 `formula_omml.py`（127 行）未合并为 `ooxml/`（T7）。

其中 T6 与 T7 与"完成标准第 4 条（无单文件超过 900 行）"相关，若后续要求满足该条，需执行这两项。

---

## 5. 归档位置

| 内容 | 路径 |
|---|---|
| 执行前完整备份（src/tests/scripts/tools） | `D:\pdf_validation_tmp_archive\planA_backup_sept12` |
| 删除的 `src/model_cache`、`src/service_data` | `D:\pdf_validation_tmp_archive\src_runtime_dirs_sept12` |
| 勘查与改写脚本、校验输出 | `D:\pdf_validation_tmp_archive\section140-sept-2026-09-12` |
