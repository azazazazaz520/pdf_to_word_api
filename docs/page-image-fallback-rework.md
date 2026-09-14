# 整页贴图兜底重构：从图像相似度到文字覆盖率

编制日期：2026-09-13
适用范围：`D:\pdf_validation`（PDF 转 Word 独立验证项目）
状态：**已实施并验证**

本文承接 [外部 PDF 转换测试报告](external-conversion-test-2026-09.md)。该报告记录的"6 份样本里只有 1 份可用"并非导出能力不足，而是本文所述的两个缺陷所致。缺陷修复后同一批样本的实际可编辑页由 3 页升至 40 页。

---

## 1. 问题现象

以 GPT-3 样本（20 页）为例，`process_job` 报告的结论与打开成品数出的结果完全相反：

| 项目 | 报告声称 | 打开 DOCX 实际 |
|---|---|---|
| 可编辑页 | 18 | 3 |
| 整页图片页 | 2 | 17 |
| 触发兜底的页 | 无 | 未记录 |

源 PDF 第 13—19 页各有 3137—3349 个字符，导出后这些页面上一个字都没有，全部变成整页截图。而它们的文字覆盖率实测为 0.992—0.997，即文字几乎完整地写入过。

## 2. 根因一：兜底依据选了图像相似度

`src/pdf_worker.py` 的兜底候选筛选同时使用两个条件：

```python
if score < fallback_threshold:          # 图像相似度低于 0.80
    candidates.append(page_number)
elif coverage < coverage_threshold:     # 文字覆盖率低于 0.90
    candidates.append(page_number)
```

第一个条件与实际目标相矛盾。`src/worker/quality.py` 的 `_apply_page_image_fallback` 命中后的处理是**丢弃整页全部文字块**，只保留一张整页图片：

```python
page.blocks = [IRBlock(kind="page_image", ...)]
page.route = "page_image"
page.editable = False
```

也就是说，"图像相似度偏低"这一条会把文字内容完整、位置正确的页面替换为不可编辑的截图。GPT-3 样本的实际数据：

| 页 | 文字覆盖率 | 图像相似度 | 原判定 | 保留的文字 |
|---|---|---|---|---|
| 13 | 0.9952 | 0.769 | 贴图 | 全部丢弃 |
| 14 | 0.9932 | 0.783 | 贴图 | 全部丢弃 |
| 15 | 0.9918 | 0.765 | 贴图 | 全部丢弃 |
| 16 | 0.9946 | 0.759 | 贴图 | 全部丢弃 |
| 18 | 0.9966 | 0.787 | 贴图 | 全部丢弃 |

这类页面共 15 页。它们的文字覆盖率均在 0.85 以上，相似度在 0.76—0.93 之间——像素比对偏低来自渲染细节差异，而非内容缺失。

## 3. 根因二：兜底结果没有回到调用方

`_run_render_validation` 在兜底与第二次导出后重建了质量报告：

```python
quality = ir.quality_report(compact=page_count > 200)
```

但该函数接收的 `quality` 是形参，这一行只重新绑定了**函数内的局部名字**，调用方 `process_job` 持有的字典从未更新。后果是：

- 最终报告沿用"兜底前"的页面状态，`editable_page_count` 恒为兜底前的值；
- `fidelity_auto_fallback_pages` 恒为 `None`（该键由兜底分支写入局部字典后即丢弃）；
- 调用方在函数返回后继续写入 `quality["fidelity"]`，作用于旧字典。

这正是 [下一阶段规划](next-phase-plan.md) 记录但未定位的"`fidelity_auto_fallback_pages` 恒为 `None`"。

## 4. 修改内容

### 4.1 兜底只依据文字覆盖率

`src/pdf_worker.py` 的候选筛选改为单一条件：

```python
for page_number, coverage in coverage_pages.items():
    page = ir.pages[page_number - 1]
    if page.route in {"page_image", "blank"}:
        continue
    if coverage < coverage_threshold:
        candidates.append(page_number)
```

图像相似度不再参与兜底判定，仅保留在报告的 `ssim` 字段中作为参考值。服务层已失效的 `ssim_fallback_threshold` 配置项、环境变量校验与下发字段一并删除。

### 4.2 报告随兜底结果更新

`_run_render_validation` 改为返回重建后的质量报告：

```python
quality = _run_render_validation(payload, quality, ...)
```

调用方接收返回值，使 `editable_page_count`、`visual_only_page_count`、`route_summary`、`fidelity_auto_fallback_pages` 均反映兜底后的真实状态。

### 4.3 兜底原因标识

`_PAGE_IMAGE_FALLBACK_REASON` 由 `auto_fallback_ssim_below_threshold` 改为 `auto_fallback_text_coverage_below_threshold`，阶段日志的 `route_reason` 同步改为 `text_coverage_below_threshold`。图像相似度的验收告警 `fidelity_ssim_below_threshold` 与之无关，保持原样。

## 5. 验证结果

### 5.1 GPT-3 样本逐页对照

修改后，报告与成品逐页一致（不一致页数 0）：

| 结果 | 页 | 依据 |
|---|---|---|
| 保留可编辑文字 | 2、4—7、9、10、13—20（15 页） | 覆盖率 0.91—0.997 |
| 整页贴图兜底 | 1、3、8 | 覆盖率 0.69 / 0.89 / 0.85，文字确有丢失 |
| 原样保留的贴图页 | 11、12 | 源页仅 177 / 87 字，路由本就是 page_image |

兜底页由 18 页降至 3 页，可编辑页由 3 页升至 15 页，且报告值等于实际值。

### 5.2 六份外部样本复测

直接打开生成的 DOCX 统计可编辑页，与引擎自报并列：

| 样本 | 页 | 耗时(秒) | 自报可编辑 | 实际可编辑 | 整页图片 | 门禁 |
|---|---|---|---|---|---|---|
| Attention | 15 | 168.74 | 9 | 9 | 6 | passed |
| NIST SP 800-53r5 | 20 | 80.48 | 3 | 3 | 17 | passed |
| ResNet | 12 | 53.22 | 0 | 0 | 12 | passed |
| GPT-3 | 20 | 50.38 | 15 | 15 | 5 | passed |
| NASA 系统工程手册 | 20 | 42.84 | 7 | 7 | 13 | passed |
| BERT | 16 | 21.28 | 6 | 6 | 10 | error |

实际可编辑页合计 40 / 103（修改前为 3 页），**自报与实际不符的样本 0 个**。

Attention 的兜底页为 1、4、8、13、14、15，其覆盖率实测均为 0.0000。逐页比对源 PDF 与 Word 回读结果确认该判定成立：这 6 页在回读文件中确无任何文字，而其余 9 页的字符数与源页相差不超过 3 个字（如 3631 / 3634、2673 / 2674）。兜底的依据是准确的。

### 5.3 门禁

| 项目 | 结果 |
|---|---|
| `python -m unittest discover -s tests` | 85 个测试通过 |
| `tmp/check_imports_dynamic.py` | 41 个模块，问题 0 个 |
| `tmp/check_names.py` | 可疑名字 0 个 |
| `tmp/verify_against_baseline.py` | 实验指导书前 20 页与 F103 的关键保真度指标均与基线一致 |

F103 样本此前长期存在基线差异，原因之一即 4.2 节的报告覆盖；修复后该样本的关键指标亦与基线一致。

## 6. 仍然存在的问题

**Attention 与 BERT 的 Word 渲染校验失败。** 前者报 `returncode=0, 未能生成 PDF`（耗时 110—169 秒），后者报 `WORD_OPEN_FAILED`。这两份文档的 6 页与 10 页在导出时丢失了文字，兜底贴图只是抢救结果，**导出环节本身为何丢字尚未查明**。

**BERT 的形状尺寸超出范围。** `<wp:extent>` 的 cx 取值最大达 1 796 527 316 EMU（A4 页宽约 7 560 056 EMU），Word 因此拒绝打开文档。该项与文字丢失可能同源，均指向版面提取在该文档上产出退化包围盒。
