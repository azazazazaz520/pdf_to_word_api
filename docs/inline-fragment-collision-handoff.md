# 行内文字重叠缺陷：修复记录与交接

编制日期：2026-09-13
适用范围：`D:\pdf_validation`（PDF 转 Word 独立验证项目）

本文是 [inline-fragment-collision-fix.md](inline-fragment-collision-fix.md) 的收尾摘要。完整过程（含每一次尝试的数据）见该文档第 1—20 节。

---

## 一、已完成并验证的修复

### 1. 字符抽取按 pdfium 字符索引（`src/layout/text.py`）

**根因**：`pypdfium2` 的 `get_text_range(i, i+1)` 第二个参数不是长度，调用它会返回从该位置到结尾的整段文本；而 `get_charbox(index)` 使用字符索引。两者分属不同索引体系，在连字与 `\ufffe` 断行标记处错位。

**修法**：改用 pdfium 原始 API 按字符索引取码位。

```python
def _character_count(text_page) -> int:
    return int(pdfium_raw.FPDFText_CountChars(text_page.raw))

def _character_at(text_page, index) -> str:
    code = int(pdfium_raw.FPDFText_GetUnicode(text_page.raw, index))
    ...
```

同时补上 XML 控制字符过滤（原始 API 会返回 C0 控制字符，写入 DOCX 会触发写入库校验失败）。

**验证**：修复后字符抽取层、版面模型、IR 三层的 `Kaiming` 均正确（修复前只有 IR 层正确）。

### 2. 片段分组与归带（`src/layout/text.py`、`src/layout/models.py`）

| 成因 | 修法 |
|---|---|
| 片段按 0.25pt 精确字号切分，导致每字符一片段（一行 111 字切 76 片） | 改为按整行中位字号、分档容差 |
| 同一词内因逐字形度量差异被拆开 | 增加词界判定，同词内不断开 |
| 退化字符（零高度空格，字号 1.00）被选为片段代表，整段字号被写成 1.00 | 代表字符与 bbox 只在正常外框字符中选取 |
| 固定 6.0pt 归带容差与小字号不匹配，零高度空格把相邻行并入同一带 | 容差改为随字号缩放；退化字符按纵向位置归带 |

**效果**：GPT-3 成品穿插由 8 行降为 **0**；Attention 由 4 行降为 **1**；其余 5 个样本保持 0。

### 3. 验收指标

| 指标 | 基线 | 当前 | 判定 |
|---|---|---|---|
| first20 `page_delta` | 0 | **0** | 一致 |
| first20 匹配行 | 275 | **276** | 优于 |
| first20 未匹配行 | 18 | **17** | 优于 |
| first20 bbox 中位 | 0.37 | **0.37** | 一致 |
| first20 bbox 最大 | 13.451 | **13.451** | 一致 |
| first20 SSIM 均值 | 0.9603 | **0.9612** | 优于 |
| first20 SSIM 最低 | 0.8866 | **0.8867** | 优于 |
| F103 | 与基线一致 | **与基线一致** | 一致 |
| 单元测试 | 87 | **87 通过** | 通过 |
| 导入 / 命名检查 | — | 41 模块 0 问题 / 0 可疑命名 | 通过 |
| 字形穿插（7 样本） | 4 行 | **1 行** | **未达成** |

---

## 二、剩余 1 处的诊断（已定位，未修）

`Attention Is All You Need` 第 11 页 `[11]` 参考文献条目。

### 现象

源 PDF 正常：

```text
'[11] Kaiming He, Xiangyu Zhang, Shaoqing Ren, and Jian Sun. Deep residual learning for im\ufffeage r'
```

成品该处叠加了两层文字（渲染图确认）：一层是正确文本，另一层是乱码文本。

### 已排除

| 环节 | 状态 |
|---|---|
| 字符抽取 | 正确 |
| 版面模型 | 正确（`Kaiming` 存在、`Kagaeim` 不存在） |
| IR（导出前） | 正确 |
| **DOCX（导出后）** | **错误** |

因此偏差产生于**导出阶段**。

### 关键数据

该页 DOCX 第 11 节有 **322 个定位框，但只有 97 个不同的 `y` 值**——多个框落在同一纵坐标（最多一处 5 个框同为 y=240.8）。这是"两层文字叠加"的直接来源。

### 已尝试且失败的修法（均在 `inline-fragment-collision-fix.md` 有完整数据）

| # | 修法 | 结果 |
|---|---|---|
| 1 | 导出层整行一框（字号取中位 / 90 分位 / 补 `pdf_font_name`） | 匹配行 275→268、262 |
| 2 | 导出层一框内写多个文字运行（每片段自带字号） | 匹配行改善至 279，但 bbox 中位由 0.37 恶化到 4.68pt |
| 3 | 空格归入前一片段 | 未改善空格，4 项测试失败 |
| 4 | 片段文本直接拼接（绕过归一化） | 空格保留成功，5 项测试失败 |
| 5 | 占位字符替换 → 归一化 → 还原 | 空格保留成功、测试失败降至 1 项，但 first20 匹配 275→266、F103 校验失败 |

### 下一步方向

**修法 5 在功能上正确**（空格得以保留、单测接近全绿），但会改变行文本内容，进而影响 **`src/validate/text_compare.py`** 的行匹配——该函数以"整行文本相等或公共前缀不少于 4 字符"为条件，对词间空格敏感。

因此继续推进需**同时**检查该匹配逻辑，而不是只改片段文本。这是跨模块工作，建议作为独立任务（本项目的"导出字符顺序与回读匹配"任务）推进，不要与已验证的修复混在一起。

---

## 三、仓库状态

**全部改动尚未提交。** 涉及 12 个文件改动 + 5 个新文件：

```
 M src/export/fidelity.py          兜底原因字符串
 M src/fonts/resolver.py           删除失效的 style_key
 M src/layout/images.py            APP14 JPEG 直通条件
 M src/layout/models.py            归带容差随字号缩放；删除 style_key
 M src/layout/text.py              字符索引抽取、片段分组、退化字符处理
 M src/pdf_routing.py              路由改用 pypdfium2、类别判据、按比例罚分
 M src/pdf_worker.py               兜底依据改覆盖率、报告回传调用方
 M src/service/app.py              删除失效的 ssim_fallback_threshold
 M src/worker/quality.py           兜底原因标识
 M tests/test_pdf_routing.py       路由测试更新
?? docs/external-conversion-test-2026-09.md
?? docs/inline-fragment-collision-fix.md
?? docs/page-image-fallback-rework.md
?? docs/routing-text-detection-plan.md
?? tests/test_text_span_grouping.py
```

建议分两个提交：

1. **路由与贴图机制**：`pdf_routing.py`、`pdf_worker.py`、`service/app.py`、`worker/quality.py`、`export/fidelity.py`、`tests/test_pdf_routing.py` 与两份对应文档
2. **行内片段与归带**：`layout/text.py`、`layout/models.py`、`fonts/resolver.py`、`layout/images.py`、`tests/test_text_span_grouping.py` 与 `inline-fragment-collision-fix.md`

---

## 四、如何继续

1. 复现：转换 `tmp/external_pdfs/attention-is-all-you-need.pdf`，检查第 11 页 `[11]` 条目
2. 判定：`tmp/collision_pairs.py` 对比源与成品的字形穿插行数
3. 起点：导出层为每行写框的位置与数量（该页 322 框 / 97 个不同 y 值）
4. 守门：每次改动后运行 `python -m unittest discover -s tests` 与 `tmp/verify_against_baseline.py`
