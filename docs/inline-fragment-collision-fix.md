# 行内文字片段被拆成独立定位框的修复记录

编制日期：2026-09-13
适用范围：`D:\pdf_validation`（PDF 转 Word 独立验证项目）
状态：**版面层已修复，导出层不变量受阻**

本文记录"同一行文字被拆成多个绝对定位框、导致字形互相穿插"缺陷的定位与修复过程。

---

## 1. 现象

`Attention Is All You Need` 与 `GPT-3` 两份样本的部分页面，在 Word 成品里出现逐字交错的乱码，例如：

```text
'wofit ihn- rceolanttievxetly le satrrnaiignhgtf iosr awlsaord si smcaillianr'
```

该字符串由两句话逐字符交替而成。源 PDF 同一位置抽取结果为正常文字。

## 2. 根因

### 2.1 片段分组按精确字号匹配

`src/fonts/resolver.py::style_key` 把字号量化到 0.25pt 后作为分组键：

```python
round(float(font_size) * 4) / 4
```

而 `src/layout/text.py::_extract_text_characters` 把**字形外框高度**当作字号：

```python
height = max(abs(y1 - y0), 1.0)
```

数学字体的逐字形外框高度差异极大——实测同一行的取值跨越 1.1—6.9pt。量化后每个字符落入不同档位，于是**每个字符各自成为一个片段**：一行 111 字符得到 76 个片段。

### 2.2 导出按片段逐个绝对定位

`src/export/fidelity.py::_place_text_block` 对每个片段单独调用 `_place_text_span`，框的纵坐标由该片段自身的包围盒推算。同一基线上相差 0.1pt 的片段因此落在不同高度，字形互相覆盖。

## 3. 已实施的修复

`src/layout/text.py` 改为以**整行中位字号**作为固定参照，字号只在相对偏差超过按绝对尺度分档的容差时才断开片段：

```python
tolerance = (0.6 if line_size < 4.0 else 0.35) * line_size
```

同时删除了失效的 `style_key`（`src/fonts/resolver.py`）与 `_TextCharacter.style_key`（`src/layout/models.py`）。

**效果**：`GPT-3` 样本的成品穿插由 8 行降为 **0 行**。

## 4. 导出层不变量受阻

目标要求"源文件的一条文字行对应一个定位框，行内不同字体与字号用文字片段表示"。在导出层实现该规则共尝试五种方案，全部未通过基线门禁：

| 方案 | 匹配行数 | bbox 中位 | bbox 最大 | SSIM 均值 |
|---|---|---|---|---|
| 基线（逐片段定位） | 275 | 0.37 | 13.45 | 0.9603 |
| 整行一框，字号取中位 | 268 | — | 20.28 | 0.9322 |
| 字号取 90 分位 | 268 | — | 20.28 | 0.9322 |
| 补 `pdf_font_name` 以保留字体替换 | 262 | — | 30.96 | 0.9309 |
| 字体尺寸偏差超过 25% 的片段单独定位 | 262 | — | 30.96 | 0.9309 |
| 锚点片段定位 + 整行文字合并 | 265 | — | 301.13 | 0.9152 |
| 一框多运行（每个片段独立字号） | **279** | **4.68** | 30.00 | 0.9330 |

最后一种方案在匹配行数上有改善（275 → 279），但 bbox 中位误差由 0.37pt 恶化到 **4.68pt**：Word 对同一段落内的多个运行按文字流排布，无法保留每个片段在源 PDF 中的精确水平位置。

**因此存在原理性冲突**：逐片段绝对定位保证水平精度，整行一框保证同一基线不被拆散，二者不可兼得。要同时满足，需要让 Word 在保留逐片段水平位置的前提下接受一个段落——这超出 `w:framePr` 的表达能力。

## 5. 当前验收状态

| 编号 | 验收项 | 状态 |
|---|---|---|
| 1 | 版面行数与定位框数不变量 | 已建立，但发现两者单位不同（行与元素），已改用字形级判据 |
| 2 | 字形穿插检测器对全部样本归零 | **6 / 7 通过**，`Attention` 余 4 行 |
| 3 | 三个合成受控用例 | **3 / 3 通过**（用例使用标准字体，不触发本缺陷） |
| 4 | 语料影响面 | 113 个 PDF 扫描：16 个存在行拆分，均在 5 行以内 |
| 5 | 基线与单测 | **87 个测试通过**，first20 与 F103 关键指标与基线一致 |

字形穿插检测器的判据已排除两类假阳性：同坐标多码位（数学斜体字体的映射重复）与连字紧排；判定改为**与源文件对比**，成品穿插行数不得多于源文件。

## 6. 剩余缺陷的定位

`Attention` 余下 4 行穿插**不是导出层造成的**：版面模型本身已经产出交错文本。第 6 页该区域的三条记录为：

```text
y=155.79  x=124.72-181.55  'RCSeeolcnfuv-Arorlteutnettin otnioanl'
y=156.24  x=258.54-309.19  'OOO(k(( nn· 2n· ·d· d2d))2 )'
y=156.63  x=351.04-463.81  'OOO(((n11)))' / 'O(lOOog((nk1())n))'
```

三条都是同一视觉行、纵坐标相差不超过 0.5pt 的兄弟行。直接检查源 PDF 的字符抽取结果，同一高度上的字符串同样已经交错：

```text
top≈141  'SlfAiOdO1O1'    x 范围 125.0 ~ 443.6
top≈142  'ttt'            x 范围 151.2 ~ 166.2
top≈143  'eenonn'         x 范围 130.3 ~ 276.4
```

`'SlfAiOdO1O1'` 由 `Self-Attention` 与 `O(1)` 逐字符交替而成，`'()()()'` 由三个 `O(...)` 的括号交替而成。即公式表格里多个单元格的字符被按横坐标排序后混成了同一串，**版面模型自身产生了交错文本**，导出层只是如实写出。

这一处与本文第 2 节记录的缺陷成因不同：第 2 节是同一段文字被拆成逐字符片段，这里是同一视觉行内**不同单元格**的字符被混成一行。两者需要分别处理。

## 7. 剩余缺陷的根因（已定位到字符级）

`Attention` 的剩余穿插在**分带**阶段产生。`src/layout/text.py::_extract_text_lines` 把字符按 `(top, x0)` 排序后逐字归入文字带，归带判据 `_same_text_band` 为：

```python
overlap = min(left.bottom, right.bottom) - max(left.top, right.top)
minimum_height = min(left.bottom - left.top, right.bottom - right.top)
if overlap >= minimum_height * 0.35:
    return True
return abs(left.center_y - right.center_y) <= max(
    6.0,
    max(left.font_size, right.font_size) * 0.6,
)
```

空格等字符的外框**高度为 0**，于是 `overlap` 为 0、`minimum_height` 为 0，判据落到兜底的固定 6.0pt 容差。实测该页某个空格字符：

```text
' '  x=181.55  top=147.74  bottom=147.74  高=0.00
'R'  x=124.72  top=152.53  bottom=159.12  高=6.60   ← 下一行的首字符
```

两者纵向中心相差 **4.76pt < 6.0**，因此**下一行被并进同一带**。带内字符再按横坐标排序，就产生了交错文本：

```text
y=155.79  'RCSeeolcnfuv-Arorlteutnettin otnioanl'
y=156.24  'OOO(k(( nn· 2n· ·d· d2d))2 )'
y=156.63  'OOO(((n11)))' / 'O(lOOog((nk1())n))'
```

### 尝试的修复与取舍

围绕"零高度字符不得把相邻行并入同一带"试了四种改法，实测数据如下。

| 方案 | attention 交错 | first20 匹配行 | bbox 中位 / 最大 | 单元测试 |
|---|---|---|---|---|
| 原逻辑（未改动） | 4 行 | **275** | 0.37 / 13.451 | **87 通过** |
| 参考基准改用带内最高字符 | 已修复 | 264 | — | — |
| 探测对象替换为带内最高字符 | 已修复 | 274 | 0.37 / 13.451 | 3 项失败 |
| 原判据或探测判据任一成立 | 4 行 | **275** | 0.37 / 13.451 | **87 通过** |
| 双方之一零高度时按另一侧高度收紧容差 | 已修复 | 264 | — | 1 项失败 |
| 归带比较基准跳过零高度字符 | 已修复 | 264 | — | **87 通过** |

**共同规律**：凡改变归带结果的方案，`Attention` 的交错都能修掉，但实验指导书前 20 页的匹配行数一律由 275 降到 264（少 11 行）。原因是该文档里零高度字符同样是行内正常成分，收紧或跳过它都会把正常行拆开。

因此这些方案都**不够外科**，已全部撤回；工作区保持全部门禁通过的状态（87 测试、first20 与 F103 与基线一致）。

### 下一步方向

已确认固定 6.0pt 容差与字号不匹配是根本问题。实测把容差改为**随字号缩放**（`字号 × 0.75`）后：

| 指标 | 基线 | 改为缩放容差 |
|---|---|---|
| `Attention` 交错 | 4 行 | **0 行** |
| first20 匹配行 | 275 | **276**（优于基线） |
| first20 未匹配行 | 18 | **17**（优于基线） |
| first20 bbox 中位 / 最大 | 0.37 / 13.451 | **0.37 / 13.451**（完全一致） |
| first20 SSIM 均值 / 最低 | 0.9603 / 0.8866 | **0.9603 / 0.8866**（完全一致） |
| F103 | 与基线一致 | **与基线一致** |
| 单元测试 | 87 通过 | 1 项失败 |

唯一未通过的是 `test_repeated_header_footer_is_marked_and_filtered`：页眉 `Quarterly Report - Internal` 的**词间空格脱离所属行**，被 `_compact_text` 丢弃，变成 `QuarterlyReport-Internal`。

该空格字符的实测几何为：

```text
' '  x=73.01  top=24.99  bottom=25.00  高=0.01  字号=1.00  cy=25.00
'y'  x=68.65  top=20.33  bottom=26.88  高=6.55  字号=6.55  cy=23.61
'R'  x=76.22  top=18.56  bottom=25.00  高=6.44  字号=6.44  cy=21.78
```

空格与前一字母 `'y'` 的中心差为 **1.39pt**，而空格自身字号只有 1.00，因此以空格为一方计算尺度时容差不足；既有逻辑之所以接受它，是因为比较基准 `band[-1]` 是前一个**字母**（字号 6.55）而不是空格。

结论：容差尺度必须由**参与归带判定的字形**提供，而外框退化字符自身无法提供该尺度。下一步需要在分带循环中显式区分"外框退化字符的归属"与"正常字形的归带判定"——例如让退化字符无条件继承前一个正常字形的归带结果，而不是走通用判据。这一改动很小，预期可同时满足 attention 修复、first20 指标优于基线、以及 87 个单测全部通过。

## 8. 已实施的修复（第 7 轮）

在容差随字号缩放的基础上，补上“外框退化字符继承前一个正常字形的归带结果”，两处改动如下。

src/layout/models.py::_same_text_band 的纵向容差由固定下限改为随字号缩放：

`python
scale = max(left.font_size, right.font_size)
if scale <= 0:
    return False
return abs(left.center_y - right.center_y) <= scale * 0.75
`

src/layout/text.py::_extract_text_lines 增加退化字符的归属规则：

`python
if matching_band is None:
    if _is_degenerate_height(character) and bands:
        bands[-1].append(character)
        continue
    bands.append([character])
`

_is_degenerate_height 以外框高度不超过 0.5pt 为判据：这类字符（部分空格）自身字号也随之退化，无法提供有效的纵向尺度，直接跟随前一个正常字形，既不独自成带被丢弃，也不会把相邻行并入同一带。

### 验证结果

| 指标 | 基线 | 修复后 |
|---|---|---|
| Attention 成品穿插行 | 4 | **1** |
| GPT-3 成品穿插行 | 0（第 3 轮修复） | **0** |
| ResNet / NIST / NASA / 实验指导书 / F103 穿插行 | 0 | **0** |
| first20 page_delta | 0 | **0** |
| first20 匹配行 | 275 | **276** |
| first20 未匹配行 | 18 | **17** |
| first20 bbox 中位 / 最大 | 0.37 / 13.451 | **0.37 / 13.451** |
| first20 SSIM 均值 / 最低 | 0.9603 / 0.8866 | **0.9603 / 0.8866** |
| F103 全部指标 | 与基线一致 | **与基线一致** |
| 单元测试 | 87 通过 | **87 通过** |

first20 的 bbox 中位、最大误差、SSIM 均值与最低值四项与基线**完全一致**，匹配行数优于基线，仅 p90 由 5.42 变为 5.46（差 0.04pt）。

### 剩余 1 处穿插

Attention 第 11 页（参考文献页）仍有 1 行被判定穿插。该处文本已正常（'[24]Minh-ThangLuongHieuPhamandChristopherDMann'），命中来自矮标点字符：,、. 的外框高度在 0.5—2.41pt，略高于退化阈值，因此走通用判据；其字形区间与相邻字符轻微交叠。该判定属判据灵敏度边界，尚未确认是真实缺陷还是检测器噪声。

## 9. 第 8 轮补充修复与最终状态

第 8 轮发现退化字符的归属不能笼统地交给“上一条带”。Attention 第 11 页有 **384 个空格，外框高度全部为 0.00、字号为 1.00**，若一律追加到 ands[-1]，会把空格挂到无关的行上，词间空格随之丢失。改用**按纵向位置找所属带**：

`python
host = next(
    (
        band
        for band in reversed(bands)
        if min(item.top for item in band) <= character.center_y
        <= max(item.bottom for item in band)
    ),
    None,
)
`

修复后该页词间空格恢复，例如：

| 修复前 | 修复后 |
|---|---|
| 'acrosslanguages.InProceedingsofthe2009C' | 'across languages. In Proceedings of the 2009 C' |

### 最终验收状态

| 指标 | 基线 | 当前 |
|---|---|---|
| first20 page_delta | 0 | **0** |
| first20 匹配行 | 275 | **276** |
| first20 未匹配行 | 18 | **17** |
| first20 bbox 中位 | 0.37 | **0.37** |
| first20 bbox 最大 | 13.451 | **13.451** |
| first20 SSIM 均值 | 0.9603 | **0.9603** |
| first20 SSIM 最低 | 0.8866 | **0.8866** |
| F103 | 与基线一致 | **与基线一致** |
| 单元测试 | 87 | **87 通过** |
| 导入 / 命名检查 | — | 41 模块 0 问题 / 0 可疑命名 |
| 字形穿插（7 个样本） | 4 行（attention） | **1 行（attention 第 11 页）** |

### 剩余 1 行的性质

Attention 第 11 页仍有 1 行被判定穿插，来源是两条相邻的参考文献行：

`	ext
top=373.37  bot=382.33  '[14]ZhongqiangHuangandMaryHarperSelf-trainingPCFGgrammars...'
top=384.28  bot=393.24  'across languages. In Proceedings of the 2009 Conference on...'
`

两行的版面纵向范围并不交叠（373—382 与 384—393），但渲染结果的字形位置成对偏移约 2pt 并互相交叠（'a' 225.93-229.97 与 'M' 227.74-236.22），说明写出时这两行的纵坐标被压到了同一位置。行内容本身已正常，属纵向定位问题，尚未定位到具体代码位置。

## 10. 第 9 轮：剩余 1 行的定位证据

对 Attention 第 11 页的 DOCX 逐框检查，得到该处的纵坐标分布：

`	ext
y= 372.35 x= 108.40 | '[14] Zachroonsgsq laiannggua Hgueasn.g In a n'
y= 372.35 x= 220.25 | 'd'
y= 372.55 x= 227.55 | 'M'
y= 372.55 x= 252.10 | 'H'
y= 373.40 x= 305.85 | 'tr'
...
`

该节共 322 个定位框。**y≈372 处的一组框里混入了两条不同行的内容**：版面模型中该处有两条相邻行：

`	ext
top=373.37  bot=382.33   '[14] Zhongqiang Huang and Mary Harper, Self-training PCFG grammars...'
top=384.28  bot=393.24   'across languages. In Proceedings of the 2009 Conference on...'
`

两条行的版面纵向范围**并不交叠**（373—382 与 384—393），但导出后它们的文字被写到了同一个纵坐标附近，字形位置成对偏移约 2pt 并互相交叠。该页**不含表格**（len(page.tables) == 0），因此不是表格单元格路径。

已排除的原因：不是逐字符片段拆分（第 3 轮已修）、不是分带把两行并入同一带（分带结果中两行仍各为一条）。

尚未定位：这两条行的框为何落在同一纵坐标。可能方向是 _extract_text_characters 按 (top, x0) 排序后，某条行的字符归属被判给了另一条行，从而使 _split_text_band 在同一条带内产出交错文本，再被 _compact_text 压成一行。

### 当前完整状态

| 指标 | 基线 | 当前 |
|---|---|---|
| first20 page_delta | 0 | **0** |
| first20 匹配行 | 275 | **276** |
| first20 未匹配行 | 18 | **17** |
| first20 bbox 中位 / 最大 | 0.37 / 13.451 | **0.37 / 13.451** |
| first20 SSIM 均值 / 最低 | 0.9603 / 0.8866 | **0.9603 / 0.8866** |
| F103 | 与基线一致 | **与基线一致** |
| 单元测试 | 87 | **87 通过** |
| 字形穿插（7 样本） | 4 行 | **1 行** |

## 11. 第 12 轮：改为按词分组

第 11 轮确认导出后残留区域的片段仍为 1—2 个字符，与最初缺陷同源。第 12 轮改用**按词分组**：同一词内的字符无论外框高度如何都属同一片段，字号差异只在跨词处才断开。

src/layout/text.py::_same_font_run 增加词界判定，_build_text_spans 传入词内间隔阈值：

`python
word_gap = max(line_size * 0.25, 0.4)
...
if character.x0 - previous.x1 < word_gap:
    return True
`

### 效果

Attention 第 11 页的片段粒度由逐字（1—3 字符）变为**整词**：

`	ext
修复前：'ro' 'ce' 'e' 'ng' 's' 'o' 'e' ...
修复后：'across' 'languages.' 'In' 'Proceedings' 'of' 'the' ...
`

第 11 页片段数由逐字级降至平均 **6.6 个/行**。

### 验收状态

| 指标 | 基线 | 当前 |
|---|---|---|
| first20 page_delta | 0 | **0** |
| first20 匹配行 / 未匹配行 | 275 / 18 | **276 / 17** |
| first20 bbox 中位 / 最大 | 0.37 / 13.451 | **0.37 / 13.451** |
| first20 SSIM 均值 / 最低 | 0.9603 / 0.8866 | **0.9603 / 0.8866** |
| F103 | 与基线一致 | **与基线一致** |
| 单元测试 | 87 | **87 通过** |
| 导入 / 命名检查 | — | 41 模块 0 问题 / 0 可疑命名 |
| 字形穿插（7 样本） | 4 行 | **1 行** |

### 未解决

Attention 第 11 页仍有 1 行被判定穿插。导出后该处有 **50 个定位框，每个仅 1—2 个字符**（'d' 'h' 'al' 'l' 'f' '2' '0' '0'），与版面模型中的整词片段不一致——说明这些逐字框并非来自当前版面模型的行片段，而是另有来源，尚未定位。

已排除：表格路径（该页无表格）、分带合并（55 条带，无一带同时含这两行）、检测器噪声（字形 x 区间确实交叠）。

## 12. 第 13 轮：片段样式来源修正

第 13 轮找到残留区域片段异常的第三个成因：**退化字符污染片段样式**。

_build_text_spans 原先按字符数取代表字符（max(group, key=len(text))）。退化空格同样只有一个字符，长度相同时会被选中，于是整个片段的代表字号变成空格的 ont_size = 1.00，片段 bbox 也按空格收紧。实测该页多处片段字号被写成 1.00：

`	ext
'In'  bbox=(205.3, 384.5, 213.7, 391.1)  字号=1.00   ← 实际高度 6.6
'of'  bbox=(269.7, 384.3, 279.0, 393.1)  字号=1.00
'on'  bbox=(367.8, 386.7, 377.7, 391.2)  字号=1.00
`

修正为：代表字符与 bbox 均只在**外框正常**的字符中选取。

`python
styled = [item for item in group if not _is_degenerate_height(item)] or group
representative = max(styled, key=lambda item: len(item.text))
`

修正后同一批片段的字号恢复正常：

`	ext
'In'  字号=6.60   'of'  字号=4.49   'the'  字号=5.54
`

### 效果

| 指标 | 基线 | 第 12 轮 | 第 13 轮 |
|---|---|---|---|
| first20 匹配行 | 275 | 276 | **276** |
| first20 未匹配行 | 18 | 17 | **17** |
| first20 bbox 中位 / 最大 | 0.37 / 13.451 | 0.37 / 13.451 | **0.37 / 13.451** |
| first20 SSIM 均值 / 最低 | 0.9603 / 0.8866 | 0.9603 / 0.8866 | **0.9612 / 0.8867**（优于基线）|
| F103 | 与基线一致 | 与基线一致 | **与基线一致** |
| 单元测试 | 87 | 87 | **87 通过** |
| 字形穿插 | 4 行 | 1 行（第 11 页 y≈411） | **1 行（第 11 页 y≈514）** |

原先 y≈411 处的穿插已消除，但第 11 页另一处（y≈514）出现同类穿插。该处涉及两条相邻行：

`	ext
top=507.12  '[18] Nal Kalchbrenner, Lasse Espeholt, Karen Simonyan, Aaron van den Oord...'
top=518.00  'ray Kavukcuoglu. Neural machine translation in linear time...'
`

字形位置确认互相交叠（'t' 223.16-225.84 与 'Z' 223.74-229.42 与 'i' 226.12-228.35），属真实缺陷而非检测噪声。词本身已完整，属纵向定位问题。

## 13. 第 14 轮：渲染图确认残留缺陷的形态

把 Attention 第 11 页 y≈514 区域渲染出来直接查看，得到决定性证据：

`	ext
[10] A l𝑒𝑥  Graves      , G 2e0n1e3r.ating sequences with recurrent neural networks.arXivv preprint
     𝑎𝑟𝑋𝑖𝑣:1308.0850
[11] Kagaeim reinco gg Hni𝑒𝑗𝑍ℎ𝑎𝑋𝑛𝑎𝑛𝑔. Sℎ𝑎𝑛𝑔𝑖𝑛𝑔 Ren, and Jian Sun. Deep residual learning for im
     𝑎𝑟𝑋𝑖𝑣𝑃𝑟𝑜𝑐𝑒𝑒𝑑𝑖𝑛𝑔𝑠 of the IEEE Conference on Computer Vision and Pattern
     Recognition, pages 770–778, 2016.
`

[10] 条目渲染正常，[11] 条目出现**文字重叠加倍**：同一位置叠加了两层文字，字距被压缩。该页 DOCX 侧只有 3 个定位框，且位置正确：

`	ext
y= 506.10 x= 108.40 | '[18] Nal Kalchbrenner, Lasse Espeholt, Karen'
y= 517.00 x= 129.15 | 'ray Kavukcuoglu. Neural machine translation '
y= 517.00 x= 364.55 | 'arXivp reprint arXiv:1610.10099v2'
`

即版面模型与 DOCX 都正常，**损坏发生在渲染阶段或 DOCX 中尚未检查到的部分**。字形几何确认互相交叠（'t' 223.16-225.84、'Z' 223.74-229.42、'i' 226.12-228.35），并非检测器噪声。

### 累计状态

| 指标 | 基线 | 当前 |
|---|---|---|
| first20 page_delta | 0 | **0** |
| first20 匹配行 / 未匹配行 | 275 / 18 | **276 / 17** |
| first20 bbox 中位 / 最大 | 0.37 / 13.451 | **0.37 / 13.451** |
| first20 SSIM 均值 / 最低 | 0.9603 / 0.8866 | **0.9612 / 0.8867** |
| F103 | 与基线一致 | **与基线一致** |
| 单元测试 | 87 | **87 通过** |
| 字形穿插（7 样本） | 4 行 | **1 行** |

## 14. 第 15 轮：残留缺陷的精确定位

对比源 PDF 与成品在第 11 页 [11] 条目的文本，得到决定性证据。

源 PDF 的该条目完全正常：

`	ext
'[11] Kaiming He, Xiangyu Zhang, Shaoqing Ren, and Jian Sun. Deep residual learning for im\ufffeage recognition. In Proceedings of the IEEE Conference on Computer Vision and Pattern\r\nRecognition, pages 770–'
`

成品同一位置的文本被打乱：

`	ext
'[11] Kagaeim reincogg Hnieti,o Xni.an Ing y\r\nP\nu\r\nro\r\nZ\r\nc\r\nh\r\ne\r\na\r\ne\r\nn\r\nd\r\ng\r\nin,'
`

三个判定要点：

1. Kagaeim 是 Kaiming 的**字符重排**（K-a-i-m-i-n-g 被打乱为 K-a-g-a-e-i-m），该词在源 PDF 中不存在；
2. 后半段字符被**逐字拆行**（P u 
o Z c h e … 各自成行）；
3. 同页的 [18] 条目完整正常（'[18] Nal Kalchbrenner, Lasse Espeholt, Karen Simonyan, Aaron van den Oord, Alex Graves, and Ko\uffferay Kavukcuoglu. Neural machi'）。

因此该残留**不是源文件缺陷，也不是检测器噪声，而是导出阶段对这一段文本的字符重排**。

进一步线索：源文本含 \ufffe（行内断行标记）与 \r\n，而损坏的条目恰好是含 \ufffe 的那几条（[9]、[11] 均含，[10]、[12] 不含或位置不同）。方向指向 _split_text_band 与 _build_text_spans 在处理含 \ufffe 的行时，字符顺序被破坏。

### 累计成效

| 指标 | 基线 | 当前 |
|---|---|---|
| first20 page_delta | 0 | **0** |
| first20 匹配行 / 未匹配行 | 275 / 18 | **276 / 17** |
| first20 bbox 中位 / 最大 | 0.37 / 13.451 | **0.37 / 13.451** |
| first20 SSIM 均值 / 最低 | 0.9603 / 0.8866 | **0.9612 / 0.8867** |
| F103 | 与基线一致 | **与基线一致** |
| 单元测试 | 87 | **87 通过** |
| 字形穿插（7 样本） | 4 行 | **1 行** |

## 15. 第 16 轮：根因确定为字符索引错位

对比同一段文本在版面模型与成品中的形态，定位到最终根因。

**版面模型完全正常**（Kaiming 及全部词正确）：

`	ext
'[11] Kaiming He, Xiangyu Zhang, Shaoqing Ren, and Jian Sun. Deep residual learning for im'
片段：'[' '1' '1' ']' 'Kaiming' 'He,' 'Xiangyu' 'Zhang,' 'Shaoqing' 'Ren,' ...
`

**成品被打乱**（Kaiming → Kagaeim，字符逐字拆行）。

问题在 src/layout/text.py::_extract_text_characters：该函数以**文本码位索引**调用 	ext_page.get_charbox(index)，而 pdfium 的索引是**字符索引**。两者在正常情况下一致，但当一个索引产出多个字符时（连字、数学字符、行内断行标记 \ufffe）便发生错位，字符的坐标与文本张冠李戴，表现为字符重排与逐字拆行。

含 \ufffe 的条目（[9]、[11]）出现该现象，不含的条目（[10]、[12]）正常，与该解释一致。

**修法**：循环按每个索引实际产出的字符数推进（index += len(raw_text)），而不是逐码位推进。已尝试该改法，但直接替换会使 13 项既有测试失败——因为 _compact_text 会剥掉首尾空格，产出宽度与原始宽度不再一致。正确改法是保留原始 
aw_text 的宽度用于推进索引、仅对文本内容做归一化，需要在不改变既有归带与分组行为的前提下重构该循环。

该改动已按项目规范撤回，工作区保持全部门禁通过。

## 16. 第 20 轮：字符抽取改用 pdfium 字符索引（已实施）

### 根因确认

`pypdfium2` 的 `get_text_range(i, i+1)` **第二个参数不是长度**，调用它会返回从该位置到结尾的整段文本（实测 `count_chars=3251` 时拼接结果达 200 万字符）。而 `get_charbox(index)` 使用的是**字符索引**。两者分属不同索引体系，在连字与 `\ufffe` 断行标记处错位，导致文本与字符盒指向不同字符——这正是"字符重排 + 逐字拆行"的来源。

验证：直接调用 pdfium 原始 API 按字符索引取码位，结果完全正确：

```text
raw:  '[5] Kyunghyun Cho, Bart van Merrienboer, Caglar Gulcehre, Fethi Bougares, Holger'
Kaiming@961: '0, 2013.\r\n[11] Kaiming He, Xiangyu Zhang, Shaoqing Ren, and Jian Sun. Deep '
长度 3251 = count_chars
```

### 已实施的改动

`src/layout/text.py::_extract_text_characters` 改为按字符索引迭代，新增两个辅助函数：

```python
def _character_count(text_page) -> int:
    return int(pdfium_raw.FPDFText_CountChars(text_page.raw))

def _character_at(text_page, index) -> str:
    code = int(pdfium_raw.FPDFText_GetUnicode(text_page.raw, index))
    ...
```

同时补上 XML 控制字符过滤——原始 API 会返回 C0 控制字符，直接写入 DOCX 会触发写入库校验失败（本轮实测该失败）。

### 验证结果：偏差收敛到导出层

| 层 | `Kaiming` | 结论 |
|---|---|---|
| 字符抽取 | 正确 | 本轮修复生效 |
| 版面模型 | 正确 | 修复生效 |
| IR（导出前） | 正确 | 修复生效 |
| **DOCX（导出后）** | **`Kagaeim`** | **导出层仍引入偏差** |

版面与 IR 均为 `'[11] Kaiming He, Xiangyu Zhang, Shaoqing Ren, and Jian Sun. Deep residual learning for im'`，而 DOCX 变为交错文本：

```text
'[11] Kagaeim reincogg Hnieti,o Xni.an Ing y\r\nP\r\nu\r\nro\r\nZ\r\nc\r\nh\r\ne\r\na\r\ne\r\nn\r\nd\r\ng\r\nin,'
```

下一轮应从 `export_fidelity_docx` 的 `_place_text_block` 入手，检查它写入的片段是否来自正确的行。

### 门禁状态

| 指标 | 基线 | 当前 |
|---|---|---|
| first20 `page_delta` | 0 | **0** |
| first20 匹配行 / 未匹配行 | 275 / 18 | **276 / 17** |
| first20 bbox 中位 / 最大 | 0.37 / 13.451 | **0.37 / 13.451** |
| first20 SSIM 均值 / 最低 | 0.9603 / 0.8866 | **0.9612 / 0.8867** |
| F103 | 与基线一致 | **与基线一致** |
| 单元测试 | 87 | **87 通过** |
| 导入 / 命名检查 | — | 41 模块 0 问题 / 0 可疑命名 |
| 字形穿插（7 样本） | 4 行 | **1 行** |

## 17. 第 21 轮：空格在片段拼接时丢失

本轮在追查"IR 行文本正确、DOCX 错乱"的过程中，发现一个独立且更基础的事实。

### 观测

IR 中该行的**行文本含有空格**，但**片段拼接不含空格**：

```text
行文本  : '[11] Kaiming He, Xiangyu Zhang, Shaoqing Ren, and Jian Sun. Deep residual learni'
片段拼接: '[11]KaimingHe,XiangyuZhang,ShaoqingRen,andJianSun.Deepresiduallearningforim'
```

导出层使用的是**片段**，因此 DOCX 继承的是丢空格的结果。

### 已排除的原因

字符抽取环节完全正确——`FPDFText_GetUnicode` 报告该页 **384 个空格字符**，`_extract_text_characters` 的输出经核对也保留了全部 384 个空格字符对象。空格是在**片段分组**阶段丢失的。

同时发现该页空格的横坐标与相邻词**重叠**：

```text
']' x=121.6
' ' x=124.6
'K' x=129.9
...
' ' x=179.9
'C' x=180.2   ← 空格与 'C' 相差 0.3
```

### 尝试与结果

按"空格处必须断开片段"修改 `_same_font_run`（空格归入其前面的片段），实测**未修复空格丢失**（片段拼接仍为 0 个空格），且使 4 项既有测试失败，已按项目规范撤回。

该结果说明空格丢失发生在 `_same_font_run` 之外的环节，可能是 `_split_text_band` 的分组、或 `_compact_text` 对 `"".join(...)` 结果的归一化。下一轮应从 `_build_text_spans` 中 `span_text = _compact_text("".join(item.text for item in group))` 这一行入手，检查 group 内是否真的包含空格字符对象。

### 门禁状态

| 指标 | 基线 | 当前 |
|---|---|---|
| first20 `page_delta` | 0 | **0** |
| first20 匹配行 / 未匹配行 | 275 / 18 | **276 / 17** |
| first20 bbox 中位 / 最大 | 0.37 / 13.451 | **0.37 / 13.451** |
| first20 SSIM 均值 / 最低 | 0.9603 / 0.8866 | **0.9612 / 0.8867** |
| F103 | 与基线一致 | **与基线一致** |
| 单元测试 | 87 | **87 通过** |
| 字形穿插（7 样本） | 4 行 | **1 行** |

## 18. 第 22 轮：空格丢失的根因定位到代码行，修法受阻

本轮把"片段拼接丢空格"定位到**确切的代码位置**，但三种修法均引入回归，已全部撤回。

### 根因（已用最小用例证明）

`_build_text_spans` 第 560 行：

```python
span_text = _compact_text("".join(item.text for item in group))
```

空格被分到了**后一个片段的开头**。实测分组内容：

```text
组3: '] '        空格在末尾（正确）
组5: ' He,'      空格在开头（错误）
组6: ' Xiangyu'  空格在开头（错误）
```

而 `_compact_text` 会剥掉片段首尾空格，于是 `' He,'` → `'He,'`，**词间空格全部丢失**。

最小复现：

```python
_build_text_spans(['a', ' ', 'b'])  →  'ab'   （期望 'a b'）
_build_text_spans([']', ' ', 'K'])  →  ']K'
```

根因是两条规则的叠加：分组时空格留给后一组的开头 + 归一化时剥掉首尾空格。

### 尝试的三种修法及结果

| 修法 | 最小用例 | 单元测试 | 结论 |
|---|---|---|---|
| 空格归入前一片段（`_same_font_run`） | 未改善 | 4 项失败 | 撤回 |
| 片段文本直接拼接（绕过 `_compact_text`） | `'a b'` ✓ | 5 项失败 | 撤回 |
| 新增 `keep_trailing_space` 参数只作用于片段 | `'a b'` ✓ | 5 项失败 | 撤回 |

失败模式明确：保留尾随空格后，`normalize_page_text` 会把尾随空格视作行分隔，导致**每个词各占一行**（`'Worker \nprocess \npage \none...'`），进而破坏 5 项既有测试。

**因此修法必须是"把空格归入前一片段的末尾、同时不改变 `_compact_text` 的行处理语义"**。第二、三种尝试表明直接在片段拼接处保留空格会触发 `normalize_page_text` 的行分隔行为；第一种尝试表明只改 `_same_font_run` 不足以让空格真正落在前一组。

下一轮应从 `_same_font_run` 的调用顺序入手：`_build_text_spans` 按 `x0` 排序后逐字判定，而空格与相邻词的 x 坐标重叠（实测 `' '` x=179.9 与 `'C'` x=180.2 相差 0.3），排序结果决定空格落入哪一组。需要先确认排序后的实际顺序，再决定归组规则。

### 门禁状态

| 指标 | 基线 | 当前 |
|---|---|---|
| first20 `page_delta` | 0 | **0** |
| first20 匹配行 / 未匹配行 | 275 / 18 | **276 / 17** |
| first20 bbox 中位 / 最大 | 0.37 / 13.451 | **0.37 / 13.451** |
| first20 SSIM 均值 / 最低 | 0.9603 / 0.8866 | **0.9612 / 0.8867** |
| F103 | 与基线一致 | **与基线一致** |
| 单元测试 | 87 | **87 通过** |
| 字形穿插（7 样本） | 4 行 | **1 行** |

## 19. 第 23 轮：空格丢失的两个叠加因素（已全部查清）

### 因素一：空格的横坐标是零宽且位置错位

真实数据中，空格的 `x0` 与 `x1` 相等（零宽），且位置卡在**其前面那个词的位置**：

```text
' ' x0=168.23 x1=168.23   ← 零宽，位置在 'H' 之前
'H' x0=168.43 x1=175.37
'e' x0=175.83 x1=179.88
',' x0=180.65 x1=182.06
' ' x0=185.87 x1=185.87   ← 零宽，位置在 'X' 之前
'X' x0=185.97 x1=193.02
```

`_build_text_spans` 按 `x0` 排序，因此顺序为 `'] He, Xiangyu ...`——**空格排到了词的前面**，而不是它实际分隔的 `He,` 与 `Xiangyu` 之间。

### 因素二：归一化剥离首尾空格

`normalize_page_text` 会剥掉首尾空格：

```text
'He, '     -> 'He,'
'] '       -> ']'
' Kaiming' -> 'Kaiming'
```

因此无论空格落在片段的哪一端，`_compact_text` 都会把它去掉。这解释了第 22 轮三种修法为何全部失败：把空格归入前一片段（成为尾随空格）或被后一片段携带（成为前导空格），都会被这一步剥掉。

### 结论

**片段文本无法承载词间空格**，因为 `normalize_page_text` 的设计目的就是剥掉首尾空白。可行的修法有两条：

1. **不依赖片段文本承载空格**：导出时按片段的 `x0` 与源宽度分别定位，词间距由坐标体现。但第 4 节已实测该路线会破坏基线（逐片段定位的精度与整行一框不可兼得）。
2. **在片段文本中保留空格并让归一化不剥离**：需要为片段路径提供不剥首尾空格的归一化。第 22 轮尝试过该方向，但 `normalize_page_text` 会把尾随空格视作行分隔，导致每个词各占一行，破坏 5 项既有测试。正确做法是**先替换占位、归一化后再还原**，避免触发其行分隔逻辑——该实现需要绕开 `normalize_page_text` 的行处理，属于对 `_compact_text` 的谨慎重构。

### 门禁状态

| 指标 | 基线 | 当前 |
|---|---|---|
| first20 `page_delta` | 0 | **0** |
| first20 匹配行 / 未匹配行 | 275 / 18 | **276 / 17** |
| first20 bbox 中位 / 最大 | 0.37 / 13.451 | **0.37 / 13.451** |
| first20 SSIM 均值 / 最低 | 0.9603 / 0.8866 | **0.9612 / 0.8867** |
| F103 | 与基线一致 | **与基线一致** |
| 单元测试 | 87 | **87 通过** |
| 字形穿插（7 样本） | 4 行 | **1 行** |

## 20. 第 24 轮：占位方案已实现，单测从 5 项失败降到 1 项，但基线回归

### 实现

按第 23 轮确定的方案实现"占位替换 → 归一化 → 还原"：

```python
_SPACE_PLACEHOLDER = "\ue000"

def _compact_span_text(value: str) -> str:
    guarded = value
    if guarded[:1] == " ":
        guarded = _SPACE_PLACEHOLDER + guarded[1:]
    if guarded[-1:] == " ":
        guarded = guarded[:-1] + _SPACE_PLACEHOLDER
    return _compact_text(guarded).replace(_SPACE_PLACEHOLDER, " ")
```

占位字符经 `normalize_page_text` 验证可完整保留（`'He,\ue000'` → `'He,\ue000'`）。

### 结果

| 项目 | 结果 |
|---|---|
| 最小用例 | `'a b'`、`'] K'` —— **空格保留成功** |
| 单元测试 | 失败数由 **5 项降到 1 项** |
| 剩余 1 项 | `test_mixed_font_line_exports_separate_font_frames` 期望旧的 `'Plain'`，实际为 `'Plain '`（带尾随空格）——该失败恰好证明空格修复生效 |
| first20 匹配行 | 275 → **266（回归 9 行）** |
| F103 | 渲染校验失败，指标缺失 |

**结论**：占位方案在功能上正确（空格得以保留、单测接近全绿），但会改变行文本内容，进而影响渲染回读的行匹配——first20 掉 9 行、F103 校验失败。

该改动已按项目规范撤回，工作区保持全部门禁通过。

### 说明

至此已确认：**片段文本保留空格 与 渲染回读基线 存在冲突**。空格保留使片段文本更长/更贴近原文，但回读匹配逻辑对该变化敏感。

若继续推进，需要一并检查渲染回读的匹配逻辑（`src/validate/text_compare.py`）为何对词间空格敏感——该函数以"整行文本相等或公共前缀不少于 4 字符"为匹配条件，片段文本变化会影响其匹配结果。

### 门禁状态（撤回后）

| 指标 | 基线 | 当前 |
|---|---|---|
| first20 `page_delta` | 0 | **0** |
| first20 匹配行 / 未匹配行 | 275 / 18 | **276 / 17** |
| first20 bbox 中位 / 最大 | 0.37 / 13.451 | **0.37 / 13.451** |
| first20 SSIM 均值 / 最低 | 0.9603 / 0.8866 | **0.9612 / 0.8867** |
| F103 | 与基线一致 | **与基线一致** |
| 单元测试 | 87 | **87 通过** |
| 导入 / 命名检查 | — | 41 模块 0 问题 / 0 可疑命名 |
| 字形穿插（7 样本） | 4 行 | **1 行** |
