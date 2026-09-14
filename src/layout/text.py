"""文本字符、字体样式与文本行的提取。

从页面图形对象读出字符盒、字体、字号、颜色与旋转，按基线聚成行，
再按字体 run 切成 span。
"""

from __future__ import annotations
from typing import Any
from .models import (
    PdfTextLine,
    PdfTextSpan,
    _TextCharacter,
    _TextObjectStyle,
    _compact_text,
    _dominant_value,
)
import math
from bisect import bisect_left
from ctypes import c_float, c_int, c_uint
import pypdfium2.raw as pdfium_raw
from ..fonts.resolver import FontMatch, PdfFontDescriptor, resolve_font
from .models import _PAGEOBJ_TEXT

def _raw_color(
    getter: Any,
    target: Any,
) -> tuple[int, int, int] | None:
    red, green, blue, alpha = c_uint(), c_uint(), c_uint(), c_uint()
    try:
        if not getter(target, red, green, blue, alpha):
            return None
    except Exception:
        return None
    if alpha.value == 0:
        return None
    return (
        int(red.value) & 0xFF,
        int(green.value) & 0xFF,
        int(blue.value) & 0xFF,
    )


def _object_fill_color(page_object: Any) -> tuple[int, int, int] | None:
    return _raw_color(pdfium_raw.FPDFPageObj_GetFillColor, page_object)


def _object_stroke_color(page_object: Any) -> tuple[int, int, int] | None:
    return _raw_color(pdfium_raw.FPDFPageObj_GetStrokeColor, page_object)


def _stroke_width(page_object: Any) -> float:
    value = c_float()
    try:
        if not pdfium_raw.FPDFPageObj_GetStrokeWidth(page_object, value):
            return 0.0
    except Exception:
        return 0.0
    return max(float(value.value), 0.0)


def _font_is_bold(font: Any, weight: int, *names: str) -> bool:
    if weight >= 600:
        return True
    try:
        flags = int(pdfium_raw.FPDFFont_GetFlags(font.raw))
    except Exception:
        flags = 0
    if flags & (1 << 18):  # FPDF_FONT_FLAG_FORCEBOLD
        return True
    combined = " ".join(name for name in names if name).lower()
    return any(
        token in combined
        for token in ("bold", "black", "heavy", "semibold", "demibold")
    )


def _font_is_italic(font: Any, *names: str) -> bool:
    try:
        flags = int(pdfium_raw.FPDFFont_GetFlags(font.raw))
    except Exception:
        flags = 0
    if flags & (1 << 6):  # FPDF_FONT_FLAG_ITALIC
        return True
    angle = c_int()
    try:
        if pdfium_raw.FPDFFont_GetItalicAngle(font.raw, angle):
            if abs(int(angle.value)) > 4:
                return True
    except Exception:
        pass
    combined = " ".join(name for name in names if name).lower()
    return "italic" in combined or "oblique" in combined


def _resolve_object_font(
    *,
    base_name: str,
    family_hint: str,
    weight: int | None,
    flags: int | None,
    italic_angle: float | None,
    descriptors: dict[str, PdfFontDescriptor] | None,
    cache: dict[tuple[Any, ...], FontMatch],
    sample_text: str = "",
) -> FontMatch:
    """把 PDF 文本对象解析成本机可用字体。"""
    key = (base_name, family_hint, weight, flags)
    cached = cache.get(key)
    if cached is not None:
        return cached
    descriptor = None
    if descriptors:
        for candidate in (
            base_name,
            base_name.lstrip("/"),
            family_hint,
        ):
            if candidate and candidate in descriptors:
                descriptor = descriptors[candidate]
                break
    match = resolve_font(
        raw_name=base_name,
        family_hint=family_hint,
        descriptor=descriptor,
        text=sample_text,
        font_weight=weight,
        font_flags=flags,
        italic_angle=italic_angle,
    )
    cache[key] = match
    return match


def _extract_object_styles(
    page: Any,
    text_page: Any,
    page_height: float,
    *,
    font_descriptors: dict[str, PdfFontDescriptor] | None = None,
    font_cache: dict[tuple[Any, ...], FontMatch] | None = None,
) -> list[_TextObjectStyle]:
    """提取页面中所有文本对象的字体、颜色和 z-order。"""
    styles: list[_TextObjectStyle] = []
    cache = font_cache if font_cache is not None else {}
    try:
        objects = list(page.get_objects(textpage=text_page))
    except Exception:
        return styles
    for index, page_object in enumerate(objects):
        try:
            if page_object.type != _PAGEOBJ_TEXT:
                continue
            try:
                x0, y0, x1, y1 = (
                    float(value) for value in page_object.get_bounds()
                )
            except Exception:
                continue
            bbox = (
                min(x0, x1),
                page_height - max(y0, y1),
                max(x0, x1),
                page_height - min(y0, y1),
            )
            base_name = ""
            family_hint = ""
            weight: int | None = None
            flags: int | None = None
            italic_angle: float | None = None
            sample_text = ""
            try:
                font = page_object.get_font()
                try:
                    base_name = font.get_base_name() or ""
                except Exception:
                    base_name = ""
                try:
                    family_hint = font.get_family_name() or ""
                except Exception:
                    family_hint = ""
                try:
                    weight = int(font.get_weight())
                except Exception:
                    weight = None
                try:
                    flags = int(pdfium_raw.FPDFFont_GetFlags(font.raw))
                except Exception:
                    flags = None
                try:
                    angle_value = c_int()
                    if pdfium_raw.FPDFFont_GetItalicAngle(font.raw, angle_value):
                        italic_angle = float(angle_value.value)
                except Exception:
                    italic_angle = None
                try:
                    sample_text = page_object.extract()[:64]
                except Exception:
                    sample_text = ""
            except Exception:
                pass
            match = _resolve_object_font(
                base_name=base_name,
                family_hint=family_hint,
                weight=weight,
                flags=flags,
                italic_angle=italic_angle,
                descriptors=font_descriptors,
                cache=cache,
                sample_text=sample_text,
            )
            rotation = 0.0
            matrix_scale = 1.0
            try:
                matrix = page_object.get_matrix()
                matrix_values = list(matrix.get())
                rotation = _matrix_rotation(matrix)
                matrix_scale = _matrix_scale(matrix_values)
            except Exception:
                rotation = 0.0
                matrix_scale = 1.0
            try:
                font_size = float(page_object.get_font_size()) * matrix_scale
            except Exception:
                font_size = max(bbox[3] - bbox[1], 1.0)
            color = _object_fill_color(page_object) or (0, 0, 0)
            styles.append(
                _TextObjectStyle(
                    bbox=bbox,
                    font_name=match.family or family_hint or base_name or "",
                    font_size=max(font_size, 0.1),
                    color=color,
                    bold=match.bold,
                    italic=match.italic,
                    z_order=index,
                    rotation=rotation,
                    pdf_font_name=base_name or family_hint,
                    substituted=match.substituted,
                    fallback_reason=match.fallback_reason,
                )
            )
        finally:
            close_object = getattr(page_object, "close", None)
            if close_object is not None:
                close_object()
    return styles


def _matrix_scale(matrix_values: Any) -> float:
    """从文本矩阵中提取有效缩放，修正 PDF 里被缩放绘制的字号。"""
    try:
        values = list(matrix_values)
    except TypeError:
        return 1.0
    if len(values) < 4:
        return 1.0
    a, b, c, d = (float(value) for value in values[:4])
    determinant = abs(a * d - b * c)
    if determinant <= 1e-9:
        return 1.0
    scale = math.sqrt(determinant)
    if scale <= 1e-6:
        return 1.0
    return max(min(scale, 20.0), 0.01)


def _matrix_rotation(matrix: Any) -> float:
    """从 PDF 文本矩阵中提取旋转角（度）。"""
    try:
        values = list(matrix.get())
    except Exception:
        values = [getattr(matrix, name, 0.0) for name in ("a", "b", "c", "d")]
    if len(values) < 4:
        return 0.0
    a, b = float(values[0]), float(values[1])
    if abs(a) < 1e-6 and abs(b) < 1e-6:
        return 0.0
    angle = math.degrees(math.atan2(b, a))
    if abs(angle) < 0.05 or abs(abs(angle) - 180.0) < 0.05:
        return 0.0
    return round(angle, 2)


def _style_index(
    styles: list[_TextObjectStyle],
    *,
    cell: float = 64.0,
) -> dict[tuple[int, int], list[int]]:
    index: dict[tuple[int, int], list[int]] = {}
    for position, style in enumerate(styles):
        left = int(style.bbox[0] // cell)
        right = int(style.bbox[2] // cell)
        top = int(style.bbox[1] // cell)
        bottom = int(style.bbox[3] // cell)
        for column in range(left, max(right, left) + 1):
            for row in range(top, max(bottom, top) + 1):
                index.setdefault((column, row), []).append(position)
    return index


def _match_style(
    character: _TextCharacter,
    styles: list[_TextObjectStyle],
    index: dict[tuple[int, int], list[int]],
    *,
    cell: float = 64.0,
) -> _TextObjectStyle | None:
    candidates: list[_TextObjectStyle] = []
    seen: set[int] = set()
    for column in range(
        int(character.x0 // cell), int(character.x1 // cell) + 1
    ):
        for row in range(
            int(character.top // cell), int(character.bottom // cell) + 1
        ):
            for position in index.get((column, row), []):
                if position in seen:
                    continue
                seen.add(position)
                candidates.append(styles[position])
    if not candidates:
        return None
    center_x = character.center_x
    center_y = character.center_y
    containing = [
        style
        for style in candidates
        if style.bbox[0] - 1.5 <= center_x <= style.bbox[2] + 1.5
        and style.bbox[1] - 2.5 <= center_y <= style.bbox[3] + 2.5
    ]
    pool = containing or candidates
    return min(pool, key=lambda style: style.area)


def _is_xml_control_character(value: str) -> bool:
    """判断字符是否为 XML 不允许的控制字符。

    XML 1.0 只接受 TAB、换行与回车三类控制字符，
    其余 C0 控制字符写入 DOCX 会触发写入库校验失败。
    """
    if not value:
        return False
    code = ord(value)
    return code < 0x20 and code not in {0x09, 0x0A, 0x0D}


def _extract_text_characters(
    text_page: Any,
    page_height: float,
    *,
    styles: list[_TextObjectStyle] | None = None,
) -> list[_TextCharacter]:
    characters: list[_TextCharacter] = []
    style_index = _style_index(styles) if styles else {}
    text = text_page.get_text_range()
    for index, raw_text in enumerate(text):
        if raw_text in {"", "\r", "\n", "\t", "\ufffe"} or _is_xml_control_character(
            raw_text
        ):
            continue
        if raw_text == " ":
            text = raw_text
        else:
            text = _compact_text(raw_text)
        if not text:
            continue
        x0, y0, x1, y1 = (float(value) for value in text_page.get_charbox(index))
        height = max(abs(y1 - y0), 1.0)
        character = _TextCharacter(
            text=text,
            x0=x0,
            top=page_height - max(y0, y1),
            x1=x1,
            bottom=page_height - min(y0, y1),
            font_size=height,
        )
        if styles:
            style = _match_style(
                character, styles, style_index
            )
            if style is not None:
                character.font_name = style.font_name
                character.color = style.color
                character.bold = style.bold
                character.italic = style.italic
                character.rotation = style.rotation
                character.z_order = style.z_order
                character.pdf_font_name = style.pdf_font_name
                character.font_substituted = style.substituted
                character.font_fallback_reason = style.fallback_reason
                if style.font_size > 0.2:
                    character.font_size = style.font_size
        characters.append(character)
    return characters


def _group_into_rows(
    characters: list[_TextCharacter],
) -> list[list[_TextCharacter]]:
    """把一组字符按视觉行分组，组内按横向排序。

    逐字符比较外框上界不可靠：同一行里字母的上升部、下降部与 x 高度会把
    外框拉开数个点，按 ``top`` 排序会把同行靠右的字排到靠左的字前面，
    成品因此出现"作者姓名互换"这类顺序错乱。

    同一行文字的基线基本重合，因此以字符下界为依据分行：以中位下界估计
    文字基线，容差取半个字高，逐字符并入距离最近的已有行。标点与上标的
    下界偏离基线但仍在容差内，因此不会自成一行。
    """
    if not characters:
        return []
    # 用下界的中位数估计文字基线：句点等标点的下界比文字低，取众数会被它们带偏
    bottoms = sorted(round(item.bottom, 1) for item in characters)
    baseline = bottoms[len(bottoms) // 2]
    # 容忍下降部（g、p、y）与标点把下界推离基线半个字高
    tolerance = max(
        (item.bottom - item.top for item in characters), default=1.0
    ) * 0.6
    ordered = sorted(
        characters,
        key=lambda item: (
            round(abs(item.bottom - baseline), 1),
            item.x0,
        ),
    )
    # 逐字符匹配已有行：标点与下标的基线低于文字基线，只认单一基线会把
    # 它们挤成独立的一行，行内文字因此被切成碎片
    rows: list[list[_TextCharacter]] = []
    baselines: list[float] = []
    for character in ordered:
        if rows:
            distances = [abs(value - character.bottom) for value in baselines]
            nearest = min(range(len(distances)), key=distances.__getitem__)
            if distances[nearest] <= tolerance:
                rows[nearest].append(character)
                continue
        rows.append([character])
        baselines.append(character.bottom)
    for row in rows:
        row.sort(key=lambda item: item.x0)
    return rows


def _merge_fragmented_groups(
    groups: list[list[_TextCharacter]],
    rectangles: tuple[tuple[float, float, float, float], ...],
) -> list[list[_TextCharacter]]:
    """合并纵向对齐且横向紧邻的行矩形分组。

    部分 PDF 把一整行拆成多个文本对象，pdfium 因此给出多个很窄的行矩形
    （同一页脚的不同栏目、同一行上的多个姓名）。这些矩形纵向互相重叠且
    横向间距在字号量级内，属于同一行，需要合并后再按横向次序成行，
    否则页面文字会被切成大量碎片。
    """
    if not any(groups):
        return groups
    order = sorted(
        range(len(rectangles)), key=lambda index: rectangles[index][1]
    )
    merged: list[list[_TextCharacter]] = []
    bounds: list[tuple[float, float]] = []
    extents: list[tuple[float, float]] = []
    for index in order:
        group = groups[index]
        if not group:
            continue
        top, bottom = rectangles[index][1], rectangles[index][3]
        size = max((item.font_size or 0.0) for item in group)
        placed = False
        if merged:
            previous_top, previous_bottom = bounds[-1]
            previous_x0, previous_x1 = extents[-1]
            x0 = min(item.x0 for item in group)
            x1 = max(item.x1 for item in group)
            vertical = min(bottom, previous_bottom) - max(top, previous_top)
            if vertical > 0:
                gap = (
                    x0 - previous_x1
                    if x0 > previous_x1
                    else previous_x0 - x1
                    if x1 < previous_x0
                    else 0.0
                )
                if gap <= max(size, 1.0) * 2.5:
                    merged[-1].extend(group)
                    bounds[-1] = (
                        min(top, previous_top),
                        max(bottom, previous_bottom),
                    )
                    extents[-1] = (
                        min(x0, previous_x0),
                        max(x1, previous_x1),
                    )
                    placed = True
        if not placed:
            merged.append(list(group))
            bounds.append((top, bottom))
            extents.append(_group_extent(group))
    return merged


def _group_extent(group: list[_TextCharacter]) -> tuple[float, float]:
    return (
        min(item.x0 for item in group),
        max(item.x1 for item in group),
    )


def _split_text_band(characters: list[_TextCharacter]) -> list[PdfTextLine]:
    rows = _group_into_rows(characters)
    characters = [item for row in rows for item in row]
    runs: list[list[_TextCharacter]] = []
    current: list[_TextCharacter] = []
    for character in characters:
        if current:
            previous = current[-1]
            gap = character.x0 - previous.x1
            # 词间空格约半个字宽，超过一个字宽即认为是并列的独立文本
            # （同一视觉行内的多个姓名、页脚的不同栏目），拆成两行
            split_gap = max(previous.font_size, character.font_size) * 1.0
            if gap > split_gap:
                runs.append(current)
                current = []
        current.append(character)
    if current:
        runs.append(current)

    lines: list[PdfTextLine] = []
    for run in runs:
        text = _compact_text("".join(item.text for item in run))
        if not text:
            continue
        font_sizes = sorted(item.font_size for item in run if item.font_size > 0)
        font_names = [item.font_name for item in run if item.font_name]
        colors = [item.color for item in run]
        rotations = [item.rotation for item in run if abs(item.rotation) > 0.05]
        spans = _build_text_spans(run)
        dominant_span = max(
            spans, key=lambda item: len(item.text), default=None
        )
        lines.append(
            PdfTextLine(
                text=text,
                x0=min(item.x0 for item in run),
                top=min(item.top for item in run),
                x1=max(item.x1 for item in run),
                bottom=max(item.bottom for item in run),
                font_size=font_sizes[len(font_sizes) // 2] if font_sizes else 0.0,
                font_name=_dominant_value(font_names, ""),
                color=_dominant_value(colors, (0, 0, 0)),
                bold=sum(item.bold for item in run) * 2 >= len(run),
                italic=sum(item.italic for item in run) * 2 >= len(run),
                rotation=_dominant_value(rotations, 0.0),
                z_order=min(item.z_order for item in run),
                spans=spans,
                pdf_font_name=(
                    dominant_span.pdf_font_name if dominant_span else ""
                ),
                font_substituted=any(
                    item.font_substituted for item in run
                ),
                font_fallback_reason=(
                    dominant_span.fallback_reason if dominant_span else ""
                ),
            )
        )
    return lines


def _is_degenerate_height(character: _TextCharacter) -> bool:
    """判断字符外框是否退化到不足以参与纵向与样式判定。

    部分 PDF 的空格等字符外框高度接近 0，字号也随之退化为极小值，
    既不能用于纵向归带，也不能作为片段样式的来源。
    """
    return (character.bottom - character.top) <= 0.5


def _is_point_like(character: _TextCharacter) -> bool:
    """判断字符外框是否只剩一个点。

    多数 PDF 的空格只给出一个点的外框，纵向中心因此落在基线上，
    无法据此判断所属行；这类字符按字符流顺序跟随前一个实心字符。
    句点、逗号等标点虽小仍有实际高度，按外框正常归行。
    """
    return (character.bottom - character.top) <= 0.05


def _same_font_run(
    character: _TextCharacter,
    previous: _TextCharacter,
    group: list[_TextCharacter],
    line_size: float,
    *,
    word_gap: float,
) -> bool:
    """判断字符是否仍属于当前字体 run。

    字体名、粗体、斜体、旋转任一不同即断开。
    字号差异只在跨词处才断开：同一词内的字符无论外框高度如何都属同一片段，
    否则逐字形度量差异会把整行拆成单个字母。
    """
    first = group[0]
    if character.font_name != first.font_name:
        return False
    if bool(character.bold) != bool(first.bold):
        return False
    if bool(character.italic) != bool(first.italic):
        return False
    if abs(float(character.rotation) - float(first.rotation)) >= 2.0:
        return False
    if character.x0 - previous.x1 < word_gap:
        return True
    size = float(character.font_size)
    if size <= 0 or line_size <= 0:
        return True
    tolerance = (0.6 if line_size < 4.0 else 0.35) * line_size
    return abs(size - line_size) <= tolerance


def _build_text_spans(
    characters: list[_TextCharacter],
) -> tuple[PdfTextSpan, ...]:
    """把一行字符按字体与其他排版变化切成连续的 span。"""
    if not characters:
        return ()
    ordered = sorted(characters, key=lambda item: item.x0)
    sizes = sorted(
        float(item.font_size) for item in ordered if float(item.font_size) > 0
    )
    line_size = sizes[len(sizes) // 2] if sizes else 0.0
    # 词内间隔远小于词间空格，用它区分"同一词"与"跨词"
    word_gap = max(line_size * 0.25, 0.4)
    groups: list[list[_TextCharacter]] = []
    current: list[_TextCharacter] = []
    for character in ordered:
        if current and not _same_font_run(
            character,
            current[-1],
            current,
            line_size,
            word_gap=word_gap,
        ):
            groups.append(current)
            current = []
        current.append(character)
    if current:
        groups.append(current)

    spans: list[PdfTextSpan] = []
    for group in groups:
        span_text = _compact_text("".join(item.text for item in group))
        if not span_text:
            continue
        # 外框退化的字符（如零高度空格）字号也随之退化，不能作为片段样式来源
        styled = [item for item in group if not _is_degenerate_height(item)] or group
        representative = max(styled, key=lambda item: len(item.text))
        spans.append(
            PdfTextSpan(
                text=span_text,
                bbox=(
                    min(item.x0 for item in styled),
                    min(item.top for item in styled),
                    max(item.x1 for item in group),
                    max(item.bottom for item in styled),
                ),
                font_name=representative.font_name,
                pdf_font_name=representative.pdf_font_name,
                font_size=representative.font_size,
                color=representative.color,
                bold=representative.bold,
                italic=representative.italic,
                rotation=representative.rotation,
                z_order=min(item.z_order for item in group),
                substituted=representative.font_substituted,
                fallback_reason=representative.font_fallback_reason,
            )
        )
    return tuple(spans)


def _line_rectangles(
    text_page: Any,
    page_height: float,
) -> tuple[tuple[float, float, float, float], ...]:
    """读取 pdfium 自身的行矩形，作为物理行切分的依据。

    返回值为 ``(x0, top, x1, bottom)``。行矩形由 pdfium 在解析页面时生成，
    同一行内跨字体对象会被拆成多个矩形，由后续归并逻辑重新合并。
    """
    try:
        count = int(text_page.count_rects())
    except Exception:
        return ()
    rectangles: list[tuple[float, float, float, float]] = []
    for index in range(max(count, 0)):
        try:
            left, bottom, right, top = text_page.get_rect(index)
        except Exception:
            continue
        x0 = min(float(left), float(right))
        x1 = max(float(left), float(right))
        rectangles.append(
            (
                x0,
                page_height - max(float(top), float(bottom)),
                x1,
                page_height - min(float(top), float(bottom)),
            )
        )
    return tuple(rectangles)


# 行矩形与字符外框的匹配容差：字形外框常略超出所属行矩形。
_LINE_RECT_TOLERANCE = 3.0


def _assign_characters_to_rectangles(
    characters: list[_TextCharacter],
    rectangles: tuple[tuple[float, float, float, float], ...],
) -> list[list[_TextCharacter]]:
    """把字符分配到所属的行矩形，返回与 ``rectangles`` 等长的分组。

    行矩形由 pdfium 按同一文本对象内的行给出，横向范围可能小于行内文字
    的实际范围；因此只按纵向位置取舍：字符纵向中心落在矩形范围内即归入，
    多个矩形同时覆盖时取纵向中心更接近的一个。

    只剩一个点的字符（空格、零高度连字符）纵向中心落在基线上，无法据此
    判断行归属；这类字符按字符流顺序跟随同一行的前一个实心字符，因为行内
    文字在字符流中连续出现，而空格与连字符都紧跟在相邻的文字之后。
    """
    groups: list[list[_TextCharacter]] = [[] for _ in rectangles]
    if not rectangles:
        return groups
    # 矩形按上边界排序后即可用二分定位候选区间，避免逐行扫描
    order = sorted(range(len(rectangles)), key=lambda index: rectangles[index][1])
    tops = [rectangles[index][1] for index in order]
    bottoms = [
        max(rectangles[index][3], rectangles[index][1]) for index in order
    ]
    # 每个分组已覆盖的横向范围，用于把同一矩形内的续接文字并回该组
    spans: list[tuple[float, float] | None] = [None] * len(rectangles)
    previous_group = None
    for character in characters:
        if _is_point_like(character):
            if previous_group is not None:
                previous_group.append(character)
            continue
        center = character.center_y
        position = bisect_left(tops, center - _LINE_RECT_TOLERANCE)
        best_index = None
        best_key: tuple[float, float] | None = None
        covered_index = None
        covered_key: tuple[float, float] | None = None
        for candidate in range(max(position - 1, 0), len(order)):
            index = order[candidate]
            top = rectangles[index][1]
            if top > center + _LINE_RECT_TOLERANCE:
                break
            bottom = bottoms[candidate]
            if bottom < center - _LINE_RECT_TOLERANCE:
                continue
            key = (
                abs((top + bottom) / 2 - center),
                -min(bottom, center + _LINE_RECT_TOLERANCE)
                + max(top, center - _LINE_RECT_TOLERANCE),
            )
            if best_key is None or key < best_key:
                best_key = key
                best_index = index
            span = spans[index]
            if span is not None and span[0] <= character.center_x <= span[1]:
                if covered_key is None or key < covered_key:
                    covered_key = key
                    covered_index = index
        # 同一矩形内可能并列多段文字（同一行上的多个姓名、页脚的多个栏目），
        # 已经写过的横向范围优先续写，避免把连续文字拆到相邻矩形里
        if covered_index is not None:
            best_index = covered_index
        if best_index is not None:
            groups[best_index].append(character)
            previous_group = groups[best_index]
            span = spans[best_index]
            if span is None:
                spans[best_index] = (
                    min(character.x0, character.x1),
                    max(character.x0, character.x1),
                )
            else:
                spans[best_index] = (
                    min(span[0], character.x0, character.x1),
                    max(span[1], character.x0, character.x1),
                )
    return groups


def _lines_in_reading_order(
    lines: list[PdfTextLine],
) -> list[PdfTextLine]:
    """按阅读次序排列文本行，同一视觉行内按横向次序。

    直接按 ``(top, x0)`` 排序时，同一视觉行各段的上界差异（字形高低造成的
    零点几到几个点）会成为主键，使同行靠右的文字排到靠左的文字之前，
    成品因此出现姓名互换这类顺序错乱。这里先按上界扫描出视觉行，
    行内按横向排序，再按视觉行输出。
    """
    ordered = sorted(lines, key=lambda line: (line.top, line.x0))
    result: list[PdfTextLine] = []
    index = 0
    while index < len(ordered):
        end = index + 1
        while end < len(ordered) and _same_line_row(ordered[index], ordered[end]):
            end += 1
        result.extend(sorted(ordered[index:end], key=lambda line: line.x0))
        index = end
    return result


def _same_line_row(left: PdfTextLine, right: PdfTextLine) -> bool:
    """判断两条文本行是否属于同一视觉行。

    同一行文字若被拆成多段，各段的上界会因字形高低相差一两个点，而下界
    （基线）基本重合。上界或基线落在行高的一小段以内即视为同一行；
    正文行的行距远大于该容差，不会被并入。
    """
    scale = max(
        float(left.font_size or 0.0),
        left.bottom - left.top,
        float(right.font_size or 0.0),
        right.bottom - right.top,
    )
    if scale <= 0:
        return True
    return (
        abs(left.top - right.top) <= scale * 0.25
        or abs(left.bottom - right.bottom) <= scale * 0.6
    )


def _extract_text_lines(
    text_page: Any,
    page_height: float,
    *,
    styles: list[_TextObjectStyle] | None = None,
) -> list[PdfTextLine]:
    characters = _extract_text_characters(
        text_page,
        page_height,
        styles=styles,
    )
    rectangles = _line_rectangles(text_page, page_height)
    groups = _assign_characters_to_rectangles(characters, rectangles)
    if not any(groups):
        groups = [characters]
    else:
        assigned = sum(len(group) for group in groups)
        if assigned < len(characters):
            # 未被任何行矩形覆盖的字符仍按自身外框重建，避免静默丢字
            covered = {id(item) for group in groups for item in group}
            groups.append(
                [item for item in characters if id(item) not in covered]
            )
    groups = _merge_fragmented_groups(groups, rectangles)
    lines: list[PdfTextLine] = []
    for group in groups:
        if not group:
            continue
        lines.extend(_split_text_band(group))
    return _lines_in_reading_order(lines)


