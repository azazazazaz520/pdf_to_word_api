"""文本字符、字体样式与文本行的提取。

从页面图形对象读出字符盒、字体、字号、颜色与方向，按方向投影聚成行，
再按字体 run 切成 span。
"""

from __future__ import annotations

from dataclasses import replace
from typing import Any
from .models import (
    PdfTextGlyph,
    PdfTextLine,
    PdfTextSpan,
    _TextCharacter,
    _TextObjectStyle,
    _compact_text,
    _dominant_value,
)
import math
import ctypes
from bisect import bisect_left
from ctypes import c_float, c_int, c_uint
import pypdfium2.raw as pdfium_raw
from ..fonts.resolver import FontMatch, PdfFontDescriptor, resolve_font
from .models import _PAGEOBJ_TEXT


def _raw_pointer_value(value: Any) -> int:
    """返回 PDFium 对象句柄的稳定整数键。"""
    raw = getattr(value, "raw", value)
    try:
        return int(ctypes.cast(raw, ctypes.c_void_p).value or 0)
    except Exception:
        return 0

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
    resolve_fonts: bool = True,
) -> list[_TextObjectStyle]:
    """提取页面中所有文本对象的字体、颜色和 z-order。"""
    styles: list[_TextObjectStyle] = []
    cache = font_cache if font_cache is not None else {}
    try:
        objects = list(page.get_objects(textpage=text_page))
    except Exception:
        return styles
    for index, page_object in enumerate(objects):
        object_key = _raw_pointer_value(page_object)
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
            font = None
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
            if resolve_fonts:
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
                font_name = match.family or family_hint or base_name or ""
                bold = match.bold
                italic = match.italic
                substituted = match.substituted
                fallback_reason = match.fallback_reason
            else:
                font_name = family_hint or base_name or ""
                bold = _font_is_bold(font, weight or 0, base_name, family_hint)
                italic = _font_is_italic(font, base_name, family_hint)
                substituted = False
                fallback_reason = ""
            rotation = 0.0
            matrix_values_tuple: tuple[float, float, float, float, float, float] = ()
            matrix_scale = 1.0
            try:
                matrix = page_object.get_matrix()
                matrix_values = list(matrix.get())
                if len(matrix_values) >= 6:
                    matrix_values_tuple = tuple(
                        float(value) for value in matrix_values[:6]
                    )
                rotation = _matrix_rotation(matrix)
                matrix_scale = _matrix_scale(matrix_values)
            except Exception:
                rotation = 0.0
                matrix_scale = 1.0
            try:
                nominal_font_size = float(page_object.get_font_size())
                font_size = nominal_font_size * matrix_scale
            except Exception:
                nominal_font_size = max(bbox[3] - bbox[1], 1.0)
                font_size = nominal_font_size
                font_size_source = "character_geometry"
            else:
                font_size_source = "object_font_matrix"
            color = _object_fill_color(page_object) or (0, 0, 0)
            styles.append(
                _TextObjectStyle(
                    bbox=bbox,
                    font_name=font_name,
                    font_size=max(font_size, 0.1),
                    color=color,
                    bold=bold,
                    italic=italic,
                    z_order=index,
                    rotation=rotation,
                    pdf_font_name=base_name or family_hint,
                    substituted=substituted,
                    fallback_reason=fallback_reason,
                    object_key=object_key,
                    object_index=index,
                    nominal_font_size=max(nominal_font_size, 0.0),
                    font_size_source=font_size_source,
                    font_matrix=matrix_values_tuple,
                )
            )
        finally:
            close_object = getattr(page_object, "close", None)
            if close_object is not None:
                close_object()
    return styles


def _matrix_scale(matrix_values: Any) -> float:
    """从文本矩阵中提取文字高度缩放，修正 PDF 中被缩放绘制的字号。"""
    try:
        values = list(matrix_values)
    except TypeError:
        return 1.0
    if len(values) < 4:
        return 1.0
    _, _, c, d = (float(value) for value in values[:4])
    scale = math.hypot(c, d)
    if scale <= 1e-9:
        return 1.0
    return max(min(scale, 20.0), 0.01)


def _matrix_scales(
    matrix_values: Any,
) -> tuple[float, float]:
    """返回文字方向和文字高度方向的独立缩放。"""
    try:
        values = list(matrix_values)
        if len(values) < 4:
            return 1.0, 1.0
        a, b, c, d = (float(value) for value in values[:4])
    except (TypeError, ValueError):
        return 1.0, 1.0
    return max(math.hypot(a, b), 0.01), max(math.hypot(c, d), 0.01)


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
    return _normalize_rotation(angle)


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


def _raw_character_text(text_page: Any, index: int) -> str:
    """按 PDFium 内部字符索引读取一个 Unicode 字符。

    ``get_text_range`` 是面向文本缓冲区的接口，可能过滤或插入字符；
    ``get_charbox`` 使用的却是内部字符数组，所以两者不能用同一个枚举
    下标直接配对。
    """
    getter = getattr(pdfium_raw, "FPDFText_GetUnicode", None)
    if getter is not None:
        try:
            codepoint = int(getter(text_page.raw, index))
        except (ValueError, OverflowError, TypeError):
            codepoint = None
        except Exception:
            codepoint = None
        if codepoint is not None:
            if (
                codepoint < 0
                or codepoint in {0, 0xFFFE}
                or codepoint > 0x10FFFF
            ):
                return ""
            return chr(codepoint)
    sliced = True
    try:
        value = text_page.get_text_range(index, 1)
    except Exception:
        sliced = False
        try:
            value = text_page.get_text_range()
        except Exception:
            return ""
    if not sliced:
        value = value[index : index + 1]
    return value[0] if value else ""


def _normalize_rotation(value: float) -> float:
    """把角度规范到 ``[-180, 180]``，单位为度。"""
    if not math.isfinite(value):
        return 0.0
    angle = (float(value) + 180.0) % 360.0 - 180.0
    if abs(angle + 180.0) <= 0.05:
        angle = 180.0
    if abs(angle) <= 0.05:
        angle = 0.0
    return round(angle, 2)


def _raw_character_rotation(text_page: Any, index: int) -> float | None:
    """读取 PDFium 字符方向，并将其从弧度转换为度。"""
    getter = getattr(pdfium_raw, "FPDFText_GetCharAngle", None)
    if getter is None:
        return None
    try:
        value = float(getter(text_page.raw, index))
    except Exception:
        return None
    # PDFium 返回弧度；保留对少数兼容实现返回角度值的兼容性。
    degrees = math.degrees(value) if abs(value) <= (2.0 * math.pi + 0.1) else value
    return _normalize_rotation(degrees)


def _raw_character_font_size(text_page: Any, index: int) -> float:
    """读取字符实际绘制字号。"""
    getter = getattr(pdfium_raw, "FPDFText_GetFontSize", None)
    if getter is None:
        return 0.0
    try:
        value = float(getter(text_page.raw, index))
    except Exception:
        return 0.0
    return value if math.isfinite(value) and value > 0.0 else 0.0


def _direction_from_rotation(rotation: float) -> tuple[float, float]:
    """返回左上角页面坐标中的阅读方向单位向量。"""
    radians = math.radians(rotation)
    return (round(math.cos(radians), 6), round(math.sin(radians), 6))


def _normalize_aggregated_text(
    value: str,
    *,
    trim_edges: bool = False,
) -> str:
    """在字符已按几何顺序聚合后归一化文本。"""
    if not value:
        return ""
    # 少数内嵌字体把 en dash/em dash 读成替换字符组合；其语义可由固定
    # 的 PDFium 编码模式确认，先还原再做逐字符归一化。
    value = value.replace("\ufffdC", "–").replace("\ufffd\ufffd", "—")
    normalized = "".join(
        character
        if character in {" ", "\t"}
        else " "
        if character.isspace()
        else _compact_text(character)
        for character in value
    )
    return normalized.strip(" \t") if trim_edges else normalized


def _styles_by_object(
    styles: list[_TextObjectStyle] | None,
) -> dict[int, _TextObjectStyle]:
    if not styles:
        return {}
    return {
        style.object_key: style
        for style in styles
        if style.object_key
    }


def _style_for_character(
    text_page: Any,
    character: _TextCharacter,
    styles: list[_TextObjectStyle],
    styles_by_object: dict[int, _TextObjectStyle],
    style_index: dict[tuple[int, int], list[int]],
) -> _TextObjectStyle | None:
    """优先按 PDFium 文本对象句柄取样式，空间匹配只作兼容回退。"""
    try:
        text_object = text_page.get_textobj(character.char_index)
    except Exception:
        text_object = None
    if text_object is not None:
        style = styles_by_object.get(_raw_pointer_value(text_object))
        if style is not None:
            return style
    return _match_style(character, styles, style_index)


def _extract_text_characters(
    text_page: Any,
    page_height: float,
    *,
    styles: list[_TextObjectStyle] | None = None,
    page_number: int = 1,
) -> list[_TextCharacter]:
    characters: list[_TextCharacter] = []
    style_index = _style_index(styles) if styles else {}
    styles_by_object = _styles_by_object(styles)
    try:
        count = int(text_page.count_chars())
    except Exception:
        try:
            count = len(text_page.get_text_range())
        except Exception:
            count = 0
    for index in range(max(count, 0)):
        raw_text = _raw_character_text(text_page, index)
        if raw_text in {"", "\r", "\n", "\ufffe"} or _is_xml_control_character(
            raw_text
        ):
            continue
        if not raw_text:
            continue
        try:
            x0, y0, x1, y1 = (
                float(value) for value in text_page.get_charbox(index)
            )
        except Exception:
            continue
        height = max(abs(y1 - y0), 0.1)
        raw_rotation = _raw_character_rotation(text_page, index)
        raw_font_size = _raw_character_font_size(text_page, index)
        rotation = raw_rotation if raw_rotation is not None else 0.0
        direction = _direction_from_rotation(rotation)
        character = _TextCharacter(
            text=raw_text,
            x0=x0,
            top=page_height - max(y0, y1),
            x1=x1,
            bottom=page_height - min(y0, y1),
            font_size=raw_font_size or height,
            rotation=rotation,
            direction=direction,
            char_index=index,
            source_char_id=f"{page_number}:{index}",
            nominal_font_size=raw_font_size or height,
            font_size_source="character_metrics" if raw_font_size > 0.2 else "character_geometry",
        )
        if styles:
            style = _style_for_character(
                text_page,
                character,
                styles,
                styles_by_object,
                style_index,
            )
            if style is not None:
                character.font_name = style.font_name
                character.color = style.color
                character.bold = style.bold
                character.italic = style.italic
                if raw_rotation is None:
                    character.rotation = style.rotation
                    character.direction = _direction_from_rotation(
                        character.rotation
                    )
                character.z_order = style.z_order
                character.object_index = style.object_index
                character.pdf_font_name = style.pdf_font_name
                character.font_substituted = style.substituted
                character.font_fallback_reason = style.fallback_reason
                character.nominal_font_size = style.nominal_font_size or character.nominal_font_size
                character.font_size_source = style.font_size_source or character.font_size_source
                character.font_matrix = style.font_matrix
                if style.font_size > 0.0:
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


_ORIENTATION_TOLERANCE = 8.0


def _angle_distance(left: float, right: float) -> float:
    """返回两个有方向角之间的最小差值。"""
    return abs((float(left) - float(right) + 180.0) % 360.0 - 180.0)


def _mean_rotation(characters: list[_TextCharacter]) -> float:
    """用字符方向向量求一组字符的代表角度。"""
    if not characters:
        return 0.0
    x_total = 0.0
    y_total = 0.0
    for character in characters:
        dx, dy = character.direction
        x_total += dx
        y_total += dy
    if abs(x_total) <= 1e-6 and abs(y_total) <= 1e-6:
        return _normalize_rotation(float(characters[0].rotation))
    return _normalize_rotation(math.degrees(math.atan2(y_total, x_total)))


def _projection(
    character: _TextCharacter,
    axis: tuple[float, float],
) -> float:
    return character.center_x * axis[0] + character.center_y * axis[1]


def _line_normal_projection(
    character: _TextCharacter,
    normal: tuple[float, float],
    rotation: float,
) -> float:
    """返回更接近文字基线的行法向坐标。

    横排大字的字符外框可能向上延伸到相邻小字行。使用外框中心会让大字
    把两行错误合并；横排文字的下边界更接近基线，旋转文字继续使用中心
    投影以避免改变原有方向处理。
    """
    if _angle_distance(rotation, 0.0) <= 5.0:
        return character.bottom
    if _angle_distance(rotation, 180.0) <= 5.0:
        return -character.top
    return _projection(character, normal)


def _projected_extent(
    character: _TextCharacter,
    axis: tuple[float, float],
) -> tuple[float, float]:
    points = (
        (character.x0, character.top),
        (character.x0, character.bottom),
        (character.x1, character.top),
        (character.x1, character.bottom),
    )
    values = [point[0] * axis[0] + point[1] * axis[1] for point in points]
    return min(values), max(values)


def _orientation_groups(
    characters: list[_TextCharacter],
) -> list[tuple[float, list[_TextCharacter]]]:
    """按有方向角拆分字符，避免把竖排文字放进横排文字行。"""
    groups: list[tuple[float, list[_TextCharacter]]] = []
    for character in sorted(
        characters,
        key=lambda item: (item.char_index if item.char_index >= 0 else 0),
    ):
        best_index = None
        best_distance = None
        for index, (rotation, group) in enumerate(groups):
            distance = _angle_distance(character.rotation, rotation)
            if distance <= _ORIENTATION_TOLERANCE and (
                best_distance is None or distance < best_distance
            ):
                best_index = index
                best_distance = distance
        if best_index is None:
            groups.append((character.rotation, [character]))
            continue
        rotation, group = groups[best_index]
        group.append(character)
        groups[best_index] = (_mean_rotation(group), group)
    return groups


def _split_geometry_row(
    characters: list[_TextCharacter],
    rotation: float,
) -> list[list[_TextCharacter]]:
    """按阅读方向投影，把同一视觉行中的大间隔拆为独立文字段。"""
    direction = _direction_from_rotation(rotation)
    ordered = sorted(
        characters,
        key=lambda item: (
            _projection(item, direction),
            item.char_index if item.char_index >= 0 else 0,
        ),
    )
    sizes = [
        float(item.font_size)
        for item in ordered
        if float(item.font_size) > 0.2
    ]
    line_size = sizes[len(sizes) // 2] if sizes else 8.0
    split_gap = max(line_size * 1.3, 8.0)
    runs: list[list[_TextCharacter]] = []
    current: list[_TextCharacter] = []
    previous_end = None
    for character in ordered:
        start, end = _projected_extent(character, direction)
        if (
            current
            and previous_end is not None
            and start - previous_end > split_gap
            and not character.text.isspace()
        ):
            runs.append(current)
            current = []
        current.append(character)
        if _is_point_like(character) and character.text.isspace():
            continue
        previous_end = end if previous_end is None else max(previous_end, end)
    if current:
        runs.append(current)
    return runs


def _geometry_character_groups(
    characters: list[_TextCharacter],
) -> list[tuple[float, list[_TextCharacter]]]:
    """依据字符中心投影聚合成方向一致的视觉行。"""
    result: list[tuple[float, list[_TextCharacter]]] = []
    for orientation, oriented in _orientation_groups(characters):
        direction = _direction_from_rotation(orientation)
        normal = (-direction[1], direction[0])
        entries = sorted(
            (
                _line_normal_projection(character, normal, orientation),
                character.char_index if character.char_index >= 0 else 0,
                character,
            )
            for character in oriented
            if not (_is_point_like(character) and character.text.isspace())
        )
        rows: list[list[_TextCharacter]] = []
        row_centers: list[float] = []
        row_sizes: list[float] = []
        for normal_value, _, character in entries:
            size = max(
                float(character.font_size or 0.0),
                abs(character.x1 - character.x0),
                abs(character.bottom - character.top),
                1.0,
            )
            best_index = None
            best_distance = None
            for index, row_center in enumerate(row_centers):
                tolerance = max(
                    1.5,
                    0.55 * max(size, row_sizes[index]),
                )
                distance = abs(normal_value - row_center)
                if distance <= tolerance and (
                    best_distance is None or distance < best_distance
                ):
                    best_index = index
                    best_distance = distance
            if best_index is None:
                rows.append([character])
                row_centers.append(normal_value)
                row_sizes.append(size)
            else:
                rows[best_index].append(character)
                row_centers[best_index] = sum(
                    _line_normal_projection(item, normal, orientation)
                    for item in rows[best_index]
                ) / len(rows[best_index])
                sizes = sorted(
                    max(
                        float(item.font_size or 0.0),
                        abs(item.x1 - item.x0),
                        abs(item.bottom - item.top),
                        1.0,
                    )
                    for item in rows[best_index]
                )
                middle = len(sizes) // 2
                row_sizes[best_index] = (
                    sizes[middle]
                    if len(sizes) % 2
                    else (sizes[middle - 1] + sizes[middle]) / 2.0
                )
        row_for_character = {
            id(character): row_index
            for row_index, row in enumerate(rows)
            for character in row
        }
        for character in oriented:
            if not (_is_point_like(character) and character.text.isspace()):
                continue
            neighbours = sorted(
                (
                    abs(character.char_index - other.char_index),
                    other.char_index,
                    other,
                )
                for other in oriented
                if not (_is_point_like(other) and other.text.isspace())
                and other.char_index >= 0
                and character.char_index >= 0
                and other.char_index != character.char_index
            )
            if not neighbours:
                continue
            if not rows:
                continue
            previous = next(
                (item[2] for item in neighbours if item[1] < character.char_index),
                None,
            )
            following = next(
                (item[2] for item in neighbours if item[1] > character.char_index),
                None,
            )
            if previous is not None and following is not None:
                font_scale = max(
                    float(previous.font_size or 0.0),
                    float(following.font_size or 0.0),
                    1.0,
                )
                left_gap = max(character.x0 - previous.x1, 0.0)
                right_gap = max(following.x0 - character.x1, 0.0)
                if max(left_gap, right_gap) > max(font_scale * 2.5, 8.0):
                    continue
            candidate_rows = [
                row_for_character[id(item)]
                for item in (previous, following)
                if item is not None and id(item) in row_for_character
            ]
            target_row = None
            if candidate_rows and len(set(candidate_rows)) == 1:
                target_row = candidate_rows[0]
            elif candidate_rows:
                target_row = candidate_rows[0]
            else:
                target_row = min(
                    range(len(rows)),
                    key=lambda index: abs(
                        _line_normal_projection(character, normal, orientation)
                        - row_centers[index]
                    ),
                )
            if target_row is not None:
                rows[target_row].append(character)
        for row in rows:
            row_rotation = _mean_rotation(row)
            for run in _split_geometry_row(row, row_rotation):
                if run:
                    result.append((row_rotation, run))
    return result


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


def _ordered_characters(
    characters: list[_TextCharacter],
    rotation: float | None = None,
) -> list[_TextCharacter]:
    line_rotation = _mean_rotation(characters) if rotation is None else rotation
    object_indices = {
        item.object_index
        for item in characters
        if item.object_index >= 0
    }
    if len(object_indices) == 1 and all(item.char_index >= 0 for item in characters):
        # 一个 PDF 文本对象本身已经携带了可靠的字符流顺序，尤其适用于
        # 竖排或任意角度文字；多个对象组成的同行再使用几何投影排序。
        return sorted(characters, key=lambda item: item.char_index)
    direction = _direction_from_rotation(line_rotation)
    ordered = sorted(
        characters,
        key=lambda item: (
            _projection(item, direction),
            item.char_index if item.char_index >= 0 else 0,
        ),
    )
    if _angle_distance(line_rotation, 0.0) <= 5.0 or _angle_distance(
        line_rotation, 180.0
    ) <= 5.0:
        # 空格常只有一个点的字框，PDFium 给出的 x 坐标可能落在前一个字
        # 左侧几十分之一点。按几何坐标排序会把空格插入词内；空格自身的
        # 字符索引仍可靠，因此只对这类空格恢复字符流中的相邻关系。
        non_point_spaces = [
            item
            for item in ordered
            if not (_is_point_like(item) and item.text.isspace())
        ]
        point_spaces = sorted(
            (
                item
                for item in characters
                if _is_point_like(item)
                and item.text.isspace()
                and item.char_index >= 0
            ),
            key=lambda item: item.char_index,
        )
        if point_spaces:
            ordered = list(non_point_spaces)
            for space in point_spaces:
                previous = [
                    index
                    for index, item in enumerate(ordered)
                    if item.char_index >= 0
                    and item.char_index < space.char_index
                ]
                insert_at = previous[-1] + 1 if previous else 0
                ordered.insert(insert_at, space)
    return ordered


def _glyph_from_character(character: _TextCharacter) -> PdfTextGlyph:
    return PdfTextGlyph(
        text=character.text,
        bbox=(
            min(character.x0, character.x1),
            min(character.top, character.bottom),
            max(character.x0, character.x1),
            max(character.top, character.bottom),
        ),
        font_name=character.font_name,
        pdf_font_name=character.pdf_font_name,
        font_size=character.font_size,
        color=character.color,
        bold=character.bold,
        italic=character.italic,
        rotation=character.rotation,
        direction=character.direction,
        z_order=character.z_order,
        char_index=character.char_index,
        object_index=character.object_index,
        substituted=character.font_substituted,
        fallback_reason=character.font_fallback_reason,
        source_char_id=character.source_char_id,
        nominal_font_size=character.nominal_font_size,
        font_size_source=character.font_size_source,
        font_matrix=character.font_matrix,
    )


def _build_text_line(
    characters: list[_TextCharacter],
    *,
    rotation: float | None = None,
) -> PdfTextLine | None:
    if not characters:
        return None
    line_rotation = _mean_rotation(characters) if rotation is None else rotation
    ordered = _ordered_characters(characters, line_rotation)
    text = _normalize_aggregated_text(
        "".join(item.text for item in ordered),
        trim_edges=True,
    )
    if not text:
        return None
    font_sizes = sorted(
        item.font_size for item in ordered if item.font_size > 0
    )
    font_names = [item.font_name for item in ordered if item.font_name]
    colors = [item.color for item in ordered]
    spans = _trim_span_edges(
        list(_build_text_spans(ordered, rotation=line_rotation))
    )
    dominant_span = max(spans, key=lambda item: len(item.text), default=None)
    glyphs = tuple(_glyph_from_character(item) for item in ordered)
    return PdfTextLine(
        text=text,
        x0=min(item.x0 for item in ordered),
        top=min(item.top for item in ordered),
        x1=max(item.x1 for item in ordered),
        bottom=max(item.bottom for item in ordered),
        font_size=font_sizes[len(font_sizes) // 2] if font_sizes else 0.0,
        font_name=_dominant_value(font_names, ""),
        color=_dominant_value(colors, (0, 0, 0)),
        bold=sum(item.bold for item in ordered) * 2 >= len(ordered),
        italic=sum(item.italic for item in ordered) * 2 >= len(ordered),
        rotation=line_rotation,
        direction=_direction_from_rotation(line_rotation),
        z_order=min(item.z_order for item in ordered),
        spans=spans,
        pdf_font_name=dominant_span.pdf_font_name if dominant_span else "",
        font_substituted=any(item.font_substituted for item in ordered),
        font_fallback_reason=(
            dominant_span.fallback_reason if dominant_span else ""
        ),
        glyphs=glyphs,
        source_char_ids=tuple(
            glyph.source_char_id for glyph in glyphs if glyph.source_char_id
        ),
        nominal_font_size=(
            dominant_span.nominal_font_size if dominant_span else 0.0
        ),
        font_size_source=(
            dominant_span.font_size_source if dominant_span else ""
        ),
    )


def _split_text_band(
    characters: list[_TextCharacter],
    *,
    rotation: float | None = None,
) -> list[PdfTextLine]:
    """把字符组装成一条或多条方向一致的文本行。"""
    if rotation is not None:
        line = _build_text_line(characters, rotation=rotation)
        return [line] if line is not None else []
    result: list[PdfTextLine] = []
    for row_rotation, row in _geometry_character_groups(characters):
        line = _build_text_line(row, rotation=row_rotation)
        if line is not None:
            result.append(line)
    return result


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
    if character.pdf_font_name != first.pdf_font_name:
        return False
    if bool(character.bold) != bool(first.bold):
        return False
    if bool(character.italic) != bool(first.italic):
        return False
    if character.color != first.color:
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


def _move_boundary_whitespace(
    groups: list[list[_TextCharacter]],
) -> list[list[_TextCharacter]]:
    """把字体切换处的前导空白归到前一个可见文本组。"""
    result: list[list[_TextCharacter]] = []
    for source in groups:
        group = list(source)
        if result:
            prefix_length = 0
            while (
                prefix_length < len(group)
                and group[prefix_length].text.isspace()
            ):
                prefix_length += 1
            if prefix_length:
                result[-1].extend(group[:prefix_length])
                group = group[prefix_length:]
        if group:
            result.append(group)
    return result


def _trim_span_edges(
    spans: list[PdfTextSpan],
) -> tuple[PdfTextSpan, ...]:
    """移除整行两端的布局空白，保留 span 之间的词间空格。"""
    if not spans:
        return ()
    result = list(spans)
    if len(result) == 1:
        text = result[0].text.strip(" \t")
        return (replace(result[0], text=text),) if text else ()
    result[0] = replace(result[0], text=result[0].text.lstrip(" \t"))
    result[-1] = replace(result[-1], text=result[-1].text.rstrip(" \t"))
    return tuple(span for span in result if span.text)


def _build_text_spans(
    characters: list[_TextCharacter],
    *,
    rotation: float | None = None,
) -> tuple[PdfTextSpan, ...]:
    """把一行字符按字体与其他排版变化切成连续的 span。"""
    if not characters:
        return ()
    ordered = _ordered_characters(characters, rotation=rotation)
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
    groups = _move_boundary_whitespace(groups)

    spans: list[PdfTextSpan] = []
    for group in groups:
        span_text = _normalize_aggregated_text(
            "".join(item.text for item in group)
        )
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
                direction=representative.direction,
                z_order=min(item.z_order for item in group),
                substituted=representative.font_substituted,
                fallback_reason=representative.font_fallback_reason,
                glyphs=tuple(_glyph_from_character(item) for item in group),
                source_char_ids=tuple(
                    item.source_char_id
                    for item in group
                    if item.source_char_id
                ),
                nominal_font_size=representative.nominal_font_size,
                font_size_source=representative.font_size_source,
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
    if _angle_distance(left.rotation, right.rotation) > _ORIENTATION_TOLERANCE:
        return False
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
    page_number: int = 1,
) -> list[PdfTextLine]:
    characters = _extract_text_characters(
        text_page,
        page_height,
        styles=styles,
        page_number=page_number,
    )
    lines: list[PdfTextLine] = []
    for rotation, group in _geometry_character_groups(characters):
        lines.extend(_split_text_band(group, rotation=rotation))
    return _lines_in_reading_order(lines)


