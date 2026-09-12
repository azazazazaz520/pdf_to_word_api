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
    _same_text_band,
)
import math
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
        if raw_text in {"", "\r", "\n", "\t", "\ufffe"}:
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


def _split_text_band(characters: list[_TextCharacter]) -> list[PdfTextLine]:
    characters.sort(key=lambda item: item.x0)
    runs: list[list[_TextCharacter]] = []
    current: list[_TextCharacter] = []
    for character in characters:
        obj = character
        if current:
            previous = current[-1]
            gap = character.x0 - previous.x1
            split_gap = max(
                14.0,
                max(previous.font_size, character.font_size) * 1.3,
            )
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


def _build_text_spans(
    characters: list[_TextCharacter],
) -> tuple[PdfTextSpan, ...]:
    """把一行字符按字体/字号/颜色切成连续的 span。"""
    if not characters:
        return ()
    ordered = sorted(characters, key=lambda item: item.x0)
    groups: list[list[_TextCharacter]] = []
    current: list[_TextCharacter] = []
    previous_key: tuple[Any, ...] | None = None
    for character in ordered:
        key = character.style_key()
        if current and key != previous_key:
            groups.append(current)
            current = []
        current.append(character)
        previous_key = key
    if current:
        groups.append(current)

    spans: list[PdfTextSpan] = []
    for group in groups:
        span_text = _compact_text("".join(item.text for item in group))
        if not span_text:
            continue
        representative = max(group, key=lambda item: len(item.text))
        spans.append(
            PdfTextSpan(
                text=span_text,
                bbox=(
                    min(item.x0 for item in group),
                    min(item.top for item in group),
                    max(item.x1 for item in group),
                    max(item.bottom for item in group),
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
    bands: list[list[_TextCharacter]] = []
    for character in sorted(characters, key=lambda item: (item.top, item.x0)):
        matching_band = next(
            (
                band
                for band in reversed(bands)
                if _same_text_band(band[-1], character)
            ),
            None,
        )
        if matching_band is None:
            bands.append([character])
        else:
            matching_band.append(character)

    lines: list[PdfTextLine] = []
    for band in bands:
        lines.extend(_split_text_band(band))
    return sorted(lines, key=lambda line: (line.top, line.x0))


