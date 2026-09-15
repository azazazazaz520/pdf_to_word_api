"""保真导出：一页一个 section，所有内容按源坐标绝对定位。

文本走 w:framePr 逐 span 定位，图片、矢量、表格、公式按 bbox 锚定，
无法重建的区域贴源区域图片。转换服务可显式选择该兼容路径。
"""

from __future__ import annotations

import math
from collections import Counter
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from io import BytesIO
from itertools import count
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from PIL import Image

from ..ir.model import IRBlock, IRDocument, IRPage, IRTextLine, IRWarning
from ..fonts.embedding import FontPlan, attach_embedded_fonts, measure_text_width
from ..formula_omml import formula_text_to_omml
from ..ooxml_positioning import (
    add_absolute_picture,
    add_absolute_shape,
    add_absolute_text_box,
    add_absolute_text_paragraph,
    header_footer_paragraph,
    make_flow_paragraph_minimal,
    new_canvas_paragraph,
    detach_header_footer,
    position_table,
    prepare_footer,
    prepare_header,
    rgb_to_hex,
    set_exact_page,
)
from ..layout.models import PdfTable

from .content import _layout_line_role
from .document_setup import (
    StageCallback,
    _notify_stage,
    _remove_initial_empty_paragraph,
    _scaled_page_size,
    _set_document_styles,
    _set_section_page,
)
from .table_render import _add_pdf_table

from .document_setup import render_pdf_region

# 目录域、书签与公式 OMML 与流式导出共用，定义在导出主模块；
# 在函数体内导入以避开导出主模块 ← 本模块的导入环。


FIDELITY_MODES = frozenset({"fidelity", "fidelity_hybrid"})
DEFAULT_FIDELITY_TEXT_MIN_CONFIDENCE = 0.0
DEFAULT_FIDELITY_HYBRID_TEXT_MIN_CONFIDENCE = 0.85
DEFAULT_FIDELITY_FALLBACK_DPI = 200.0
DEFAULT_FIDELITY_FALLBACK_MAX_PIXELS = 2_000_000
FIDELITY_DEFAULT_FONT = "Microsoft YaHei"
FIDELITY_TEXT_DX_FACTOR = 0.042
FIDELITY_TEXT_DX_BASE = 0.05
FIDELITY_TEXT_DY_FACTOR = 0.074
FIDELITY_TEXT_DY_BASE = 0.02
"""Word 中 framePr 原点到字形框的偏移经验系数。

framePr 的 x/y 是段落框原点，而 IR 中的 bbox 来自 PDF 字形框；Word
渲染时存在与字号成正比的系统性偏移。导出时按字号补偿，使渲染后的
字形框与源 PDF 对齐（实测残差 < 0.3pt）。
"""
_FIDELITY_TEXT_KINDS = frozenset(
    {
        "title",
        "heading1",
        "heading2",
        "heading3",
        "paragraph",
        "body",
        "line",
        "ordered",
        "bullet",
        "plugin",
        "caption",
        "quote",
        "code",
    }
)


def _fidelity_text_origin(
    bbox: tuple[float, float, float, float],
    font_size: float,
    *,
    scale_x: float = 1.0,
    scale_y: float = 1.0,
    font_name: str = "",
    offsets: Mapping[str, float] | None = None,
) -> tuple[float, float]:
    """把字形框 bbox 换算为 framePr 原点，补偿 Word 的字形偏移。

    字体经过标定时使用实测系数 ``k``（``frame_y = glyph_top - k*size``），
    否则退回历史经验值。
    """
    size = max(float(font_size), 1.0)
    offset_x = FIDELITY_TEXT_DX_FACTOR * size + FIDELITY_TEXT_DX_BASE
    if offsets and font_name in offsets:
        offset_y = float(offsets[font_name]) * size
    else:
        offset_y = FIDELITY_TEXT_DY_FACTOR * size + FIDELITY_TEXT_DY_BASE
    return bbox[0] - offset_x * scale_x, bbox[1] - offset_y * scale_y


_FIDELITY_TEXT_KINDS = frozenset(
    {
        "title",
        "heading1",
        "heading2",
        "heading3",
        "paragraph",
        "body",
        "line",
        "ordered",
        "bullet",
        "plugin",
        "caption",
        "quote",
        "code",
    }
)


@dataclass
class _FidelityContext:
    """高保真导出的共享上下文。"""

    document: Any
    source_pdf: Path | None
    mode: str
    min_confidence: float
    fallback_dpi: float
    fallback_max_pixels: int
    page_width_points: float = 612.0
    scale_x: float = 1.0
    scale_y: float = 1.0
    object_ids: Any = None
    font_plan: FontPlan | None = None
    font_offsets: Mapping[str, float] | None = None
    font_programs: Mapping[str, Any] | None = None
    compensate_advance: bool = False

    def next_object_id(self) -> int:
        return next(self.object_ids)

    def scaled_bbox(
        self,
        bbox: tuple[float, float, float, float] | None,
    ) -> tuple[float, float, float, float] | None:
        if bbox is None:
            return None
        return (
            bbox[0] * self.scale_x,
            bbox[1] * self.scale_y,
            bbox[2] * self.scale_x,
            bbox[3] * self.scale_y,
        )


def _planned_font(
    context: _FidelityContext,
    raw_name: str,
    fallback_name: str,
    bold: bool,
    italic: bool,
) -> tuple[str, bool, bool]:
    """按字体计划返回实际使用的 Word 字体与粗斜体标记。"""
    if context.font_plan is None:
        return fallback_name, bool(bold), bool(italic)
    name = context.font_plan.word_name_for(raw_name, fallback_name)
    resolved_bold, resolved_italic = context.font_plan.style_for(
        raw_name, bold, italic
    )
    return name or fallback_name, resolved_bold, resolved_italic


def _compensate_advance(
    context: _FidelityContext,
    *,
    font_name: str,
    text: str,
    font_size: float,
    source_width: float,
) -> float:
    if not bool(getattr(context, "compensate_advance", False)):
        return 0.0
    """按内嵌字体度量计算字符间距补偿，抵消 w:sz 只能取半磅带来的累计漂移。"""
    programs = context.font_programs or {}
    program = programs.get(font_name)
    if program is None or len(text) < 3 or source_width <= 0:
        return 0.0
    expected = measure_text_width(program.data, text, font_size)
    if expected <= 0:
        return 0.0
    rounded_size = max(round(font_size * 2) / 2.0, 0.5)
    if abs(rounded_size - font_size) > 0.001:
        expected *= rounded_size / font_size
    spacing = (source_width - expected) / max(len(text) - 1, 1)
    if not -1.0 < spacing < 1.5:
        return 0.0
    return round(spacing, 4)


def _fidelity_frame_width(
    *,
    x_points: float,
    line_width: float,
    page_width: float,
    font_size: float,
) -> float:
    """计算 framePr 宽度。

    绝对定位的 frame 宽度只影响换行，不影响其他内容；这里给足冗余，
    避免 Word 使用的替代字体比源 PDF 略宽时发生意外换行。
    """
    remaining = max(float(page_width) - float(x_points), 0.0)
    return max(
        float(line_width) + 12.0,
        remaining + max(float(font_size), 6.0) * 4.0,
        12.0,
    )


def _fidelity_min_confidence(mode: str, override: float | None) -> float:
    if override is not None:
        return float(override)
    if mode == "fidelity_hybrid":
        return DEFAULT_FIDELITY_HYBRID_TEXT_MIN_CONFIDENCE
    return DEFAULT_FIDELITY_TEXT_MIN_CONFIDENCE


def _fidelity_fallback_reason(
    block: IRBlock,
    *,
    min_confidence: float,
) -> str:
    """判断非正文块是否需要区域图片兜底。

    正文即使置信度较低，也必须保留为可编辑文字；置信度只用于报告，
    不再触发文字区域图片。
    """
    if block.kind in _FIDELITY_TEXT_KINDS:
        return ""
    if block.fallback_reason:
        return block.fallback_reason
    if block.fallback_image is not None:
        return "explicit_fallback_image"
    if block.kind in {"html_table"} and not block.bbox:
        return "table_without_geometry"
    if block.kind == "table" and block.table is None:
        return "table_without_geometry"
    if block.kind == "vector" and getattr(block.vector, "complex", False):
        return "complex_vector"
    return ""


def _region_image_bytes(
    context: _FidelityContext,
    page: IRPage,
    block: IRBlock,
    *,
    bbox: tuple[float, float, float, float] | None = None,
    override: bytes | None = None,
) -> bytes | None:
    """渲染块所在区域的原始页面图像，作为图片兜底。"""
    if override:
        return override
    if block.fallback_image:
        return block.fallback_image
    if context.source_pdf is None:
        return None
    region = bbox or block.bbox
    if region is None:
        return None
    page_index = block.meta.get("page_index")
    if not isinstance(page_index, int):
        page_index = page.page_number - 1
    scale = max(1.0, float(context.fallback_dpi) / 72.0)
    try:
        return render_pdf_region(
            context.source_pdf,
            page_index,
            region,
            max_pixels=context.fallback_max_pixels,
            jpeg_quality=88,
            scale=scale,
        )
    except Exception:
        return None


def _record_placement(
    report: dict[str, Any],
    block: IRBlock,
    *,
    status: str,
    reason: str = "",
    bbox: tuple[float, float, float, float] | None = None,
) -> None:
    region = bbox or block.bbox
    report["placements"].append(
        {
            "kind": block.kind,
            "layer": block.layer,
            "status": status,
            "reason": reason,
            "bbox": [round(value, 2) for value in region] if region else None,
        }
    )
    if status == "native":
        report["native_block_count"] = report.get("native_block_count", 0) + 1
    elif status == "image_fallback":
        report["image_fallback_count"] = report.get("image_fallback_count", 0) + 1
    else:
        report["skipped_block_count"] = report.get("skipped_block_count", 0) + 1


def _place_region_fallback(
    context: _FidelityContext,
    canvas: Any,
    page: IRPage,
    block: IRBlock,
    *,
    report: dict[str, Any],
    reason: str,
    bbox: tuple[float, float, float, float] | None = None,
    behind_text: bool = False,
) -> bool:
    region = bbox or block.bbox
    if region is None:
        _record_placement(report, block, status="skipped", reason=reason)
        return False
    image_bytes = _region_image_bytes(
        context,
        page,
        block,
        bbox=region,
        override=block.fallback_image,
    )
    if not image_bytes:
        _record_placement(report, block, status="skipped", reason=f"{reason}:no_image")
        return False
    scaled = context.scaled_bbox(region) or region
    width = max(scaled[2] - scaled[0], 0.5)
    height = max(scaled[3] - scaled[1], 0.5)
    add_absolute_picture(
        canvas.container,
        image_bytes,
        x_points=scaled[0],
        y_points=scaled[1],
        width_points=width,
        height_points=height,
        z_order=block.z_order,
        object_id=context.next_object_id(),
        name=f"fallback_{page.page_number}_{len(report['placements'])}",
        behind_text=behind_text,
        paragraph=canvas.paragraph,
    )
    report["fallback_regions"].append(
        {
            "kind": block.kind,
            "layer": block.layer,
            "bbox": [round(value, 2) for value in region],
            "reason": reason,
        }
    )
    _record_placement(report, block, status="image_fallback", reason=reason)
    return True


def _place_text_block(
    context: _FidelityContext,
    page: IRPage,
    block: IRBlock,
    *,
    report: dict[str, Any],
) -> None:
    """按行（或按块）绝对定位写入文本。"""
    lines = list(block.lines)
    if not lines:
        if block.bbox is None:
            _record_placement(report, block, status="skipped", reason="missing_bbox")
            return
        lines = [
            IRTextLine(
                text=block.text,
                bbox=block.bbox,
                font_name=block.font_name,
                font_size=block.font_size,
                color=block.color,
                bold=block.bold,
                italic=block.italic,
                alignment=block.alignment or "left",
                line_spacing=block.line_spacing,
                rotation=block.rotation,
                z_order=block.z_order,
                confidence=block.confidence,
            )
        ]
    offset_max = 0.0
    placed_lines = 0
    for line in lines:
        spans = list(getattr(line, "spans", ()) or ())
        if spans:
            for span in spans:
                if _place_text_span(
                    context,
                    page,
                    block,
                    span,
                    report=report,
                    sequence=placed_lines,
                ):
                    placed_lines += 1
                    offset_max = max(
                        offset_max,
                        0.045 * float(span.font_size or block.font_size or 10.0),
                    )
            continue
        if not str(line.text or "").strip():
            continue
        bbox = context.scaled_bbox(line.bbox)
        if bbox is None:
            continue
        font_size = float(line.font_size or block.font_size or 10.0)
        font_name = (
            line.font_name
            or block.font_name
            or FIDELITY_DEFAULT_FONT
        )
        width = _fidelity_frame_width(
            x_points=bbox[0],
            line_width=max(bbox[2] - bbox[0], 4.0),
            page_width=context.page_width_points,
            font_size=font_size,
        )
        height = max(bbox[3] - bbox[1], font_size)
        color = line.color if line.color is not None else block.color
        bold = bool(line.bold)
        italic = bool(line.italic)
        font_name, bold, italic = _planned_font(
            context,
            str(getattr(line, "pdf_font_name", "") or ""),
            font_name,
            bold,
            italic,
        )
        rotation = float(line.rotation or block.rotation or 0.0)
        origin_x, origin_y = _fidelity_text_origin(
            bbox,
            font_size,
            scale_x=context.scale_x,
            scale_y=context.scale_y,
            font_name=font_name,
            offsets=context.font_offsets,
        )
        if abs(rotation) >= 1.0:
            add_absolute_text_box(
                context.document,
                x_points=origin_x,
                y_points=origin_y,
                width_points=width,
                height_points=max(height, font_size * 1.2),
                text=line.text,
                font_name=font_name,
                font_size=font_size,
                bold=bold,
                italic=italic,
                color=color,
                alignment=line.alignment or "left",
                rotation=rotation,
                z_order=line.z_order,
                object_id=context.next_object_id(),
                name=f"textbox_{page.page_number}_{placed_lines}",
            )
        else:
            add_absolute_text_paragraph(
                context.document,
                line.text,
                x_points=origin_x,
                y_points=origin_y,
                width_points=width,
                font_name=font_name,
                font_size=font_size,
                bold=bold,
                italic=italic,
                color=color,
                alignment="left",
                line_spacing_points=font_size,
                character_spacing_points=_compensate_advance(
                    context,
                    font_name=font_name,
                    text=line.text,
                    font_size=font_size,
                    source_width=bbox[2] - bbox[0],
                ),
            )
        offset_max = max(offset_max, 0.045 * font_size)
        placed_lines += 1
    if placed_lines == 0:
        _record_placement(report, block, status="skipped", reason="empty_text")
        return
    report["text_line_count"] = report.get("text_line_count", 0) + placed_lines
    report["estimated_bbox_offset_max"] = max(
        report.get("estimated_bbox_offset_max", 0.0), offset_max
    )
    report["font_size_error_max"] = report.get("font_size_error_max", 0.0)
    _record_placement(report, block, status="native", reason="text_frame")


def _place_text_span(
    context: _FidelityContext,
    page: IRPage,
    block: IRBlock,
    span: Any,
    *,
    report: dict[str, Any],
    sequence: int,
) -> bool:
    """按 span 的字体信息写入一段绝对定位文本。"""
    text = str(getattr(span, "text", "") or "")
    if not text.strip():
        return False
    bbox = context.scaled_bbox(tuple(float(v) for v in span.bbox))
    if bbox is None:
        return False
    span_size = float(getattr(span, "font_size", 0.0) or 0.0)
    font_size = span_size or float(block.font_size or 10.0)
    font_name = (
        getattr(span, "font_name", "")
        or block.font_name
        or FIDELITY_DEFAULT_FONT
    )
    color = getattr(span, "color", None)
    if color is None:
        color = block.color
    bold = bool(getattr(span, "bold", False))
    italic = bool(getattr(span, "italic", False))
    font_name, bold, italic = _planned_font(
        context,
        str(getattr(span, "pdf_font_name", "") or ""),
        font_name,
        bold,
        italic,
    )
    rotation = float(getattr(span, "rotation", 0.0) or 0.0)
    origin_x, origin_y = _fidelity_text_origin(
        bbox,
        font_size,
        scale_x=context.scale_x,
        scale_y=context.scale_y,
        font_name=font_name,
        offsets=context.font_offsets,
    )
    width = _fidelity_frame_width(
        x_points=origin_x,
        line_width=max(bbox[2] - bbox[0], 4.0),
        page_width=context.page_width_points,
        font_size=font_size,
    )
    if abs(rotation) >= 1.0:
        add_absolute_text_box(
            context.document,
            x_points=origin_x,
            y_points=origin_y,
            width_points=width,
            height_points=max(bbox[3] - bbox[1], font_size * 1.2),
            text=text,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            color=color,
            alignment="left",
            rotation=rotation,
            z_order=int(getattr(span, "z_order", 0) or 0),
            object_id=context.next_object_id(),
            name=f"textbox_{page.page_number}_{sequence}",
        )
        return True
    add_absolute_text_paragraph(
        context.document,
        text,
        x_points=origin_x,
        y_points=origin_y,
        width_points=width,
        font_name=font_name,
        font_size=font_size,
        bold=bold,
        italic=italic,
        color=color,
        alignment="left",
        line_spacing_points=font_size,
        character_spacing_points=_compensate_advance(
            context,
            font_name=font_name,
            text=text,
            font_size=font_size,
            source_width=bbox[2] - bbox[0],
        ),
    )
    return True


def _place_vector_block(
    context: _FidelityContext,
    canvas: Any,
    page: IRPage,
    block: IRBlock,
    *,
    report: dict[str, Any],
    behind_text: bool,
) -> None:
    vector = block.vector
    scaled = context.scaled_bbox(block.bbox)
    if vector is None or scaled is None:
        _place_region_fallback(
            context,
            canvas,
            page,
            block,
            report=report,
            reason="vector_without_geometry",
            behind_text=behind_text,
        )
        return
    kind = str(getattr(vector, "kind", "path") or "path")
    stroke = getattr(vector, "stroke_color", None)
    fill = getattr(vector, "fill_color", None)
    stroke_width = max(float(getattr(vector, "stroke_width", 0.0) or 0.0), 0.4)
    segments = list(getattr(vector, "segments", ()) or ())
    if getattr(vector, "stroked", True) and stroke is not None:
        # pdfium 的 get_bounds 会把描边宽度算进 bbox，这里换算回路径边界。
        inset = stroke_width / 2.0
        left = min(scaled[0] + inset, scaled[2])
        top = min(scaled[1] + inset, scaled[3])
        right = max(scaled[2] - inset, left)
        bottom = max(scaled[3] - inset, top)
        scaled = (left, top, right, bottom)
    if kind == "path":
        if len(segments) > 24:
            _place_region_fallback(
                context,
                canvas,
                page,
                block,
                report=report,
                reason="complex_vector",
                behind_text=behind_text,
            )
            return
        if segments and _draw_vector_segments(
            context,
            canvas,
            segments=segments,
            stroke_hex=rgb_to_hex(stroke) if stroke is not None else None,
            fill_hex=rgb_to_hex(fill) if fill is not None else None,
            stroke_width=stroke_width,
            z_order=block.z_order,
            behind_text=behind_text,
            name=f"vector_{page.page_number}",
        ):
            _record_placement(report, block, status="native", reason="vector_segments")
            return
    geometry = "rect" if kind in {"rect", "path"} else "line"
    width = max(scaled[2] - scaled[0], 0.4)
    height = max(scaled[3] - scaled[1], 0.4)
    if geometry == "line" and kind == "line":
        height = max(stroke_width, 0.4)
        fill = None
    elif geometry == "line":
        height = max(height, 0.4)
        fill = None
    add_absolute_shape(
        context.document,
        geometry=geometry,
        x_points=scaled[0],
        y_points=scaled[1],
        width_points=width,
        height_points=height,
        fill_color=fill,
        line_color=stroke,
        line_width_points=stroke_width,
        z_order=block.z_order,
        behind_text=behind_text,
        object_id=context.next_object_id(),
        name=f"shape_{page.page_number}",
        paragraph=canvas.paragraph,
    )
    _record_placement(report, block, status="native", reason=f"vector_{geometry}")


def _draw_vector_segments(
    context: _FidelityContext,
    canvas: Any,
    *,
    segments: list[tuple[str, float, float]],
    stroke_hex: str | None,
    fill_hex: str | None,
    stroke_width: float,
    z_order: int,
    behind_text: bool,
    name: str,
) -> bool:
    """把路径分段拆成若干绝对定位的线条形状。"""
    drawn = 0
    previous: tuple[float, float] | None = None
    for label, x, y in segments:
        if label == "M":
            previous = (x, y)
            continue
        if previous is None:
            previous = (x, y)
            continue
        x0, y0 = previous
        previous = (x, y)
        left = min(x0, x)
        top = min(y0, y)
        width = abs(x - x0)
        height = abs(y - y0)
        if width < 0.2 and height < 0.2:
            continue
        scaled_left = left * context.scale_x
        scaled_top = top * context.scale_y
        scaled_width = max(width * context.scale_x, 0.4)
        scaled_height = max(height * context.scale_y, 0.4)
        flip_v = y > y0
        flip_h = x < x0
        add_absolute_shape(
            context.document,
            geometry="line",
            x_points=scaled_left,
            y_points=scaled_top,
            width_points=scaled_width,
            height_points=scaled_height,
            line_color=stroke_hex or fill_hex or "000000",
            line_width_points=stroke_width,
            z_order=z_order,
            behind_text=behind_text,
            object_id=context.next_object_id(),
            name=name,
            paragraph=canvas.paragraph,
            flip_h=flip_h,
            flip_v=flip_v,
        )
        drawn += 1
    return drawn > 0


def _table_text_padding(
    table_data: Any,
) -> tuple[float, float] | None:
    """根据源单元格文本与单元格边界推算左右/上内边距。"""
    cells = list(getattr(table_data, "cells", ()) or ())
    if not cells:
        return None
    left_paddings: list[float] = []
    top_paddings: list[float] = []
    for cell in cells:
        text_bbox = getattr(cell, "text_bbox", None)
        if not text_bbox or text_bbox == (0.0, 0.0, 0.0, 0.0):
            continue
        left_paddings.append(text_bbox[0] - cell.bbox[0])
        top_paddings.append(text_bbox[1] - cell.bbox[1])
    if not left_paddings:
        return None
    left = max(min(left_paddings), 0.0)
    top = max(min(top_paddings), 0.0)
    first_font = max(float(getattr(cells[0], "font_size", 0.0) or 0.0), 1.0)
    # Word 单元格单倍行距会在文本上方留出约 0.53em 的行距空间，这里扣除。
    top = max(top - 0.53 * first_font, 0.0)
    return left, top


def _place_table_block(
    context: _FidelityContext,
    canvas: Any,
    page: IRPage,
    block: IRBlock,
    *,
    report: dict[str, Any],
    behind_text: bool,
) -> None:
    table_data = block.table
    scaled = context.scaled_bbox(block.bbox)
    if table_data is None or scaled is None:
        _place_region_fallback(
            context,
            canvas,
            page,
            block,
            report=report,
            reason="table_without_geometry",
            behind_text=behind_text,
        )
        return
    padding = _table_text_padding(table_data)
    try:
        table = _add_pdf_table(
            context.document,
            table_data,
            bordered=bool(getattr(table_data, "has_borders", True)),
            exact_widths=True,
            cell_vertical_alignment="top",
            tight_cell_margins=True,
            cell_padding_points=padding,
        )
        if table is None:
            raise ValueError("empty table")
        position_table(
            table,
            x_points=scaled[0] + (padding[0] if padding else 0.0),
            y_points=scaled[1],
            z_order=block.z_order,
        )
        make_flow_paragraph_minimal(context.document.add_paragraph())
    except Exception as error:
        _place_region_fallback(
            context,
            canvas,
            page,
            block,
            report=report,
            reason=f"table_rebuild_failed:{type(error).__name__}",
            behind_text=behind_text,
        )
        return
    _record_placement(report, block, status="native", reason="absolute_table")


def _place_image_block(
    context: _FidelityContext,
    canvas: Any,
    page: IRPage,
    block: IRBlock,
    *,
    report: dict[str, Any],
    behind_text: bool,
    status: str = "native",
    reason: str = "absolute_picture",
) -> None:
    if not block.image_bytes:
        _record_placement(report, block, status="skipped", reason="image_bytes_missing")
        return
    scaled = context.scaled_bbox(block.bbox)
    if scaled is None:
        if block.image_width and block.image_height:
            scaled = (0.0, 0.0, block.image_width, block.image_height)
        else:
            _record_placement(report, block, status="skipped", reason="missing_bbox")
            return
    add_absolute_picture(
        canvas.container,
        block.image_bytes,
        x_points=scaled[0],
        y_points=scaled[1],
        width_points=max(scaled[2] - scaled[0], 0.5),
        height_points=max(scaled[3] - scaled[1], 0.5),
        z_order=block.z_order,
        object_id=context.next_object_id(),
        name=f"image_{page.page_number}_{len(report['placements'])}",
        behind_text=behind_text
        or block.layer in {"background", "header", "footer"},
        paragraph=canvas.paragraph,
    )
    _record_placement(report, block, status=status, reason=reason)


def _place_formula_block(
    context: _FidelityContext,
    canvas: Any,
    page: IRPage,
    block: IRBlock,
    *,
    report: dict[str, Any],
    behind_text: bool,
) -> None:
    scaled = context.scaled_bbox(block.bbox)
    if scaled is not None and block.text.strip():
        formula_size = float(block.font_size or 10.0)
        origin_x, origin_y = _fidelity_text_origin(
            scaled,
            formula_size,
            scale_x=context.scale_x,
            scale_y=context.scale_y,
            font_name=block.font_name,
            offsets=context.font_offsets,
        )
        paragraph = add_absolute_text_paragraph(
            context.document,
            "",
            x_points=origin_x,
            y_points=origin_y,
            width_points=_fidelity_frame_width(
                x_points=scaled[0],
                line_width=max(scaled[2] - scaled[0], 6.0),
                page_width=context.page_width_points,
                font_size=formula_size,
            ),
            font_size=formula_size,
        )
        from ..page_render import _add_formula_omml_element

        if _add_formula_omml_element(paragraph, block.text):
            _record_placement(report, block, status="native", reason="omml")
            return
        try:
            paragraph._p.getparent().remove(paragraph._p)
        except Exception:
            pass
    _place_region_fallback(
        context,
        canvas,
        page,
        block,
        report=report,
        reason="formula_omml_failed",
        behind_text=behind_text,
    )


def _page_background_blocks(page: IRPage) -> list[IRBlock]:
    """找出整页填充矩形（页面背景）。"""
    result: list[IRBlock] = []
    for block in page.blocks:
        if block.kind != "vector" or block.bbox is None:
            continue
        x0, top, x1, bottom = block.bbox
        if (x1 - x0) < page.width * 0.9 or (bottom - top) < page.height * 0.9:
            continue
        vector = block.vector
        if vector is None:
            continue
        if getattr(vector, "fill_color", None) is None:
            continue
        if getattr(vector, "stroke_color", None) is not None:
            continue
        result.append(block)
    return result


def _is_white_fill(block: IRBlock) -> bool:
    vector = block.vector
    color = getattr(vector, "fill_color", None) if vector is not None else None
    if color is None:
        return False
    try:
        return all(int(channel) >= 250 for channel in color[:3])
    except Exception:
        return False


def _place_fidelity_block(
    context: _FidelityContext,
    canvas: Any,
    page: IRPage,
    block: IRBlock,
    *,
    report: dict[str, Any],
    behind_text: bool,
) -> None:
    reason = _fidelity_fallback_reason(
        block, min_confidence=context.min_confidence
    )
    if reason:
        _place_region_fallback(
            context,
            canvas,
            page,
            block,
            report=report,
            reason=reason,
            behind_text=behind_text,
        )
        return
    if block.kind == "page_image":
        _place_image_block(
            context, canvas, page, block, report=report, behind_text=True
        )
        return
    if block.kind == "image":
        _place_image_block(
            context, canvas, page, block, report=report, behind_text=behind_text
        )
        return
    if block.kind == "vector":
        _place_vector_block(
            context,
            canvas,
            page,
            block,
            report=report,
            behind_text=behind_text,
        )
        return
    if block.kind == "table":
        _place_table_block(
            context,
            canvas,
            page,
            block,
            report=report,
            behind_text=behind_text,
        )
        return
    if block.kind == "html_table":
        _place_region_fallback(
            context,
            canvas,
            page,
            block,
            report=report,
            reason="html_table_without_geometry",
            behind_text=behind_text,
        )
        return
    if block.kind == "formula":
        _place_formula_block(
            context,
            canvas,
            page,
            block,
            report=report,
            behind_text=behind_text,
        )
        return
    if block.kind in _FIDELITY_TEXT_KINDS:
        _place_text_block(context, page, block, report=report)
        return
    _record_placement(report, block, status="skipped", reason="unsupported_kind")


def _compute_rebuild_confidence(report: dict[str, Any]) -> float:
    weights: list[tuple[float, float]] = []
    for placement in report.get("placements", []):
        bbox = placement.get("bbox") or [0.0, 0.0, 1.0, 1.0]
        width = max(float(bbox[2]) - float(bbox[0]), 1.0)
        height = max(float(bbox[3]) - float(bbox[1]), 1.0)
        weight = math.sqrt(width * height)
        status = placement.get("status")
        if status == "native":
            score = 1.0
        elif status == "image_fallback":
            score = 0.8
        else:
            score = 0.0
        weights.append((weight, score))
    if not weights:
        return 1.0
    total_weight = sum(item[0] for item in weights)
    if total_weight <= 0:
        return 1.0
    return round(
        sum(weight * score for weight, score in weights) / total_weight,
        4,
    )


def _write_header_footer_blocks(
    context: _FidelityContext,
    section: Any,
    page: IRPage,
) -> bool:
    """把页眉页脚写入 Word header/footer part；失败返回 False。"""
    header_blocks = sorted(page.header_blocks, key=lambda item: item.z_order)
    footer_blocks = sorted(page.footer_blocks, key=lambda item: item.z_order)
    if not header_blocks and not footer_blocks:
        return False
    prepared: dict[str, list[tuple[Any, bytes | None]]] = {}
    for name, blocks in (("header", header_blocks), ("footer", footer_blocks)):
        items: list[tuple[Any, bytes | None]] = []
        for block in blocks:
            image_bytes: bytes | None = None
            if not _can_place_header_footer_item(block):
                image_bytes = _region_image_bytes(
                    context, page, block, override=block.fallback_image
                )
                if image_bytes is None:
                    return False
            elif block.kind == "vector" and getattr(
                block.vector, "complex", True
            ):
                image_bytes = _region_image_bytes(context, page, block)
                if image_bytes is None:
                    return False
            items.append((block, image_bytes))
        prepared[name] = items
    touched: list[Any] = []
    try:
        for name in ("header", "footer"):
            items = prepared.get(name) or []
            if not items:
                continue
            container = (
                prepare_header(section)
                if name == "header"
                else prepare_footer(section)
            )
            touched.append(container)
            first_paragraph = header_footer_paragraph(container)
            for index, (block, image_bytes) in enumerate(items):
                paragraph = (
                    first_paragraph if index == 0 else container.add_paragraph()
                )
                make_flow_paragraph_minimal(paragraph)
                _place_header_footer_item(
                    context,
                    container,
                    paragraph,
                    page,
                    block,
                    image_bytes=image_bytes,
                )
    except Exception:
        for container in touched:
            try:
                for paragraph in list(container.paragraphs):
                    element = paragraph._p
                    parent = element.getparent()
                    if parent is not None:
                        parent.remove(element)
            except Exception:
                continue
        return False
    page.header_footer_native = True
    return True


def _can_place_header_footer_item(block: IRBlock) -> bool:
    """页眉页脚块是否可以在不渲染图片的情况下重建。"""
    if block.kind in _FIDELITY_TEXT_KINDS:
        return bool(str(block.text or "").strip() or block.lines)
    if block.fallback_reason or block.fallback_image is not None:
        return False
    if block.kind == "image":
        return bool(block.image_bytes)
    if block.kind == "vector":
        return block.vector is not None and block.bbox is not None
    return False


def _place_header_footer_item(
    context: _FidelityContext,
    container: Any,
    paragraph: Any,
    page: IRPage,
    block: IRBlock,
    *,
    image_bytes: bytes | None,
) -> None:
    scaled = context.scaled_bbox(block.bbox)
    if image_bytes is not None:
        if scaled is None:
            raise ValueError("header/footer image without bbox")
        add_absolute_picture(
            container,
            image_bytes,
            x_points=scaled[0],
            y_points=scaled[1],
            width_points=max(scaled[2] - scaled[0], 0.5),
            height_points=max(scaled[3] - scaled[1], 0.5),
            z_order=block.z_order,
            object_id=context.next_object_id(),
            name=f"hf_{page.page_number}_{block.layer}",
            behind_text=False,
            paragraph=paragraph,
        )
        return
    if block.kind == "image":
        if not block.image_bytes or scaled is None:
            raise ValueError("header/footer image missing data")
        add_absolute_picture(
            container,
            block.image_bytes,
            x_points=scaled[0],
            y_points=scaled[1],
            width_points=max(scaled[2] - scaled[0], 0.5),
            height_points=max(scaled[3] - scaled[1], 0.5),
            z_order=block.z_order,
            object_id=context.next_object_id(),
            name=f"hf_{page.page_number}_image",
            behind_text=False,
            paragraph=paragraph,
        )
        return
    if block.kind == "vector":
        vector = block.vector
        if vector is None or scaled is None:
            raise ValueError("header/footer vector without geometry")
        stroke = getattr(vector, "stroke_color", None)
        fill = getattr(vector, "fill_color", None)
        stroke_width = max(float(getattr(vector, "stroke_width", 0.0) or 0.0), 0.4)
        if getattr(vector, "stroked", True) and stroke is not None:
            inset = stroke_width / 2.0
            left = min(scaled[0] + inset, scaled[2])
            top = min(scaled[1] + inset, scaled[3])
            right = max(scaled[2] - inset, left)
            bottom = max(scaled[3] - inset, top)
            scaled = (left, top, right, bottom)
        geometry = "rect" if getattr(vector, "kind", "path") == "rect" else "line"
        width = max(scaled[2] - scaled[0], 0.4)
        height = max(scaled[3] - scaled[1], 0.4)
        segments = list(getattr(vector, "segments", ()) or ())
        if segments and len(segments) <= 24:
            canvas = _CanvasContext(document=container, paragraph=paragraph)
            if _draw_vector_segments(
                context,
                canvas,
                segments=segments,
                stroke_hex=rgb_to_hex(stroke) if stroke is not None else None,
                fill_hex=rgb_to_hex(fill) if fill is not None else None,
                stroke_width=stroke_width,
                z_order=block.z_order,
                behind_text=False,
                name=f"hf_{page.page_number}_vector",
            ):
                return
        if geometry == "line":
            height = max(stroke_width, 0.4)
        add_absolute_shape(
            container,
            geometry=geometry,
            x_points=scaled[0],
            y_points=scaled[1],
            width_points=width,
            height_points=height,
            fill_color=fill,
            line_color=stroke,
            line_width_points=stroke_width,
            z_order=block.z_order,
            object_id=context.next_object_id(),
            name=f"hf_{page.page_number}_vector",
            paragraph=paragraph,
        )
        return
    lines = list(block.lines)
    if not lines and block.text and scaled is not None:
        lines = [
            IRTextLine(
                text=block.text,
                bbox=block.bbox,
                font_name=block.font_name,
                font_size=block.font_size,
                color=block.color,
                bold=block.bold,
                italic=block.italic,
                alignment="left",
                z_order=block.z_order,
            )
        ]
    if not lines:
        raise ValueError("header/footer block without content")
    used_paragraph = paragraph
    for position, line in enumerate(lines):
        line_spans = list(getattr(line, "spans", ()) or ())
        if line_spans:
            for span_index, span in enumerate(line_spans):
                bbox = context.scaled_bbox(span.bbox)
                if bbox is None:
                    continue
                span_size = float(span.font_size or line.font_size or 9.0)
                span_font, span_bold, span_italic = _planned_font(
                    context,
                    str(getattr(span, "pdf_font_name", "") or ""),
                    span.font_name or line.font_name or FIDELITY_DEFAULT_FONT,
                    bool(span.bold),
                    bool(span.italic),
                )
                origin_x, origin_y = _fidelity_text_origin(
                    bbox,
                    span_size,
                    scale_x=context.scale_x,
                    scale_y=context.scale_y,
                    font_name=span_font,
                    offsets=context.font_offsets,
                )
                if position or span_index:
                    used_paragraph = container.add_paragraph()
                    make_flow_paragraph_minimal(used_paragraph)
                add_absolute_text_paragraph(
                    container,
                    span.text,
                    x_points=origin_x,
                    y_points=origin_y,
                    width_points=_fidelity_frame_width(
                        x_points=origin_x,
                        line_width=max(bbox[2] - bbox[0], 4.0),
                        page_width=context.page_width_points,
                        font_size=span_size,
                    ),
                    font_name=span_font,
                    font_size=span_size,
                    bold=span_bold,
                    italic=span_italic,
                    color=span.color if span.color is not None else line.color,
                    alignment="left",
                    line_spacing_points=span_size,
                    paragraph=used_paragraph,
                )
            continue
        bbox = context.scaled_bbox(line.bbox)
        if bbox is None:
            continue
        font_size = float(line.font_size or block.font_size or 9.0)
        origin_x, origin_y = _fidelity_text_origin(
            bbox,
            font_size,
            scale_x=context.scale_x,
            scale_y=context.scale_y,
            font_name=line_font,
            offsets=context.font_offsets,
        )
        if position:
            used_paragraph = container.add_paragraph()
            make_flow_paragraph_minimal(used_paragraph)
        line_font, line_bold, line_italic = _planned_font(
            context,
            str(getattr(line, "pdf_font_name", "") or ""),
            line.font_name or block.font_name or FIDELITY_DEFAULT_FONT,
            bool(line.bold),
            bool(line.italic),
        )
        add_absolute_text_paragraph(
            container,
            line.text,
            x_points=origin_x,
            y_points=origin_y,
            width_points=_fidelity_frame_width(
                x_points=origin_x,
                line_width=max(bbox[2] - bbox[0], 4.0),
                page_width=context.page_width_points,
                font_size=font_size,
            ),
            font_name=line_font,
            font_size=font_size,
            bold=line_bold,
            italic=line_italic,
            color=line.color if line.color is not None else block.color,
            alignment="left",
            line_spacing_points=font_size,
            paragraph=used_paragraph,
        )


def export_fidelity_docx(
    ir: IRDocument,
    output_path: Path,
    *,
    title: str | None = None,
    source_pdf: Path | None = None,
    include_toc: bool = False,
    include_bookmarks: bool = True,
    mode: str = "fidelity",
    min_text_confidence: float | None = None,
    fallback_dpi: float = DEFAULT_FIDELITY_FALLBACK_DPI,
    fallback_max_pixels: int = DEFAULT_FIDELITY_FALLBACK_MAX_PIXELS,
    font_plan: FontPlan | None = None,
    font_offsets: Mapping[str, float] | None = None,
    font_programs: Mapping[str, Any] | None = None,
    compensate_advance: bool = False,
    stage_callback: StageCallback | None = None,
) -> dict[str, Any]:
    """高保真导出：一页一个 section，所有内容按页面坐标绝对定位。

    规则：

    * 每个 PDF 页对应一个 Word section，纸张尺寸/页边距等于源 PDF；
    * 文本用 ``w:framePr`` 逐行绝对定位，不压缩字号、不参与流式重排；
    * 图片、线条、矢量对象按 bbox 锚定；
    * 表格使用固定布局 + ``w:tblpPr`` 浮动定位；
    * 页眉页脚写入 Word header/footer，失败时在正文按原坐标兜底；
    * 无法重建的非正文区域渲染原区域图片兜底，并写入报告；正文始终保留为文字。
    """
    from itertools import count

    if mode not in FIDELITY_MODES:
        raise ValueError(f"不支持的高保真模式：{mode}")

    document = Document()
    _set_document_styles(document)
    document.core_properties.title = title or ir.title
    _remove_initial_empty_paragraph(document)
    if include_toc:
        document.add_heading("目录", level=1)
        from ..page_render import _add_toc_field

        _add_toc_field(document)
        document.add_page_break()

    outline_by_page: dict[int, list[dict[str, Any]]] = {}
    if include_bookmarks:
        for entry in ir.metadata.get("outline", []):
            page_number = entry.get("page")
            if isinstance(page_number, int):
                outline_by_page.setdefault(page_number, []).append(entry)

    min_confidence = _fidelity_min_confidence(mode, min_text_confidence)
    object_ids = count(1)
    bookmark_id = 0
    written_pages = 0
    for page in ir.pages:
        if written_pages == 0:
            section = document.sections[0]
        else:
            section = document.add_section(WD_SECTION.NEW_PAGE)
        written_pages += 1
        section_info = set_exact_page(
            section,
            width_points=page.width,
            height_points=page.height,
            margin_points=0.0,
            header_distance_points=0.0,
            footer_distance_points=0.0,
        )
        scale_x = section_info["width_points"] / max(page.width, 1.0)
        scale_y = section_info["height_points"] / max(page.height, 1.0)
        report: dict[str, Any] = {
            "page": page.page_number,
            "route": page.route,
            "placements": [],
            "fallback_regions": [],
            "native_block_count": 0,
            "image_fallback_count": 0,
            "skipped_block_count": 0,
            "text_line_count": 0,
            "scaled_page": bool(section_info["scaled"]),
        }
        if section_info["scaled"]:
            report["page_scale"] = round(scale_x, 4)

        first_paragraph = new_canvas_paragraph(document)
        canvas = _CanvasContext(document=document, paragraph=first_paragraph)
        context = _FidelityContext(
            document=document,
            source_pdf=source_pdf,
            mode=mode,
            min_confidence=min_confidence,
            fallback_dpi=fallback_dpi,
            fallback_max_pixels=fallback_max_pixels,
            page_width_points=section_info["width_points"],
            scale_x=scale_x,
            scale_y=scale_y,
            object_ids=object_ids,
            font_plan=font_plan,
            font_offsets=font_offsets,
            font_programs=font_programs,
            compensate_advance=bool(compensate_advance),
        )

        background_blocks = _page_background_blocks(page)
        white_backgrounds = {
            id(block) for block in background_blocks if _is_white_fill(block)
        }
        colored_background = any(
            id(block) not in white_backgrounds for block in background_blocks
        )
        header_footer_native = False
        # 有彩色页面背景时，页眉页脚必须留在正文层，否则会被背景盖住
        if (page.header_blocks or page.footer_blocks) and not colored_background:
            header_footer_native = _write_header_footer_blocks(context, section, page)
        if not header_footer_native:
            # 整页图片/空白页自带页眉页脚（或本就为空），必须断开继承避免串页
            detach_header_footer(section)

        page_image_block = next(
            (
                block
                for block in page.blocks
                if block.kind == "page_image" and block.image_bytes
            ),
            None,
        )
        forced_page_image = bool(page.fidelity.get("force_page_image"))
        page_image_fallback = page_image_block is not None and (
            forced_page_image
            or page.route in {"ocr", "page_image"}
            or not page.body_blocks
        )

        if page_image_fallback:
            if forced_page_image:
                page_image_fallback_reason = str(
                    page.fidelity.get("force_page_image_reason")
                    or "auto_fallback_text_coverage_below_threshold"
                )
            elif page.route == "ocr":
                page_image_fallback_reason = "ocr_page_preserved_as_image"
            else:
                page_image_fallback_reason = "page_preserved_as_image"
            page_image_block.fallback_reason = page_image_fallback_reason
            _place_image_block(
                context,
                canvas,
                page,
                page_image_block,
                report=report,
                behind_text=True,
                status="image_fallback",
                reason=page_image_fallback_reason,
            )
            report["page_image_fallback"] = True
            report["fallback_regions"].append(
                {
                    "kind": "page_image",
                    "layer": "background",
                    "bbox": [0.0, 0.0, round(page.width, 2), round(page.height, 2)],
                    "reason": page_image_fallback_reason,
                }
            )
            for block in page.blocks:
                if block is page_image_block:
                    continue
                _record_placement(
                    report,
                    block,
                    status="image_fallback",
                    reason="covered_by_page_image",
                )
        else:
            blocks = (
                list(page.body_blocks)
                if header_footer_native
                else list(page.blocks)
            )
            min_text_z = min(
                (
                    block.z_order
                    for block in blocks
                    if block.kind in _FIDELITY_TEXT_KINDS
                ),
                default=None,
            )
            blocks.sort(
                key=lambda item: (
                    item.z_order,
                    item.reading_order if item.reading_order >= 0 else 10**9,
                    item.kind,
                )
            )
            for block in blocks:
                if id(block) in white_backgrounds:
                    # Word 页面本身就是白色，重复画白底会盖住页眉/页脚
                    _record_placement(
                        report,
                        block,
                        status="skipped",
                        reason="page_background_white",
                    )
                    continue
                behind_text = False
                if block.kind in {"image", "vector"}:
                    behind_text = block.layer == "background" or (
                        block.layer == "body"
                        and min_text_z is not None
                        and block.z_order < min_text_z
                    )
                _place_fidelity_block(
                    context,
                    canvas,
                    page,
                    block,
                    report=report,
                    behind_text=behind_text,
                )

        report["rebuild_confidence"] = _compute_rebuild_confidence(report)
        page.reconstruction_confidence = report["rebuild_confidence"]
        page.fidelity = {**page.fidelity, **report}
        page.header_footer_native = header_footer_native
        for region in report["fallback_regions"]:
            page.add_warning(
                IRWarning(
                    code="fidelity_region_fallback",
                    message=(
                        f"第 {page.page_number} 页的区域无法重建，已使用原始图片兜底："
                        f"{region.get('reason')}"
                    ),
                    page=page.page_number,
                    severity="info",
                )
            )
        if include_bookmarks:
            page_entries = outline_by_page.get(page.page_number, [])
            if page_entries and first_paragraph is not None:
                from ..page_render import _add_bookmark, _bookmark_name

                for entry in page_entries:
                    bookmark_id += 1
                    _add_bookmark(
                        first_paragraph,
                        name=_bookmark_name(
                            str(entry.get("title") or ""), bookmark_id
                        ),
                        bookmark_id=bookmark_id,
                    )
        _notify_stage(
            stage_callback,
            "ir_page_completed",
            page=page.page_number,
            route=page.route,
            rebuild_confidence=report["rebuild_confidence"],
            fallback_region_count=len(report["fallback_regions"]),
        )

    make_flow_paragraph_minimal(document.add_paragraph())
    embedded_font_ids: dict[str, str] = {}
    if font_plan is not None and font_plan.programs:
        embedded_font_ids = attach_embedded_fonts(document, font_plan.programs)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    if embedded_font_ids:
        ir.metadata["embedded_fonts"] = embedded_font_ids
    fidelity_report = ir.fidelity_report()
    ir.metadata["fidelity_report"] = fidelity_report
    _notify_stage(
        stage_callback,
        "ir_export_completed",
        page_count=len(ir.pages),
        table_count=ir.table_count,
        media_count=ir.media_count,
        fallback_region_count=fidelity_report.get("fallback_region_count", 0),
        mean_rebuild_confidence=fidelity_report.get("mean_rebuild_confidence"),
    )
    return fidelity_report


@dataclass
class _CanvasContext:
    """页面画布：所有浮动对象共享同一个零高度段落。"""

    document: Any
    paragraph: Any

    @property
    def container(self) -> Any:
        return self.document
