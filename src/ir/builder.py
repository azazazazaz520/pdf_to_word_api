from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Any, Iterable

from .model import (
    IRBlock,
    IRDocument,
    IRPage,
    IRTextGlyph,
    IRTextLine,
    IRTextSpan,
    IRWarning,
)
from ..layout.models import (
    PdfContentBlock,
    PdfDocumentLayout,
    PdfPageLayout,
    PdfTextLine,
)
from ..export.content import (
    _group_text_lines,
    _layout_lines_to_text,
    _line_is_in_pdf_table,
    _ordered_blocks,
    _typical_body_left,
)
from ..export.text_utils import _field, _plain_text


_BLOCK_TEXT_KINDS = {
    "doc_title": "title",
    "paragraph_title": "heading1",
    "title": "title",
    "figure_title": "caption",
}


def _role_to_kind(role: str) -> str:
    if role == "heading1":
        return "heading1"
    if role == "heading2":
        return "heading2"
    if role == "heading3":
        return "heading3"
    if role == "title":
        return "title"
    if role in {"ordered", "bullet", "plugin", "formula", "code"}:
        return role
    return "paragraph"


def _block_level(kind: str) -> int:
    if kind == "title":
        return 0
    if kind == "heading1":
        return 1
    if kind == "heading2":
        return 2
    if kind == "heading3":
        return 3
    return 0


def _text_block(
    *,
    kind: str,
    page_number: int,
    text: str,
    source: str,
    confidence: float | None = None,
    bbox: tuple[float, float, float, float] | None = None,
) -> IRBlock:
    return IRBlock(
        kind=kind,
        page=page_number,
        text=text,
        role=kind,
        level=_block_level(kind),
        bbox=bbox,
        source=source,
        confidence=confidence,
    )


def _apply_content_metadata(
    block: IRBlock,
    content_block: PdfContentBlock | None,
    *,
    suffix: str = "",
) -> None:
    """把版面阶段确定的块信息传给 IR，供导出和质量报告使用。"""
    if content_block is None:
        return
    block.block_id = (
        f"{content_block.block_id}:{suffix}"
        if suffix
        else content_block.block_id
    )
    block.block_type = content_block.block_type
    block.reading_order = content_block.reading_order
    block.parent_id = content_block.parent_id
    block.region_id = content_block.region_id
    block.needs_review = content_block.needs_review
    block.meta["source_block_id"] = content_block.block_id
    block.meta["reading_order"] = content_block.reading_order
    block.meta["region_id"] = content_block.region_id
    if content_block.needs_review:
        block.meta["reading_order_needs_review"] = True


def _dominant(values: Iterable[Any], default: Any) -> Any:
    collected = [value for value in values if value is not None]
    if not collected:
        return default
    return Counter(collected).most_common(1)[0][0]


def _median(values: Iterable[float], default: float = 0.0) -> float:
    collected = sorted(float(value) for value in values if value)
    if not collected:
        return default
    return collected[len(collected) // 2]


def _union_bbox(
    lines: Iterable[PdfTextLine],
) -> tuple[float, float, float, float] | None:
    collected = [
        (line.x0, line.top, line.x1, line.bottom) for line in lines
    ]
    if not collected:
        return None
    return (
        min(item[0] for item in collected),
        min(item[1] for item in collected),
        max(item[2] for item in collected),
        max(item[3] for item in collected),
    )


def _ir_text_line(
    line: PdfTextLine,
    *,
    confidence: float | None = None,
) -> IRTextLine:
    return IRTextLine(
        text=line.text,
        bbox=(line.x0, line.top, line.x1, line.bottom),
        font_name=line.font_name,
        font_size=line.font_size,
        color=line.color,
        bold=line.bold,
        italic=line.italic,
        alignment=line.alignment,
        line_spacing=line.line_spacing,
        first_line_indent=line.first_line_indent,
        rotation=line.rotation,
        z_order=line.z_order,
        confidence=confidence,
        spans=tuple(_ir_text_span(span) for span in line.spans),
        pdf_font_name=getattr(line, "pdf_font_name", ""),
        substituted=bool(getattr(line, "font_substituted", False)),
        fallback_reason=getattr(line, "font_fallback_reason", ""),
        direction=tuple(getattr(line, "direction", (1.0, 0.0))),
        glyphs=tuple(
            _ir_text_glyph(glyph)
            for glyph in getattr(line, "glyphs", ())
        ),
    )


def _ir_text_glyph(glyph: Any) -> IRTextGlyph:
    bbox = getattr(glyph, "bbox", (0.0, 0.0, 0.0, 0.0))
    return IRTextGlyph(
        text=glyph.text,
        bbox=tuple(float(value) for value in bbox),
        font_name=getattr(glyph, "font_name", ""),
        pdf_font_name=getattr(glyph, "pdf_font_name", ""),
        font_size=float(getattr(glyph, "font_size", 0.0) or 0.0),
        color=getattr(glyph, "color", None),
        bold=bool(getattr(glyph, "bold", False)),
        italic=bool(getattr(glyph, "italic", False)),
        rotation=float(getattr(glyph, "rotation", 0.0) or 0.0),
        direction=tuple(getattr(glyph, "direction", (1.0, 0.0))),
        z_order=int(getattr(glyph, "z_order", 0) or 0),
        char_index=int(getattr(glyph, "char_index", -1)),
        object_index=int(getattr(glyph, "object_index", -1)),
        substituted=bool(getattr(glyph, "substituted", False)),
        fallback_reason=getattr(glyph, "fallback_reason", ""),
    )


def _ir_text_span(span: Any) -> IRTextSpan:
    bbox = getattr(span, "bbox", (0.0, 0.0, 0.0, 0.0))
    return IRTextSpan(
        text=span.text,
        bbox=tuple(float(value) for value in bbox),
        font_name=getattr(span, "font_name", ""),
        pdf_font_name=getattr(span, "pdf_font_name", ""),
        font_size=float(getattr(span, "font_size", 0.0) or 0.0),
        color=getattr(span, "color", None),
        bold=bool(getattr(span, "bold", False)),
        italic=bool(getattr(span, "italic", False)),
        rotation=float(getattr(span, "rotation", 0.0) or 0.0),
        z_order=int(getattr(span, "z_order", 0) or 0),
        substituted=bool(getattr(span, "substituted", False)),
        fallback_reason=getattr(span, "fallback_reason", ""),
        direction=tuple(getattr(span, "direction", (1.0, 0.0))),
        glyphs=tuple(
            _ir_text_glyph(glyph)
            for glyph in getattr(span, "glyphs", ())
        ),
    )


def _apply_fidelity_style(
    block: IRBlock,
    lines: list[PdfTextLine],
    *,
    page_width: float,
    confidence: float | None,
) -> None:
    """把行级几何和样式聚合到块级字段，供绝对定位导出使用。"""
    if not lines:
        return
    block.bbox = _union_bbox(lines)
    block.lines = tuple(
        _ir_text_line(line, confidence=confidence) for line in lines
    )
    block.font_name = _dominant((line.font_name for line in lines), "")
    block.font_size = _median((line.font_size for line in lines), 0.0)
    block.color = _dominant((line.color for line in lines), (0, 0, 0))
    block.bold = sum(line.bold for line in lines) * 2 >= len(lines)
    block.italic = sum(line.italic for line in lines) * 2 >= len(lines)
    block.alignment = _dominant((line.alignment for line in lines), "left")
    block.line_spacing = _median(
        (line.line_spacing for line in lines if line.line_spacing > 0),
        0.0,
    )
    block.rotation = _dominant(
        (line.rotation for line in lines if abs(line.rotation) > 0.05),
        0.0,
    )
    block.z_order = min(line.z_order for line in lines)
    left_edges = [line.x0 for line in lines]
    block.first_line_indent = max(lines[0].x0 - min(left_edges), 0.0)
    block.meta.setdefault("page_width", page_width)
    block.meta.setdefault("page_index", block.page - 1)
    block.meta["line_count"] = len(lines)


def build_text_page_ir(
    page: PdfPageLayout,
    *,
    page_number: int,
    is_document_start: bool,
    confidence: float | None = None,
    fidelity: bool = False,
    keep_header_footer: bool = False,
) -> IRPage:
    """把带文本层和几何布局的页面转成 IR。

    ``fidelity=True`` 时补充行级 bbox、字体、颜色、对齐、行距、z-order
    以及矢量对象；``keep_header_footer=True`` 时把页眉页脚保留为
    独立块的 header/footer 层，供 Word 原生页眉页脚导出使用。
    """
    header_footer_lines = [line for line in page.lines if line.is_header_footer]
    header_footer_count = len(header_footer_lines)
    warnings: list[IRWarning] = []
    if header_footer_count and not keep_header_footer:
        warnings.append(
            IRWarning(
                code="header_footer_filtered",
                message=f"已识别并过滤 {header_footer_count} 行页眉页脚或页码。",
                page=page_number,
                severity="info",
            )
        )
    for message in page.reading_order_warnings:
        warnings.append(
            IRWarning(
                code="reading_order_uncertain",
                message=message,
                page=page_number,
                severity="warning",
                meta={
                    "confidence": page.reading_order_confidence,
                },
            )
        )
    result = IRPage(
        page_number=page_number,
        width=page.width,
        height=page.height,
        route="text",
        confidence=confidence,
        warnings=warnings,
        editable=True,
        reading_order_confidence=page.reading_order_confidence,
        reading_order_warnings=list(page.reading_order_warnings),
    )

    body_lines = [line for line in page.lines if not line.is_header_footer]
    body_candidates = [
        line
        for line in body_lines
        if not any(_line_is_in_pdf_table(line, table) for table in page.tables)
    ]
    body_left = _typical_body_left(body_candidates)
    document_start_pending = is_document_start

    def append_text_lines(
        text_lines: list[PdfTextLine],
        content_block: PdfContentBlock | None = None,
    ) -> None:
        nonlocal document_start_pending
        if not text_lines:
            return
        content, line_roles = _layout_lines_to_text(
            text_lines,
            body_left=body_left,
            is_document_start=document_start_pending,
        )
        spans: list[tuple[int, ...]] = []
        grouped = _group_text_lines(
            content,
            is_document_start=document_start_pending,
            line_roles=line_roles,
            spans_out=spans,
        )
        for position, (role, block_text) in enumerate(grouped):
            kind = _role_to_kind(role)
            if content_block is not None:
                if content_block.block_type == "FORMULA" and kind == "paragraph":
                    kind = "formula"
                elif content_block.block_type == "CAPTION" and kind == "paragraph":
                    kind = "caption"
                elif content_block.block_type == "TITLE" and kind == "paragraph":
                    kind = "title" if document_start_pending else "heading1"
            block = _text_block(
                kind=kind,
                page_number=page_number,
                text=block_text,
                source="text_layout",
                confidence=confidence,
                bbox=(content_block.bbox if content_block else None),
            )
            _apply_content_metadata(
                block,
                content_block,
                suffix=(str(position) if len(grouped) > 1 else ""),
            )
            span = spans[position] if position < len(spans) else ()
            source_lines = [
                text_lines[index]
                for index in span
                if 0 <= index < len(text_lines)
            ]
            if fidelity:
                _apply_fidelity_style(
                    block,
                    source_lines,
                    page_width=page.width,
                    confidence=confidence,
                )
            if kind == "formula" and len(source_lines) == 1:
                line = source_lines[0]
                block.bbox = (line.x0, line.top, line.x1, line.bottom)
                block.meta["page_index"] = page_number - 1
                block.meta["page_width"] = page.width
                block.latex = _formula_latex(block.text)
                block.z_order = line.z_order
            result.blocks.append(block)
        document_start_pending = False

    def append_table(
        table: Any,
        content_block: PdfContentBlock | None = None,
    ) -> None:
        block = IRBlock(
            kind="table",
            page=page_number,
            table=table,
            bbox=table.bbox,
            source="vector_table",
            confidence=confidence,
            z_order=getattr(table, "z_order", 0),
        )
        _apply_content_metadata(block, content_block)
        result.blocks.append(block)

    def append_image(
        image: Any,
        content_block: PdfContentBlock | None = None,
    ) -> None:
        block = IRBlock(
            kind="image",
            page=page_number,
            bbox=image.bbox,
            source="embedded_image",
            image_bytes=image.data,
            image_width=image.width,
            image_height=image.height,
            image_alt=image.name,
            confidence=confidence,
            z_order=getattr(image, "z_order", 0),
            layer=getattr(image, "layer", "body"),
            meta={
                "page_width": page.width,
                "is_logo": getattr(image, "is_logo", False),
            },
        )
        _apply_content_metadata(block, content_block)
        result.blocks.append(block)

    def append_vector(
        vector: Any,
        content_block: PdfContentBlock | None = None,
    ) -> None:
        block = IRBlock(
            kind="vector",
            page=page_number,
            bbox=vector.bbox,
            source="pdf_vector",
            confidence=confidence,
            z_order=getattr(vector, "z_order", 0),
            layer=getattr(vector, "layer", "body"),
            vector=vector,
        )
        _apply_content_metadata(block, content_block)
        result.blocks.append(block)

    content_blocks = tuple(getattr(page, "content_blocks", ()) or ())
    if content_blocks:
        for content_block in content_blocks:
            if content_block.layer in {"header", "footer"}:
                continue
            if content_block.text_block is not None:
                block_lines = [
                    line
                    for line in content_block.text_block.lines
                    if not line.is_header_footer
                    and not any(
                        _line_is_in_pdf_table(line, table)
                        for table in page.tables
                    )
                ]
                append_text_lines(block_lines, content_block)
            elif content_block.table is not None:
                append_table(content_block.table, content_block)
            elif content_block.image is not None:
                append_image(content_block.image, content_block)
            elif content_block.vector is not None and fidelity:
                append_vector(content_block.vector, content_block)
    else:
        events: list[tuple[float, int, str, Any]] = []
        for table in page.tables:
            events.append((table.bbox[1], 0, "table", table))
        for image in page.images:
            events.append((image.bbox[1], 1, "image", image))
        if fidelity:
            for vector in page.vectors:
                events.append((vector.bbox[1], 1, "vector", vector))
        geometry_blocks = tuple(getattr(page, "text_blocks", ()) or ())
        if geometry_blocks:
            for text_block in geometry_blocks:
                block_lines = [
                    line
                    for line in text_block.lines
                    if not line.is_header_footer
                    and not any(
                        _line_is_in_pdf_table(line, table)
                        for table in page.tables
                    )
                ]
                if block_lines:
                    events.append(
                        (block_lines[0].top, 2, "text_block", block_lines)
                    )
        else:
            for line in body_lines:
                if not any(_line_is_in_pdf_table(line, table) for table in page.tables):
                    events.append((line.top, 2, "line", line))
        events.sort(key=lambda item: (item[0], item[1]))
        text_lines: list[PdfTextLine] = []
        for _, _, kind, item in events:
            if kind == "line":
                text_lines.append(item)
                continue
            if kind == "text_block":
                append_text_lines(item)
                continue
            append_text_lines(text_lines)
            text_lines = []
            if kind == "table":
                append_table(item)
            elif kind == "image":
                append_image(item)
            elif kind == "vector":
                append_vector(item)
        append_text_lines(text_lines)

    if fidelity and keep_header_footer:
        _append_header_footer_blocks(
            result,
            page,
            page_number=page_number,
            confidence=confidence,
        )
    if not result.blocks:
        result.route = "blank"
        result.editable = True
        result.add_warning(
            IRWarning(
                code="source_blank_page",
                message="源 PDF 该页只有页眉页脚，过滤后为空白页。",
                page=page_number,
                severity="info",
            )
        )
    return result


def _append_header_footer_blocks(
    page_ir: IRPage,
    page: PdfPageLayout,
    *,
    page_number: int,
    confidence: float | None = None,
) -> None:
    """把页眉页脚文本、logo 和分隔线整理成独立层块。"""
    line_sources = {
        id(line): content_block
        for content_block in page.content_blocks
        if content_block.text_block is not None
        for line in content_block.text_block.lines
    }
    for line in page.lines:
        if not line.is_header_footer:
            continue
        block = _text_block(
            kind="paragraph",
            page_number=page_number,
            text=line.text,
            source="header_footer_text",
            confidence=confidence,
        )
        _apply_content_metadata(block, line_sources.get(id(line)))
        _apply_fidelity_style(
            block,
            [line],
            page_width=page.width,
            confidence=confidence,
        )
        block.layer = line.layer
        page_ir.blocks.append(block)
    for image in page.images:
        if image.layer == "body":
            continue
        block = IRBlock(
            kind="image",
            page=page_number,
            bbox=image.bbox,
            source="header_footer_image",
            image_bytes=image.data,
            image_width=image.width,
            image_height=image.height,
            image_alt=image.name,
            z_order=image.z_order,
            layer=image.layer,
            confidence=confidence,
            meta={"page_width": page.width, "is_logo": image.is_logo},
        )
        content_block = next(
            (
                item
                for item in page.content_blocks
                if item.image is image
            ),
            None,
        )
        _apply_content_metadata(block, content_block)
        page_ir.blocks.append(block)
    for vector in page.vectors:
        if vector.layer == "body":
            continue
        block = IRBlock(
            kind="vector",
            page=page_number,
            bbox=vector.bbox,
            source="header_footer_vector",
            z_order=vector.z_order,
            layer=vector.layer,
            vector=vector,
            confidence=confidence,
        )
        content_block = next(
            (
                item
                for item in page.content_blocks
                if item.vector is vector
            ),
            None,
        )
        _apply_content_metadata(block, content_block)
        page_ir.blocks.append(block)


def _formula_latex(text: str) -> str:
    try:
        from ..formula_omml import formula_text_to_latex

        return formula_text_to_latex(text)
    except Exception:
        return ""


def build_page_image_ir(
    *,
    page_number: int,
    width: float,
    height: float,
    image_bytes: bytes,
    reason: str,
    confidence: float | None = 1.0,
) -> IRPage:
    """把需要视觉保真的页面转成整页图片 IR。"""
    page = IRPage(
        page_number=page_number,
        width=width,
        height=height,
        route="page_image",
        blocks=[
            IRBlock(
                kind="page_image",
                page=page_number,
                source="page_render",
                image_bytes=image_bytes,
                image_width=width,
                image_height=height,
                image_alt=f"第 {page_number} 页原始页面图像",
                confidence=confidence,
                bbox=(0.0, 0.0, width, height),
                layer="background",
            )
        ],
        confidence=confidence,
        warnings=[
            IRWarning(
                code="page_preserved_as_image",
                message=f"该页以整页图像保留，视觉内容不丢失，但文字不可编辑。原因：{reason}",
                page=page_number,
                severity="warning",
            )
        ],
        editable=False,
    )
    page.fidelity = {
        "page_image_fallback": True,
        "fallback_regions": [
            {
                "kind": "page_image",
                "layer": "background",
                "bbox": [0.0, 0.0, round(width, 3), round(height, 3)],
                "reason": reason or "page_preserved_as_image",
            }
        ],
    }
    return page


def build_ocr_page_ir(
    result: Any,
    *,
    page_number: int,
    width: float,
    height: float,
    quality: dict[str, Any] | None = None,
    include_page_image: bytes | None = None,
    ocr_scale: float | None = None,
) -> IRPage:
    """把一页 OCR 结果转成 IR。

    ``ocr_scale`` 为 OCR 渲染图相对 PDF 点坐标的缩放系数；提供后会把
    OCR 像素 bbox 换算为 point，供高保真模式定位或图片兜底使用。
    """
    quality = quality or {}
    warnings: list[IRWarning] = []
    if quality.get("ocr_needs_review"):
        warnings.append(
            IRWarning(
                code="ocr_needs_review",
                message="该页 OCR 置信度较低或存在疑似乱码，建议人工复核。",
                page=page_number,
                severity="warning",
            )
        )
    confidence = quality.get("ocr_mean_confidence")
    page = IRPage(
        page_number=page_number,
        width=width,
        height=height,
        route="ocr",
        confidence=confidence if isinstance(confidence, (int, float)) else None,
        warnings=warnings,
        editable=True,
    )
    if include_page_image:
        page.blocks.append(
            IRBlock(
                kind="page_image",
                page=page_number,
                source="hybrid_page_render",
                image_bytes=include_page_image,
                image_width=width,
                image_height=height,
                image_alt=f"第 {page_number} 页原始页面图像",
                bbox=(0.0, 0.0, width, height),
                layer="background",
            )
        )
    for block in _ordered_blocks(result):
        label = str(_field(block, "block_label") or "text")
        content = str(_field(block, "block_content") or "").strip()
        if not content:
            continue
        bbox = _ocr_block_bbox(block, ocr_scale)
        if label == "table" or "<table" in content.lower():
            page.blocks.append(
                IRBlock(
                    kind="html_table",
                    page=page_number,
                    text=content,
                    source="ocr_table",
                    confidence=page.confidence,
                    bbox=bbox,
                )
            )
            continue
        if label in _BLOCK_TEXT_KINDS:
            kind = _BLOCK_TEXT_KINDS[label]
            created = _text_block(
                kind=kind,
                page_number=page_number,
                text=_plain_text(content),
                source="ocr",
                confidence=page.confidence,
                bbox=bbox,
            )
            created.meta["page_width"] = width
            page.blocks.append(created)
            continue
        for role, block_text in _group_text_lines(content):
            created = _text_block(
                kind=_role_to_kind(role),
                page_number=page_number,
                text=block_text,
                source="ocr",
                confidence=page.confidence,
                bbox=bbox,
            )
            created.meta["page_width"] = width
            page.blocks.append(created)
    return page


def _ocr_block_bbox(
    block: Any,
    scale: float | None,
) -> tuple[float, float, float, float] | None:
    if not scale or scale <= 0:
        return None
    raw_bbox = _field(block, "block_bbox")
    if not raw_bbox or len(raw_bbox) < 4:
        return None
    try:
        x0, y0, x1, y1 = (float(value) / float(scale) for value in raw_bbox[:4])
    except Exception:
        return None
    return (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))


def build_document_ir(
    *,
    source_pdf: Path,
    page_routes: list[str],
    page_sizes: list[tuple[float, float]],
    layout: PdfDocumentLayout | None = None,
    ocr_results: list[Any] | None = None,
    ocr_quality: list[dict[str, Any]] | None = None,
    ocr_scales: list[float | None] | None = None,
    page_images: dict[int, bytes] | None = None,
    text_confidence: list[float] | None = None,
    route_mode: str = "auto",
    export_mode: str = "hybrid",
    engine: str | None = None,
    model_version: str | None = None,
    fidelity: bool = False,
    keep_header_footer: bool = False,
) -> IRDocument:
    """根据逐页路由结果组装统一 Document IR。"""
    page_images = page_images or {}
    ocr_results = ocr_results or []
    ocr_quality = ocr_quality or []
    ocr_scales = ocr_scales or []
    text_confidence = text_confidence or []
    page_count = len(page_routes)
    document = IRDocument(
        title=source_pdf.stem,
        metadata={
            "source_pdf": str(source_pdf),
            "route_mode": route_mode,
            "export_mode": export_mode,
            "engine": engine,
            "model_version": model_version,
            "fidelity_mode": fidelity,
        },
    )
    for page_index, route in enumerate(page_routes):
        page_number = page_index + 1
        width, height = (
            page_sizes[page_index]
            if page_index < len(page_sizes)
            else (612.0, 792.0)
        )
        if route == "text" and layout is not None and page_index < len(layout.pages):
            confidence = (
                text_confidence[page_index]
                if page_index < len(text_confidence)
                else None
            )
            document.pages.append(
                build_text_page_ir(
                    layout.pages[page_index],
                    page_number=page_number,
                    is_document_start=page_index == 0,
                    confidence=confidence,
                    fidelity=fidelity,
                    keep_header_footer=keep_header_footer,
                )
            )
        elif route == "ocr":
            result = (
                ocr_results[page_index]
                if page_index < len(ocr_results)
                else {"parsing_res_list": []}
            )
            quality = (
                ocr_quality[page_index]
                if page_index < len(ocr_quality)
                else None
            )
            ocr_scale = (
                ocr_scales[page_index]
                if page_index < len(ocr_scales)
                else None
            )
            include_image = page_images.get(page_index)
            document.pages.append(
                build_ocr_page_ir(
                    result,
                    page_number=page_number,
                    width=width,
                    height=height,
                    quality=quality,
                    include_page_image=include_image,
                    ocr_scale=ocr_scale,
                )
            )
        elif route == "page_image":
            image_bytes = page_images.get(page_index)
            if image_bytes is None:
                document.pages.append(
                    IRPage(
                        page_number=page_number,
                        width=width,
                        height=height,
                        route="page_image",
                        editable=False,
                        warnings=[
                            IRWarning(
                                code="page_image_missing",
                                message="页面保真图像生成失败，该页没有可导出内容。",
                                page=page_number,
                                severity="error",
                            )
                        ],
                    )
                )
            else:
                document.pages.append(
                    build_page_image_ir(
                        page_number=page_number,
                        width=width,
                        height=height,
                        image_bytes=image_bytes,
                        reason="文本层不完整或页面以图片为主",
                    )
                )
        else:
            document.pages.append(
                IRPage(
                    page_number=page_number,
                    width=width,
                    height=height,
                    route=route or "unknown",
                    editable=False,
                    warnings=[
                        IRWarning(
                            code="unknown_route",
                            message=f"未知的页面路线：{route}",
                            page=page_number,
                            severity="error",
                        )
                    ],
                )
            )
    return document
