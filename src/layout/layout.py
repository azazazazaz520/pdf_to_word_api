"""版面提取入口：把各阶段串成一次完整的页面解析。

被 worker 与基准脚本直接调用；内部按文本、矢量、图片、表格、阅读顺序
的顺序组织，并过滤掉已由 Word 表格承担的边框矢量。
"""

from __future__ import annotations
from pathlib import Path
from typing import Any
from .images import (
    DEFAULT_EMBEDDED_IMAGE_JPEG_QUALITY,
    DEFAULT_EMBEDDED_IMAGE_MAX_PIXELS,
    DEFAULT_EMBEDDED_IMAGE_PNG_OPTIMIZE,
    _extract_page_images,
)
from .models import (
    PdfDocumentLayout,
    PdfPageLayout,
    _PAGEOBJ_IMAGE,
    _PAGEOBJ_PATH,
    _PAGEOBJ_TEXT,
)
from .reading_order import (
    _build_content_blocks,
    _build_text_blocks,
    _annotate_line_metrics,
    _mark_logos,
    _mark_running_headers,
    _sort_lines_in_reading_order,
)
from .tables import (
    _find_borderless_tables,
    _find_tables,
    _line_in_tables,
    _mark_table_continuations,
)
from .text import _extract_object_styles, _extract_text_lines
from .vectors import _extract_vector_objects
from dataclasses import replace
import pypdfium2 as pdfium
from ..fonts.resolver import FontMatch, PdfFontDescriptor, extract_pdf_font_descriptors
from .models import PdfTable, PdfVectorObject
from .reading_order import _detect_column_boundaries, _estimate_body_left
from .regions import (
    assign_model_regions_to_lines,
    build_geometry_regions,
    normalize_model_regions,
    PageCoordinateTransform,
)

def extract_pdf_layout(
    source_pdf: Path,
    *,
    include_page_images: set[int] | None = None,
    image_max_pixels: int = DEFAULT_EMBEDDED_IMAGE_MAX_PIXELS,
    image_png_optimize: bool = DEFAULT_EMBEDDED_IMAGE_PNG_OPTIMIZE,
    image_jpeg_quality: int = DEFAULT_EMBEDDED_IMAGE_JPEG_QUALITY,
    include_fidelity: bool = False,
    block_reading_order_enabled: bool = True,
    block_reading_order_fallback_enabled: bool = True,
    model_regions_by_page: dict[int, Any] | None = None,
) -> PdfDocumentLayout:
    """提取文本行、图片、几何表格和矢量对象，并按多栏阅读顺序重排。

    ``include_fidelity=True`` 时额外提取字体、颜色、对齐、行距、
    z-order 和矢量图形，供高保真绝对定位导出使用。
    """
    if not source_pdf.is_file():
        raise FileNotFoundError(f"布局解析所需的 PDF 不存在：{source_pdf}")
    if image_max_pixels < 1:
        raise ValueError("内嵌图片像素上限必须大于 0")
    if not 1 <= image_jpeg_quality <= 100:
        raise ValueError("内嵌图片 JPEG 质量必须在 1 到 100 之间")

    font_descriptors: dict[str, PdfFontDescriptor] = (
        extract_pdf_font_descriptors(source_pdf) if include_fidelity else {}
    )
    font_cache: dict[tuple[Any, ...], FontMatch] = {}
    document = pdfium.PdfDocument(str(source_pdf))
    pages: list[PdfPageLayout] = []
    model_region_pages: list[tuple[Any, ...]] = []
    try:
        for page_number, page in enumerate(document, start=1):
            width, height = (float(value) for value in page.get_size())
            text_page = page.get_textpage()
            try:
                styles = _extract_object_styles(
                    page,
                    text_page,
                    height,
                    font_descriptors=font_descriptors,
                    font_cache=font_cache,
                    resolve_fonts=include_fidelity,
                )
                lines = _extract_text_lines(
                    text_page,
                    height,
                    styles=styles,
                    page_number=page_number,
                )
                model_regions: tuple[Any, ...] = ()
                model_page = (model_regions_by_page or {}).get(page_number)
                if model_page:
                    if isinstance(model_page, dict):
                        detections = (
                            model_page.get("detections")
                            or model_page.get("boxes")
                            or ()
                        )
                        pixel_width = float(model_page.get("pixel_width") or width)
                        pixel_height = float(model_page.get("pixel_height") or height)
                        rotation = int(model_page.get("rotation") or 0)
                    else:
                        detections = model_page
                        pixel_width = width
                        pixel_height = height
                        rotation = 0
                    model_regions = normalize_model_regions(
                        detections,
                        page_number=page_number,
                        transform=PageCoordinateTransform(
                            pixel_width,
                            pixel_height,
                            width,
                            height,
                            rotation=rotation,
                        ),
                        lines=lines,
                        include_source_char_ids=False,
                    )
                if (
                    include_page_images is not None
                    and (page_number - 1) not in include_page_images
                ):
                    images = ()
                else:
                    images = _extract_page_images(
                        page,
                        height,
                        page_number,
                        max_pixels=image_max_pixels,
                        png_optimize=image_png_optimize,
                        jpeg_quality=image_jpeg_quality,
                        include_raster_lines=True,
                    )
                tables = _find_tables(
                    page,
                    lines,
                    height,
                    raster_images=tuple(
                        image for image in images if image.source == "raster-line"
                    ),
                )
                tables = (
                    *tables,
                    *_find_borderless_tables(
                        lines,
                        page_width=width,
                        page_height=height,
                        existing_tables=tables,
                        confirmed_regions=model_regions,
                    ),
                )
                tables = tuple(sorted(tables, key=lambda table: table.bbox[1]))
                raw_vectors = _extract_vector_objects(page, height)
                table_border_bboxes = tuple(table.bbox for table in tables)
                images = tuple(
                    image
                    for image in images
                    if image.source != "raster-line"
                    or not any(
                        _bbox_contains(table_bbox, image.bbox, tolerance=2.0)
                        for table_bbox in table_border_bboxes
                    )
                )
                vectors: tuple[PdfVectorObject, ...] = ()
                if include_fidelity:
                    vectors = _filter_table_border_vectors(
                        raw_vectors,
                        tables,
                    )
            finally:
                text_page.close()
                page.close()
            pages.append(
                PdfPageLayout(
                    width=width,
                    height=height,
                    lines=tuple(lines),
                    tables=tables,
                    images=images,
                    vectors=vectors,
                )
            )
            model_region_pages.append(model_regions)
    finally:
        document.close()

    updated_pages = _mark_running_headers(tuple(pages))
    updated_pages = tuple(
        replace(
            page,
            columns=_detect_column_boundaries(
                list(page.lines),
                page.width,
                page.tables,
            ),
        )
        for page in updated_pages
    )
    region_pages: list[PdfPageLayout] = []
    for index, page in enumerate(updated_pages):
        geometry_regions = build_geometry_regions(
                page_number=index + 1,
                page_width=page.width,
                page_height=page.height,
                lines=page.lines,
                tables=page.tables,
                images=page.images,
                vectors=page.vectors,
                columns=page.columns,
        )
        model_regions = model_region_pages[index]
        assigned_lines, model_columns = assign_model_regions_to_lines(
            page.lines,
            model_regions=model_regions,
            page_width=page.width,
            tables=page.tables,
        )
        region_pages.append(
            replace(
                page,
                lines=assigned_lines,
                columns=model_columns or page.columns,
                regions=(*geometry_regions, *model_regions),
            )
        )
    updated_pages = tuple(region_pages)
    if include_fidelity:
        updated_pages = tuple(
            replace(
                page,
                lines=_annotate_line_metrics(
                    page.lines,
                    page_width=page.width,
                    columns=page.columns,
                    body_left=_estimate_body_left(page.lines),
                ),
            )
            for page in updated_pages
        )
    updated_pages = _mark_table_continuations(updated_pages)
    updated_pages = _mark_logos(updated_pages)
    ordered_pages: list[PdfPageLayout] = []
    for page in updated_pages:
        block_input_lines = page.lines
        if not block_reading_order_enabled:
            block_input_lines = _sort_lines_in_reading_order(
                page.lines,
                page.columns,
                page.width,
            )
        text_blocks = _build_text_blocks(
            block_input_lines,
            boundaries=page.columns,
            page_width=page.width,
            tables=page.tables,
        )
        content_blocks, order_confidence, order_warnings = _build_content_blocks(
            text_blocks,
            tables=page.tables,
            images=page.images,
            vectors=page.vectors,
            boundaries=page.columns,
            page_width=page.width,
            page_height=page.height,
            reading_order_enabled=block_reading_order_enabled,
            fallback_enabled=block_reading_order_fallback_enabled,
        )
        ordered_text_blocks = tuple(
            item.text_block
            for item in content_blocks
            if item.text_block is not None
        )
        ordered_lines = tuple(
            line
            for block in ordered_text_blocks
            for line in block.lines
        )
        ordered_pages.append(
            replace(
                page,
                lines=ordered_lines,
                text_blocks=ordered_text_blocks,
                content_blocks=content_blocks,
                reading_order_confidence=order_confidence,
                reading_order_warnings=order_warnings,
            )
        )
    updated_pages = tuple(ordered_pages)
    return PdfDocumentLayout(pages=updated_pages)


def _bbox_contains(
    outer: tuple[float, float, float, float],
    inner: tuple[float, float, float, float],
    *,
    tolerance: float,
) -> bool:
    return (
        inner[0] >= outer[0] - tolerance
        and inner[1] >= outer[1] - tolerance
        and inner[2] <= outer[2] + tolerance
        and inner[3] <= outer[3] + tolerance
    )


def _filter_table_border_vectors(
    vectors: tuple[PdfVectorObject, ...],
    tables: tuple[PdfTable, ...],
) -> tuple[PdfVectorObject, ...]:
    """剔除完全落在表格内部的边框线段，避免与 Word 表格边框重叠。"""
    if not tables:
        return vectors
    filtered: list[PdfVectorObject] = []
    for vector in vectors:
        inside_table = False
        for table in tables:
            x0, top, x1, bottom = table.bbox
            if (
                vector.bbox[0] >= x0 - 2.0
                and vector.bbox[1] >= top - 2.0
                and vector.bbox[2] <= x1 + 2.0
                and vector.bbox[3] <= bottom + 2.0
            ):
                inside_table = True
                break
        if not inside_table:
            filtered.append(vector)
    return tuple(filtered)


