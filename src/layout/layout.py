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

def extract_pdf_layout(
    source_pdf: Path,
    *,
    include_page_images: set[int] | None = None,
    image_max_pixels: int = DEFAULT_EMBEDDED_IMAGE_MAX_PIXELS,
    image_png_optimize: bool = DEFAULT_EMBEDDED_IMAGE_PNG_OPTIMIZE,
    image_jpeg_quality: int = DEFAULT_EMBEDDED_IMAGE_JPEG_QUALITY,
    include_fidelity: bool = False,
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
    try:
        for page_number, page in enumerate(document, start=1):
            width, height = (float(value) for value in page.get_size())
            text_page = page.get_textpage()
            try:
                styles = (
                    _extract_object_styles(
                        page,
                        text_page,
                        height,
                        font_descriptors=font_descriptors,
                        font_cache=font_cache,
                    )
                    if include_fidelity
                    else None
                )
                lines = _extract_text_lines(
                    text_page,
                    height,
                    styles=styles,
                )
                tables = _find_tables(page, lines, height)
                tables = (
                    *tables,
                    *_find_borderless_tables(
                        lines,
                        page_width=width,
                        page_height=height,
                        existing_tables=tables,
                    ),
                )
                tables = tuple(sorted(tables, key=lambda table: table.bbox[1]))
                vectors: tuple[PdfVectorObject, ...] = ()
                if include_fidelity:
                    vectors = _filter_table_border_vectors(
                        _extract_vector_objects(page, height),
                        tables,
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
                    )
                vectors: tuple[PdfVectorObject, ...] = ()
                if include_fidelity:
                    vectors = _filter_table_border_vectors(
                        _extract_vector_objects(page, height),
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
    updated_pages = tuple(
        replace(
            page,
            lines=_sort_lines_in_reading_order(
                page.lines,
                page.columns,
                page.width,
            ),
        )
        for page in updated_pages
    )
    if include_fidelity:
        updated_pages = _mark_logos(updated_pages)
    return PdfDocumentLayout(pages=_mark_table_continuations(updated_pages))


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


