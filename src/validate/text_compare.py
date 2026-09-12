"""文本层面的对比：字符覆盖率与行级 bbox、字号误差。

覆盖率按字符 n-gram 统计，用于发现文字丢失；行级比对用于发现位置与字号偏差。
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any


def _page_text_lines(
    pdf_path: Path,
    page_index: int,
) -> list[Any]:
    import pypdfium2 as pdfium

    from ..layout.text import _extract_object_styles, _extract_text_lines

    document = pdfium.PdfDocument(str(pdf_path))
    try:
        page = document[page_index]
        try:
            width, height = (float(value) for value in page.get_size())
            text_page = page.get_textpage()
            try:
                styles = _extract_object_styles(page, text_page, height)
                return list(
                    _extract_text_lines(text_page, height, styles=styles)
                )
            finally:
                text_page.close()
        finally:
            page.close()
    finally:
        document.close()


def _percentile(ordered_values: list[float], fraction: float) -> float:
    """线性插值分位数（输入需已排序）。"""
    if not ordered_values:
        return 0.0
    if len(ordered_values) == 1:
        return float(ordered_values[0])
    position = max(0.0, min(1.0, fraction)) * (len(ordered_values) - 1)
    lower = int(position)
    upper = min(lower + 1, len(ordered_values) - 1)
    weight = position - lower
    return float(ordered_values[lower]) * (1 - weight) + float(
        ordered_values[upper]
    ) * weight


def _normalized_text(value: str) -> str:
    return "".join(str(value or "").split()).lower()


def compare_text_coverage(
    source_pdf: Path,
    rendered_pdf: Path,
    *,
    ngram: int = 3,
    page_limit: int | None = None,
) -> dict[str, Any]:
    """用字符 n-gram 覆盖率判断"源页文字是否被重建"。

    比逐行匹配稳健：同一行被拆成多个定位帧、空格/标点微调都不影响。
    """
    from pypdf import PdfReader

    started = time.perf_counter()
    source = PdfReader(str(source_pdf))
    rendered = PdfReader(str(rendered_pdf))
    page_count = min(len(source.pages), len(rendered.pages))
    if page_limit is not None:
        page_count = min(page_count, max(0, int(page_limit)))
    pages: list[dict[str, Any]] = []
    coverages: list[float] = []
    for index in range(page_count):
        try:
            source_text = _normalized_text(source.pages[index].extract_text() or "")
        except Exception:
            source_text = ""
        try:
            rendered_text = _normalized_text(
                rendered.pages[index].extract_text() or ""
            )
        except Exception:
            rendered_text = ""
        if not source_text:
            pages.append({"page": index + 1, "coverage": 1.0, "source_chars": 0})
            coverages.append(1.0)
            continue
        grams = [
            source_text[position : position + ngram]
            for position in range(max(len(source_text) - ngram + 1, 0))
        ]
        if not grams:
            coverage = 1.0 if source_text in rendered_text else 0.0
        else:
            hits = sum(1 for gram in grams if gram in rendered_text)
            coverage = hits / len(grams)
        coverages.append(coverage)
        pages.append(
            {
                "page": index + 1,
                "coverage": round(coverage, 4),
                "source_chars": len(source_text),
            }
        )
    below = [entry["page"] for entry in pages if entry["coverage"] < 0.9]
    return {
        "status": "succeeded",
        "page_count": page_count,
        "ngram": ngram,
        "mean_coverage": (
            round(sum(coverages) / len(coverages), 4) if coverages else None
        ),
        "min_coverage": round(min(coverages), 4) if coverages else None,
        "pages_below_threshold": below,
        "pages": pages,
        "elapsed_seconds": round(time.perf_counter() - started, 3),
    }


def _common_prefix_length(left: str, right: str) -> int:
    limit = min(len(left), len(right))
    index = 0
    while index < limit and left[index] == right[index]:
        index += 1
    return index


def compare_text_layout(
    source_pdf: Path,
    rendered_pdf: Path,
    *,
    max_pages: int | None = None,
    bbox_tolerance: float = 3.0,
    font_tolerance: float = 0.5,
    page_indices: list[int] | None = None,
) -> dict[str, Any]:
    """逐页匹配文本行，统计 bbox 与字号误差。"""
    from pypdf import PdfReader

    started = time.perf_counter()
    source = PdfReader(str(source_pdf))
    rendered = PdfReader(str(rendered_pdf))
    page_count = min(len(source.pages), len(rendered.pages))
    if max_pages is not None:
        page_count = min(page_count, max(0, int(max_pages)))
    selected = (
        [int(index) for index in page_indices]
        if page_indices is not None
        else None
    )
    indices = (
        [index for index in selected if 0 <= index < page_count]
        if selected is not None
        else list(range(page_count))
    )
    pages: list[dict[str, Any]] = []
    all_bbox_errors: list[float] = []
    all_font_errors: list[float] = []
    unmatched = 0
    for page_index in indices:
        source_lines = _page_text_lines(source_pdf, page_index)
        rendered_lines = _page_text_lines(rendered_pdf, page_index)
        used: set[int] = set()
        bbox_errors: list[float] = []
        font_errors: list[float] = []
        page_unmatched = 0
        for line in source_lines:
            key = _normalized_text(line.text)
            if not key:
                continue
            best_index = None
            best_distance = None
            best_exact = False
            best_prefix = 0
            for position, candidate in enumerate(rendered_lines):
                if position in used:
                    continue
                candidate_key = _normalized_text(candidate.text)
                if not candidate_key:
                    continue
                exact = candidate_key == key
                prefix = _common_prefix_length(key, candidate_key)
                # 允许"一行被拆成多个 span 帧"或轻微空格差异的情况
                if not exact and prefix < min(4, len(key)):
                    continue
                distance = abs(candidate.center_y - line.center_y) + abs(
                    candidate.x0 - line.x0
                )
                if (
                    best_index is None
                    or (exact, prefix) > (best_exact, best_prefix)
                    or (
                        (exact, prefix) == (best_exact, best_prefix)
                        and (best_distance is None or distance < best_distance)
                    )
                ):
                    best_index = position
                    best_distance = distance
                    best_exact = exact
                    best_prefix = prefix
            if best_index is None:
                page_unmatched += 1
                continue
            used.add(best_index)
            candidate = rendered_lines[best_index]
            if best_exact:
                bbox_error = max(
                    abs(candidate.x0 - line.x0),
                    abs(candidate.top - line.top),
                    abs(candidate.x1 - line.x1),
                    abs(candidate.bottom - line.bottom),
                )
            else:
                # 部分匹配时只比较左上角，避免右边界失真
                bbox_error = max(
                    abs(candidate.x0 - line.x0),
                    abs(candidate.top - line.top),
                )
            bbox_errors.append(bbox_error)
            if line.font_size > 0 and candidate.font_size > 0:
                font_errors.append(abs(candidate.font_size - line.font_size))
        all_bbox_errors.extend(bbox_errors)
        all_font_errors.extend(font_errors)
        unmatched += page_unmatched
        ordered_errors = sorted(bbox_errors)
        pages.append(
            {
                "page": page_index + 1,
                "matched_line_count": len(bbox_errors),
                "unmatched_line_count": page_unmatched,
                "max_bbox_error": (
                    round(max(bbox_errors), 3) if bbox_errors else 0.0
                ),
                "mean_bbox_error": (
                    round(sum(bbox_errors) / len(bbox_errors), 3)
                    if bbox_errors
                    else 0.0
                ),
                "median_bbox_error": (
                    round(_percentile(ordered_errors, 0.5), 3)
                    if ordered_errors
                    else 0.0
                ),
                "p90_bbox_error": (
                    round(_percentile(ordered_errors, 0.9), 3)
                    if ordered_errors
                    else 0.0
                ),
                "max_font_size_error": (
                    round(max(font_errors), 3) if font_errors else 0.0
                ),
                "passes_bbox": (
                    _percentile(ordered_errors, 0.9) < bbox_tolerance
                    if ordered_errors
                    else True
                ),
                "passes_font_size": all(
                    error < font_tolerance for error in font_errors
                ),
            }
        )
    ordered_all = sorted(all_bbox_errors)
    return {
        "status": "succeeded",
        "page_count": page_count,
        "matched_line_count": len(all_bbox_errors),
        "unmatched_line_count": unmatched,
        "max_bbox_error": (
            round(max(all_bbox_errors), 3) if all_bbox_errors else 0.0
        ),
        "mean_bbox_error": (
            round(sum(all_bbox_errors) / len(all_bbox_errors), 3)
            if all_bbox_errors
            else 0.0
        ),
        "median_bbox_error": (
            round(_percentile(ordered_all, 0.5), 3) if ordered_all else 0.0
        ),
        "p90_bbox_error": (
            round(_percentile(ordered_all, 0.9), 3) if ordered_all else 0.0
        ),
        "max_font_size_error": (
            round(max(all_font_errors), 3) if all_font_errors else 0.0
        ),
        "bbox_tolerance": bbox_tolerance,
        "font_size_tolerance": font_tolerance,
        "passes_bbox": (
            _percentile(ordered_all, 0.9) < bbox_tolerance
            if ordered_all
            else True
        ),
        "passes_font_size": all(
            error < font_tolerance for error in all_font_errors
        ),
        "pages": pages,
        "elapsed_seconds": round(time.perf_counter() - started, 3),
    }
