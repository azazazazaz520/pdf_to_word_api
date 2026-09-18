"""阅读顺序、页眉页脚与 logo 判定。

依据行坐标识别多栏边界并排序，标记跨页重复的页眉页脚与重复出现的小图片。
"""

from __future__ import annotations

from bisect import bisect_right
from collections import Counter
from collections.abc import Iterable
from dataclasses import dataclass, replace
import math
import re
from typing import Any

from .models import (
    PdfImageBlock,
    PdfContentBlock,
    PdfPageLayout,
    PdfTable,
    PdfTextBlock,
    PdfTextLine,
    PdfVectorObject,
    _MIN_TABLE_WIDTH,
)
from .tables import _line_in_table


@dataclass(frozen=True)
class ReadingOrderConfig:
    """集中维护新增阅读顺序判断使用的阈值。"""

    min_column_candidates: int = 6
    column_min_gap: float = 12.0
    column_center_min_ratio: float = 0.2
    column_center_max_ratio: float = 0.8
    column_merge_gap: float = 24.0
    full_width_ratio: float = 0.65
    min_horizontal_overlap: float = 0.12
    block_gap_factor: float = 1.75
    font_size_delta_factor: float = 0.35
    font_size_delta_min: float = 2.0
    horizontal_alignment_factor: float = 2.5
    side_by_side_overlap: float = 0.25
    side_by_side_top_delta_factor: float = 0.75
    caption_min_horizontal_overlap: float = 0.12
    caption_center_tolerance_factor: float = 2.0
    caption_min_center_distance: float = 18.0
    caption_min_gap: float = 14.0
    caption_gap_factor: float = 3.5
    title_size_delta: float = 1.8
    short_text_limit: int = 80
    title_center_ratio: float = 0.12
    title_center_min_distance: float = 24.0
    footnote_top_ratio: float = 0.78
    footnote_size_ratio: float = 0.9


DEFAULT_READING_ORDER_CONFIG = ReadingOrderConfig()


def _is_page_number_text(value: str) -> bool:
    return bool(re.fullmatch(r"\s*(?:第\s*)?\d{1,4}\s*(?:页)?\s*", value))


def _same_text_direction(left: PdfTextLine, right: PdfTextLine) -> bool:
    """旋转方向不同的文本不参与同一视觉行归并。"""
    delta = abs(
        (float(left.rotation) - float(right.rotation) + 180.0) % 360.0
        - 180.0
    )
    return delta <= 8.0


def _running_text_key(value: str) -> str:
    normalized = re.sub(r"\d+", "#", value)
    normalized = re.sub(r"\s+", "", normalized)
    return normalized[:80]


def _is_running_candidate(line: PdfTextLine, page_height: float) -> bool:
    """只在页面顶部或底部识别页眉页脚和页码，避免误伤表格中的数字。"""
    return line.top <= page_height * 0.12 or line.bottom >= page_height * 0.88


def _mark_running_headers(
    pages: tuple[PdfPageLayout, ...],
) -> tuple[PdfPageLayout, ...]:
    """识别跨页重复出现的页眉、页脚和页码，并标记为运行内容。"""
    if not pages:
        return pages
    counts: Counter[str] = Counter()
    entries: list[tuple[int, int, str]] = []
    for page_index, page in enumerate(pages):
        for line_index, line in enumerate(page.lines):
            if not _is_running_candidate(line, page.height):
                continue
            key = _running_text_key(line.text)
            if not key:
                continue
            counts[key] += 1
            entries.append((page_index, line_index, key))
    if not entries:
        return pages
    if len(pages) == 1:
        marked = {
            (page_index, line_index)
            for page_index, line_index, _ in entries
            if _is_page_number_text(pages[page_index].lines[line_index].text)
        }
    else:
        threshold = max(2, int(len(pages) * 0.5 + 0.999))
        marked = {
            (page_index, line_index)
            for page_index, line_index, key in entries
            if counts[key] >= threshold
        }
    if not marked:
        return pages
    updated = list(pages)
    for page_index in {item[0] for item in marked}:
        page = pages[page_index]
        lines = tuple(
            replace(line, is_header_footer=True)
            if (page_index, line_index) in marked
            else line
            for line_index, line in enumerate(page.lines)
        )
        updated[page_index] = replace(page, lines=lines)
    return tuple(updated)


def _detect_column_boundaries(
    lines: list[PdfTextLine],
    page_width: float,
    tables: tuple[PdfTable, ...],
    config: ReadingOrderConfig = DEFAULT_READING_ORDER_CONFIG,
) -> tuple[float, ...]:
    """通过跨行稳定空白带检测多栏分栏边界。"""
    body_sizes = sorted(
        float(line.font_size)
        for line in lines
        if not line.is_header_footer and float(line.font_size) > 0.0
    )
    body_size = body_sizes[len(body_sizes) // 2] if body_sizes else 0.0
    candidates = [
        line
        for line in lines
        if not line.is_header_footer
        and not _line_in_table(line, tables)
        and line.x1 - line.x0 < page_width * config.full_width_ratio
        and not (
            body_size > 0.0
            and line.font_size >= body_size + config.title_size_delta
            and len(line.text.strip()) <= config.short_text_limit
            and not line.text.strip().endswith(
                ("。", ".", "；", ";", "：", ":")
            )
        )
    ]
    if len(candidates) < config.min_column_candidates:
        return ()
    intervals = [(line.x0, line.x1) for line in candidates]
    edges = sorted({value for interval in intervals for value in interval})
    gaps: list[tuple[float, float]] = []
    for left, right in zip(edges, edges[1:]):
        # 分栏留白由整栏文字让出，十几点的间距已是明显分栏；
        # 更窄的间距属于行内字距，不作为分栏依据。
        if right - left < config.column_min_gap:
            continue
        midpoint = (left + right) / 2.0
        if not (
            page_width * config.column_center_min_ratio
            <= midpoint
            <= page_width * config.column_center_max_ratio
        ):
            continue
        if any(x0 - 1.0 <= midpoint <= x1 + 1.0 for x0, x1 in intervals):
            continue
        left_count = sum(line.center_x < midpoint for line in candidates)
        right_count = sum(line.center_x > midpoint for line in candidates)
        if left_count >= 2 and right_count >= 2:
            gaps.append((left, right))
    boundaries: list[float] = []
    for left, right in gaps:
        boundary = (left + right) / 2.0
        if not boundaries or boundary - boundaries[-1] >= config.column_merge_gap:
            boundaries.append(boundary)
    return tuple(boundaries)


def _same_visual_row(left: PdfTextLine, right: PdfTextLine) -> bool:
    """判断两行是否属于同一视觉行。

    同一行文字若被拆成多段，各段的上界会因字形高低相差一两个点，而下界
    （基线）基本重合。仅按上界排序会把同一行靠右的字排到靠左的字前面，
    成品因此出现姓名互换这类顺序错乱。上界或基线落在行高的一小段以内
    即视为同一行，正文行的行距远大于该容差，不会被并入。
    """
    if not _same_text_direction(left, right):
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


def _row_tolerant_sort(
    lines: list[PdfTextLine],
) -> list[PdfTextLine]:
    """按阅读次序排序文本行，同一视觉行内按横向次序。

    直接按 ``(top, x0)`` 排序时，同行各段的上界差异（字形高低造成的
    零点几到几个点）会成为主键，使同行靠右的文字排到靠左的文字之前，
    成品因此出现姓名互换这类顺序错乱。这里改为先按上界扫描出视觉行，
    行内按横向排序，再按视觉行输出。
    """
    ordered = sorted(lines, key=lambda line: (line.top, line.x0))
    result: list[PdfTextLine] = []
    index = 0
    while index < len(ordered):
        end = index + 1
        while end < len(ordered) and _same_visual_row(ordered[index], ordered[end]):
            end += 1
        result.extend(sorted(ordered[index:end], key=lambda line: line.x0))
        index = end
    return result


def _normalize_row_order(
    lines: tuple[PdfTextLine, ...],
) -> tuple[PdfTextLine, ...]:
    """把同一视觉行内的文本行按横向次序排好。

    输入需已按 ``(top, x0)`` 排序。同一视觉行可能由多段文字组成
    （同一行上的多个姓名、被拆成多个文本对象的行），此时横向次序才是
    阅读次序。分组以组内最靠上的行为基准，避免相邻比较造成连锁归并。
    """
    result = list(lines)
    index = 0
    while index < len(result):
        end = index + 1
        while end < len(result) and _same_visual_row(result[index], result[end]):
            end += 1
        if end - index > 1:
            result[index:end] = sorted(
                result[index:end], key=lambda line: line.x0
            )
        index = end
    return tuple(result)


def _sort_lines_in_reading_order(
    lines: tuple[PdfTextLine, ...],
    boundaries: tuple[float, ...],
    page_width: float,
    config: ReadingOrderConfig = DEFAULT_READING_ORDER_CONFIG,
) -> tuple[PdfTextLine, ...]:
    """按分栏边界重排文本行；跨栏标题保持在全宽位置。"""
    ordered_lines = _normalize_row_order(
        sorted(lines, key=lambda line: (line.top, line.x0))
    )
    if not boundaries:
        return ordered_lines
    separator_ids = {
        id(line)
        for line in ordered_lines
        if line.x1 - line.x0 >= page_width * config.full_width_ratio
        or any(line.x0 < boundary < line.x1 for boundary in boundaries)
    }
    groups: list[list[PdfTextLine]] = []
    current_group: list[PdfTextLine] = []
    for line in ordered_lines:
        if id(line) in separator_ids:
            if current_group:
                groups.append(current_group)
                current_group = []
            groups.append([line])
        else:
            current_group.append(line)
    if current_group:
        groups.append(current_group)
    result: list[PdfTextLine] = []
    for group in groups:
        if len(group) == 1 and id(group[0]) in separator_ids:
            result.extend(group)
            continue
        result.extend(
            sorted(
                group,
                key=lambda line: (
                    bisect_right(boundaries, line.center_x),
                    line.top,
                    line.x0,
                ),
            )
        )
    return tuple(result)


def _line_column_index(
    line: PdfTextLine,
    boundaries: tuple[float, ...],
    page_width: float,
    config: ReadingOrderConfig = DEFAULT_READING_ORDER_CONFIG,
) -> int:
    if not boundaries:
        return 0
    if line.width >= page_width * config.full_width_ratio or any(
        line.x0 < boundary < line.x1 for boundary in boundaries
    ):
        return -1
    return bisect_right(boundaries, line.center_x)


def _region_for_bbox(
    bbox: tuple[float, float, float, float],
    *,
    boundaries: tuple[float, ...],
    page_width: float,
    config: ReadingOrderConfig = DEFAULT_READING_ORDER_CONFIG,
) -> str:
    """根据页面局部栏位给内容块分配稳定的区域编号。"""
    x0, _, x1, _ = bbox
    if (
        x1 - x0 >= page_width * config.full_width_ratio
        or any(x0 < boundary < x1 for boundary in boundaries)
    ):
        return "full-width"
    if not boundaries:
        return "region-0"
    return f"column-{bisect_right(boundaries, (x0 + x1) / 2.0)}"


def _vertical_overlap_ratio(
    left: Any,
    right: Any,
) -> float:
    overlap = max(min(left.bottom, right.bottom) - max(left.top, right.top), 0.0)
    denominator = max(min(left.bottom - left.top, right.bottom - right.top), 1.0)
    return overlap / denominator


def _text_block_size(block: PdfTextBlock) -> float:
    sizes = sorted(
        float(line.font_size)
        for line in block.lines
        if float(line.font_size) > 0.0
    )
    return sizes[len(sizes) // 2] if sizes else max(block.height, 1.0)


def _looks_like_formula_block(value: str) -> bool:
    text = value.strip()
    if not text or len(text) > 120:
        return False
    if re.search(r"[{};]", text):
        return False
    if not re.search(r"[=≤≥≈∑∏∫√∞∂∇α-ωΑ-Ω₀-₉⁰-⁹]", text):
        return False
    return bool(re.search(r"[+\-−·*/^_()=≤≥≈]", text))


def _caption_distance(
    caption: PdfTextBlock,
    target_bbox: tuple[float, float, float, float],
    *,
    config: ReadingOrderConfig,
) -> tuple[str, float] | None:
    target_x0, target_top, target_x1, target_bottom = target_bbox
    horizontal_overlap = max(
        min(caption.x1, target_x1) - max(caption.x0, target_x0),
        0.0,
    )
    horizontal = horizontal_overlap / max(
        min(caption.width, target_x1 - target_x0),
        1.0,
    )
    if horizontal < config.caption_min_horizontal_overlap and abs(
        caption.center_x - (target_x0 + target_x1) / 2.0
    ) > max(
        caption.height * config.caption_center_tolerance_factor,
        config.caption_min_center_distance,
    ):
        return None
    size = max(_text_block_size(caption), 1.0)
    max_gap = max(config.caption_min_gap, size * config.caption_gap_factor)
    if caption.bottom <= target_top:
        gap = target_top - caption.bottom
        if gap <= max_gap:
            return "before", gap
    if target_bottom <= caption.top:
        gap = caption.top - target_bottom
        if gap <= max_gap:
            return "after", gap
    return None


def _classify_text_blocks(
    blocks: tuple[PdfTextBlock, ...],
    *,
    tables: tuple[PdfTable, ...] = (),
    images: tuple[PdfImageBlock, ...] = (),
    page_width: float,
    page_height: float,
    boundaries: tuple[float, ...],
    config: ReadingOrderConfig = DEFAULT_READING_ORDER_CONFIG,
) -> tuple[PdfTextBlock, ...]:
    """给文本块标记类型、区域和可信度。"""
    body_sizes = sorted(
        _text_block_size(block)
        for block in blocks
        if not block.is_header_footer and block.text.strip()
    )
    body_size = body_sizes[len(body_sizes) // 2] if body_sizes else 0.0
    classified: list[PdfTextBlock] = []
    for block in blocks:
        text = block.text.strip()
        block_type = "TEXT"
        confidence = 0.82
        needs_review = False
        if block.is_header_footer:
            if _is_page_number_text(text):
                block_type = "PAGE_NUMBER"
            elif block.center_y <= page_height / 2.0:
                block_type = "HEADER"
            else:
                block_type = "FOOTER"
            confidence = 0.98
        elif any(
            _caption_distance(block, target.bbox, config=config) is not None
            for target in (*tables, *images)
        ) and len(block.lines) <= 2 and len(text) <= 120:
            block_type = "CAPTION"
            confidence = 0.78
        elif _looks_like_formula_block(text) and len(block.lines) <= 3:
            block_type = "FORMULA"
            confidence = 0.76
        elif (
            len(block.lines) <= 3
            and len(text) <= config.short_text_limit
            and _text_block_size(block) >= body_size + config.title_size_delta
            and not text.endswith(("。", ".", "；", ";", "：", ":"))
        ):
            block_type = "TITLE"
            confidence = 0.74
        elif (
            body_size > 0.0
            and block.top >= page_height * config.footnote_top_ratio
            and _text_block_size(block) <= body_size * config.footnote_size_ratio
            and len(text) >= 4
        ):
            block_type = "FOOTNOTE"
            confidence = 0.64
        region_id = _region_for_bbox(
            block.bbox,
            boundaries=boundaries,
            page_width=page_width,
            config=config,
        )
        if (
            block_type == "TITLE"
            and abs((block.x0 + block.x1) / 2.0 - page_width / 2.0)
            <= max(
                page_width * config.title_center_ratio,
                config.title_center_min_distance,
            )
        ):
            region_id = "full-width"
        if block_type in {"TITLE", "CAPTION", "FOOTNOTE"} and confidence < 0.7:
            needs_review = True
        classified.append(
            replace(
                block,
                block_type=block_type,
                confidence=confidence,
                region_id=region_id,
                needs_review=needs_review,
            )
        )
    return tuple(classified)


def _text_block_from_lines(
    lines: list[PdfTextLine],
    column_index: int,
) -> PdfTextBlock:
    return PdfTextBlock(
        lines=tuple(lines),
        bbox=(
            min(line.x0 for line in lines),
            min(line.top for line in lines),
            max(line.x1 for line in lines),
            max(line.bottom for line in lines),
        ),
        column_index=column_index,
        direction=(
            _dominant_direction(lines)
            if lines
            else (1.0, 0.0)
        ),
        rotation=_dominant_rotation(lines),
        is_header_footer=all(line.is_header_footer for line in lines),
        source_char_ids=tuple(
            dict.fromkeys(
                char_id
                for line in lines
                for char_id in getattr(line, "source_char_ids", ())
            )
        ),
    )


def _dominant_rotation(lines: list[PdfTextLine]) -> float:
    if not lines:
        return 0.0
    weights = [max(float(line.font_size), 1.0) for line in lines]
    x_total = sum(
        math.cos(math.radians(line.rotation)) * weight
        for line, weight in zip(lines, weights)
    )
    y_total = sum(
        math.sin(math.radians(line.rotation)) * weight
        for line, weight in zip(lines, weights)
    )
    if abs(x_total) <= 1e-6 and abs(y_total) <= 1e-6:
        return float(lines[0].rotation)
    angle = math.degrees(math.atan2(y_total, x_total))
    angle = (angle + 180.0) % 360.0 - 180.0
    if abs(angle + 180.0) <= 0.05:
        angle = 180.0
    if abs(angle) <= 0.05:
        angle = 0.0
    return round(angle, 2)


def _dominant_direction(lines: list[PdfTextLine]) -> tuple[float, float]:
    rotation = _dominant_rotation(lines)
    return _rotation_direction(rotation)


def _rotation_direction(rotation: float) -> tuple[float, float]:
    return (
        round(math.cos(math.radians(rotation)), 6),
        round(math.sin(math.radians(rotation)), 6),
    )


def _line_projection_interval(
    line: PdfTextLine,
    axis: tuple[float, float],
) -> tuple[float, float]:
    points = (
        (line.x0, line.top),
        (line.x0, line.bottom),
        (line.x1, line.top),
        (line.x1, line.bottom),
    )
    values = [point[0] * axis[0] + point[1] * axis[1] for point in points]
    return min(values), max(values)


def _can_extend_text_block(
    previous: PdfTextLine,
    current: PdfTextLine,
    previous_column: int,
    current_column: int,
    *,
    config: ReadingOrderConfig = DEFAULT_READING_ORDER_CONFIG,
) -> bool:
    if previous_column != current_column:
        return False
    if previous.is_header_footer != current.is_header_footer:
        return False
    if not _same_text_direction(previous, current):
        return False
    if previous.font_size > 0.0 and current.font_size > 0.0:
        size_delta = abs(previous.font_size - current.font_size)
        size_limit = max(
            config.font_size_delta_min,
            max(previous.font_size, current.font_size)
            * config.font_size_delta_factor,
        )
        if size_delta > size_limit:
            return False
    scale = max(
        float(previous.font_size or 0.0),
        float(current.font_size or 0.0),
        previous.height,
        current.height,
        1.0,
    )
    direction = _rotation_direction(previous.rotation)
    normal = (-direction[1], direction[0])
    previous_low, previous_high = _line_projection_interval(previous, normal)
    current_low, current_high = _line_projection_interval(current, normal)
    if current_low >= previous_high:
        gap = current_low - previous_high
    elif previous_low >= current_high:
        gap = previous_low - current_high
    else:
        gap = 0.0
    overlap = max(
        min(previous.x1, current.x1) - max(previous.x0, current.x0),
        0.0,
    )
    overlap_ratio = overlap / max(min(previous.width, current.width), 1.0)
    same_left = abs(previous.x0 - current.x0) <= max(
        8.0,
        scale * config.horizontal_alignment_factor,
    )
    same_center = abs(previous.center_x - current.center_x) <= max(
        12.0,
        scale * config.horizontal_alignment_factor * 1.5,
    )
    if (
        overlap_ratio < config.min_horizontal_overlap
        and not same_left
        and not same_center
    ):
        return False
    return gap <= max(8.0, scale * config.block_gap_factor)


def _build_text_blocks(
    lines: tuple[PdfTextLine, ...],
    *,
    boundaries: tuple[float, ...],
    page_width: float,
    tables: tuple[PdfTable, ...] = (),
    config: ReadingOrderConfig = DEFAULT_READING_ORDER_CONFIG,
) -> tuple[PdfTextBlock, ...]:
    """先按几何相邻关系建立文本块，不依赖整页阅读顺序。"""
    ordered_lines = sorted(
        lines,
        key=lambda line: (line.top, line.x0, line.z_order, line.text),
    )
    grouped: list[list[PdfTextLine]] = []
    columns: list[int] = []
    for line in ordered_lines:
        column = _line_column_index(
            line,
            boundaries,
            page_width,
            config=config,
        )
        candidates: list[tuple[float, float, int]] = []
        for index, group in enumerate(grouped):
            previous = group[-1]
            if _line_in_table(previous, tables) != _line_in_table(line, tables):
                continue
            if not _can_extend_text_block(
                previous,
                line,
                columns[index],
                column,
                config=config,
            ):
                continue
            gap = max(
                previous.top - line.bottom,
                line.top - previous.bottom,
                0.0,
            )
            candidates.append((gap, abs(previous.top - line.top), index))
        if candidates:
            _, _, selected = min(candidates)
            grouped[selected].append(line)
            continue
        grouped.append([line])
        columns.append(column)

    blocks: list[PdfTextBlock] = []
    for index, (group, column) in enumerate(zip(grouped, columns)):
        normalized = _normalize_row_order(
            tuple(sorted(group, key=lambda line: (line.top, line.x0)))
        )
        block = _text_block_from_lines(list(normalized), column)
        blocks.append(
            replace(
                block,
                block_id=f"text-{index}",
                region_id=_region_for_bbox(
                    block.bbox,
                    boundaries=boundaries,
                    page_width=page_width,
                    config=config,
                ),
            )
        )
    return tuple(blocks)


def _content_block_layer(
    block_type: str,
) -> str:
    if block_type in {"HEADER"}:
        return "header"
    if block_type in {"FOOTER", "PAGE_NUMBER"}:
        return "footer"
    return "body"


def _region_sort_key(region_id: str) -> tuple[int, str]:
    match = re.fullmatch(r"column-(\d+)", region_id)
    if match:
        return int(match.group(1)), region_id
    return 999, region_id


def _build_content_blocks(
    text_blocks: tuple[PdfTextBlock, ...],
    *,
    tables: tuple[PdfTable, ...],
    images: tuple[PdfImageBlock, ...],
    vectors: tuple[PdfVectorObject, ...],
    boundaries: tuple[float, ...],
    page_width: float,
    page_height: float,
    config: ReadingOrderConfig = DEFAULT_READING_ORDER_CONFIG,
    reading_order_enabled: bool = True,
    fallback_enabled: bool = True,
) -> tuple[tuple[PdfContentBlock, ...], float, tuple[str, ...]]:
    """把文字、表格、图片和矢量对象放入同一条阅读顺序输入。"""
    classified = _classify_text_blocks(
        text_blocks,
        tables=tables,
        images=images,
        page_width=page_width,
        page_height=page_height,
        boundaries=boundaries,
        config=config,
    )
    items: list[PdfContentBlock] = []
    for index, block in enumerate(classified):
        block_id = block.block_id or f"text-{index}"
        block = replace(block, block_id=block_id)
        items.append(
            PdfContentBlock(
                block_id=block_id,
                block_type=block.block_type,
                bbox=block.bbox,
                text_block=block,
                source="text_layout",
                confidence=block.confidence,
                region_id=block.region_id,
                layer=_content_block_layer(block.block_type),
                needs_review=block.needs_review,
                source_char_ids=block.source_char_ids,
            )
        )
    for index, table in enumerate(tables):
        items.append(
            PdfContentBlock(
                block_id=f"table-{index}",
                block_type="TABLE",
                bbox=table.bbox,
                table=table,
                source="vector_table",
                region_id=_region_for_bbox(
                    table.bbox,
                    boundaries=boundaries,
                    page_width=page_width,
                    config=config,
                ),
                source_char_ids=table.source_char_ids,
            )
        )
    for index, image in enumerate(images):
        items.append(
            PdfContentBlock(
                block_id=f"image-{index}",
                block_type="IMAGE",
                bbox=image.bbox,
                image=image,
                source=image.source,
                region_id=_region_for_bbox(
                    image.bbox,
                    boundaries=boundaries,
                    page_width=page_width,
                    config=config,
                ),
                layer=image.layer,
                source_char_ids=(),
            )
        )
    for index, vector in enumerate(vectors):
        items.append(
            PdfContentBlock(
                block_id=f"vector-{index}",
                block_type="VECTOR",
                bbox=vector.bbox,
                vector=vector,
                source="pdf_vector",
                region_id=_region_for_bbox(
                    vector.bbox,
                    boundaries=boundaries,
                    page_width=page_width,
                    config=config,
                ),
                layer=vector.layer,
                source_char_ids=(),
            )
        )
    return _sort_content_blocks_in_reading_order(
        tuple(items),
        boundaries=boundaries,
        page_width=page_width,
        config=config,
        reading_order_enabled=reading_order_enabled,
        fallback_enabled=fallback_enabled,
    )


def _add_order_edge(
    edges: list[set[int]],
    source: int,
    target: int,
) -> None:
    if source != target:
        edges[source].add(target)


def _is_body_item(item: PdfContentBlock) -> bool:
    return item.layer not in {"header", "footer"}


def _participates_in_reading_order(item: PdfContentBlock) -> bool:
    """装饰性矢量对象参与输出，但不单独制造文字先后关系。"""
    return item.block_type != "VECTOR"


def _is_full_width_item(
    item: PdfContentBlock,
    *,
    page_width: float,
    config: ReadingOrderConfig,
) -> bool:
    return item.region_id == "full-width" or (
        item.width >= page_width * config.full_width_ratio
    )


def _add_caption_edges(
    items: tuple[PdfContentBlock, ...],
    edges: list[set[int]],
    *,
    config: ReadingOrderConfig,
) -> dict[int, int]:
    parents: dict[int, int] = {}
    targets = [
        index
        for index, item in enumerate(items)
        if item.block_type in {"IMAGE", "TABLE"}
    ]
    for caption_index, caption in enumerate(items):
        if caption.block_type != "CAPTION":
            continue
        candidates: list[tuple[float, int, str]] = []
        text_block = caption.text_block
        if text_block is None:
            continue
        for target_index in targets:
            relation = _caption_distance(
                text_block,
                items[target_index].bbox,
                config=config,
            )
            if relation is None:
                continue
            position, distance = relation
            candidates.append((distance, target_index, position))
        if not candidates:
            continue
        _, target_index, position = min(candidates)
        parents[caption_index] = target_index
        if position == "before":
            _add_order_edge(edges, caption_index, target_index)
        else:
            _add_order_edge(edges, target_index, caption_index)
    return parents


def _fallback_content_order(
    items: tuple[PdfContentBlock, ...],
    *,
    boundaries: tuple[float, ...],
    page_width: float,
    config: ReadingOrderConfig,
) -> list[int]:
    """关系图冲突时，按旧的分栏/坐标规则生成确定性顺序。"""
    def geometry_key(index: int) -> tuple[Any, ...]:
        item = items[index]
        column, region = _region_sort_key(item.region_id)
        return (column, item.top, region, item.x0, index)

    header_indices = sorted(
        (index for index, item in enumerate(items) if item.layer == "header"),
        key=geometry_key,
    )
    footer_indices = sorted(
        (index for index, item in enumerate(items) if item.layer == "footer"),
        key=geometry_key,
    )
    body_indices = [
        index for index, item in enumerate(items) if item.layer == "body"
    ]
    if not boundaries:
        body_order = sorted(body_indices, key=lambda index: (
            items[index].top,
            items[index].x0,
            index,
        ))
        return header_indices + body_order + footer_indices

    anchors = sorted(
        (
            index
            for index in body_indices
            if _participates_in_reading_order(items[index])
            and _is_full_width_item(
                items[index],
                page_width=page_width,
                config=config,
            )
        ),
        key=lambda index: (items[index].top, items[index].x0, index),
    )
    non_anchor_indices = [index for index in body_indices if index not in anchors]
    sections: dict[int, list[int]] = {}
    for index in non_anchor_indices:
        section = sum(
            items[anchor].bottom <= items[index].top + 1.0
            for anchor in anchors
        )
        sections.setdefault(section, []).append(index)

    body_order: list[int] = []
    for section in range(len(anchors) + 1):
        body_order.extend(
            sorted(
                sections.get(section, ()),
                key=geometry_key,
            )
        )
        if section < len(anchors):
            body_order.append(anchors[section])
    return header_indices + body_order + footer_indices


def _sort_content_blocks_in_reading_order(
    items: tuple[PdfContentBlock, ...],
    *,
    boundaries: tuple[float, ...],
    page_width: float,
    config: ReadingOrderConfig = DEFAULT_READING_ORDER_CONFIG,
    reading_order_enabled: bool = True,
    fallback_enabled: bool = True,
) -> tuple[tuple[PdfContentBlock, ...], float, tuple[str, ...]]:
    """用内容块关系图生成页面阅读顺序。"""
    if not items:
        return (), 1.0, ()
    edges: list[set[int]] = [set() for _ in items]
    body_indices = [
        index
        for index, item in enumerate(items)
        if _is_body_item(item) and _participates_in_reading_order(item)
    ]
    header_indices = [
        index
        for index, item in enumerate(items)
        if item.layer == "header" and _participates_in_reading_order(item)
    ]
    footer_indices = [
        index
        for index, item in enumerate(items)
        if item.layer == "footer" and _participates_in_reading_order(item)
    ]
    for header in header_indices:
        for body in body_indices:
            _add_order_edge(edges, header, body)
        for footer in footer_indices:
            _add_order_edge(edges, header, footer)
    for body in body_indices:
        for footer in footer_indices:
            _add_order_edge(edges, body, footer)

    anchors = sorted(
        [
            index
            for index in body_indices
            if _is_full_width_item(
                items[index],
                page_width=page_width,
                config=config,
            )
        ],
        key=lambda index: (items[index].top, items[index].x0, index),
    )
    for left, right in zip(anchors, anchors[1:]):
        _add_order_edge(edges, left, right)
    for anchor in anchors:
        anchor_item = items[anchor]
        for item_index in body_indices:
            if item_index == anchor:
                continue
            item = items[item_index]
            if item.bottom <= anchor_item.top + 1.0:
                _add_order_edge(edges, item_index, anchor)
            elif anchor_item.bottom <= item.top + 1.0:
                _add_order_edge(edges, anchor, item_index)

    sections: dict[int, list[int]] = {}
    for item_index in body_indices:
        if item_index in anchors:
            continue
        section = sum(
            items[anchor_index].bottom <= items[item_index].top + 1.0
            for anchor_index in anchors
        )
        sections.setdefault(section, []).append(item_index)

    for section_indices in sections.values():
        region_groups: dict[str, list[int]] = {}
        for item_index in section_indices:
            region_groups.setdefault(items[item_index].region_id, []).append(
                item_index
            )
        for group in region_groups.values():
            ordered = sorted(
                group,
                key=lambda index: (items[index].top, items[index].x0, index),
            )
            for left, right in zip(ordered, ordered[1:]):
                if items[left].bottom <= items[right].top + 1.0:
                    _add_order_edge(edges, left, right)
                elif items[right].bottom <= items[left].top + 1.0:
                    _add_order_edge(edges, right, left)

        column_groups = {
            region: group
            for region, group in region_groups.items()
            if re.fullmatch(r"column-\d+", region)
        }
        column_names = sorted(column_groups, key=_region_sort_key)
        for left_name, right_name in zip(column_names, column_names[1:]):
            for left in column_groups[left_name]:
                for right in column_groups[right_name]:
                    if (
                        items[left].block_type == "CAPTION"
                        or items[right].block_type == "CAPTION"
                    ):
                        continue
                    _add_order_edge(edges, left, right)

    if not boundaries:
        for left_index, left in enumerate(items):
            if not _is_body_item(left) or not _participates_in_reading_order(left):
                continue
            for right_index in range(left_index + 1, len(items)):
                right = items[right_index]
                if (
                    not _is_body_item(right)
                    or not _participates_in_reading_order(right)
                ):
                    continue
                if abs(left.top - right.top) > max(
                    12.0,
                    min(left.height, right.height)
                    * config.side_by_side_top_delta_factor,
                ):
                    continue
                if _vertical_overlap_ratio(left, right) < config.side_by_side_overlap:
                    continue
                if left.x1 <= right.x0:
                    _add_order_edge(edges, left_index, right_index)
                elif right.x1 <= left.x0:
                    _add_order_edge(edges, right_index, left_index)

    caption_parents = _add_caption_edges(items, edges, config=config)
    for footnote_index, footnote in enumerate(items):
        if footnote.block_type != "FOOTNOTE":
            continue
        for body_index in body_indices:
            if body_index != footnote_index:
                _add_order_edge(edges, body_index, footnote_index)

    if not reading_order_enabled:
        ordered_indices = _fallback_content_order(
            items,
            boundaries=boundaries,
            page_width=page_width,
            config=config,
        )
        warnings = ["内容块阅读顺序已关闭，已使用坐标备用顺序。"]
        confidence = 0.6
    else:
        warnings = []
        confidence = 0.96 if boundaries else 0.86

    indegree = [0] * len(items)
    for source_edges in edges:
        for target in source_edges:
            indegree[target] += 1

    def ready_key(index: int) -> tuple[Any, ...]:
        item = items[index]
        layer = {"header": 0, "body": 1, "footer": 2}.get(item.layer, 1)
        region_index, region_name = _region_sort_key(item.region_id)
        return (layer, item.top, region_index, region_name, item.x0, index)

    if reading_order_enabled:
        ready = [index for index, value in enumerate(indegree) if value == 0]
        ordered_indices = []
        while ready:
            ready.sort(key=ready_key)
            current = ready.pop(0)
            ordered_indices.append(current)
            for target in sorted(edges[current]):
                indegree[target] -= 1
                if indegree[target] == 0:
                    ready.append(target)
        if len(ordered_indices) != len(items):
            ordered_indices = _fallback_content_order(
                items,
                boundaries=boundaries,
                page_width=page_width,
                config=config,
            )
            if fallback_enabled:
                warnings.append("内容块关系存在冲突，已使用坐标备用顺序。")
            else:
                warnings.append("内容块关系存在冲突，已强制使用确定性坐标顺序。")
            confidence = 0.45

    ordered: list[PdfContentBlock] = []
    for reading_order, item_index in enumerate(ordered_indices):
        item = items[item_index]
        parent_index = caption_parents.get(item_index)
        parent_id = (
            items[parent_index].block_id
            if parent_index is not None
            else item.parent_id
        )
        text_block = item.text_block
        if text_block is not None:
            text_block = replace(
                text_block,
                parent_id=parent_id,
                reading_order=reading_order,
                needs_review=item.needs_review or bool(warnings),
            )
        ordered.append(
            replace(
                item,
                text_block=text_block,
                parent_id=parent_id,
                reading_order=reading_order,
                needs_review=item.needs_review or bool(warnings),
            )
        )
    if not boundaries and len(body_indices) > 1:
        confidence = min(confidence, 0.82)
    return tuple(ordered), confidence, tuple(warnings)


def _line_column_bounds(
    line: PdfTextLine,
    columns: tuple[float, ...],
    page_width: float,
) -> tuple[float, float]:
    """返回文本行所在栏的左右边界。"""
    if not columns:
        return 0.0, page_width
    index = bisect_right(columns, line.center_x)
    left = 0.0 if index == 0 else columns[index - 1]
    right = page_width if index == len(columns) else columns[index]
    if right <= left:
        return 0.0, page_width
    return left, right


def _annotate_line_metrics(
    lines: tuple[PdfTextLine, ...],
    *,
    page_width: float,
    columns: tuple[float, ...],
    body_left: float = 0.0,
) -> tuple[PdfTextLine, ...]:
    """补充对齐方式、行距倍数和首行缩进等排版信息。"""
    ordered = sorted(lines, key=lambda line: (line.top, line.x0))
    annotated: list[PdfTextLine] = []
    previous: PdfTextLine | None = None
    tolerance = 6.0
    body_right = _estimate_body_right(ordered)
    for line in ordered:
        if columns:
            left_bound, right_bound = _line_column_bounds(
                line, columns, page_width
            )
        else:
            left_bound = body_left
            right_bound = body_right or page_width
            if right_bound <= left_bound:
                left_bound, right_bound = 0.0, page_width
        align_tolerance = max(4.0, min(10.0, (right_bound - left_bound) * 0.04))
        near_left = line.x0 <= left_bound + align_tolerance
        near_right = line.x1 >= right_bound - align_tolerance
        if near_left and near_right and line.width >= (right_bound - left_bound) * 0.5:
            alignment = "justify"
        elif near_left:
            alignment = "left"
        elif near_right:
            alignment = "right"
        else:
            column_center = (left_bound + right_bound) / 2.0
            centered_short_line = (
                line.width < (right_bound - left_bound) * 0.75
                and abs(line.center_x - column_center)
                <= max(3.0, (right_bound - left_bound) * 0.05)
                and line.font_size >= 11.0
            )
            if centered_short_line:
                alignment = "center"
            else:
                alignment = "left"
        spacing = 0.0
        if (
            previous is not None
            and previous.font_size > 0
            and previous.is_header_footer == line.is_header_footer
        ):
            delta = line.bottom - previous.bottom
            if 0.3 * previous.font_size < delta < 3.5 * previous.font_size:
                spacing = round(delta / previous.font_size, 4)
        first_line_indent = 0.0
        if line.x0 - body_left > tolerance and line.x0 - body_left < 72.0:
            first_line_indent = round(line.x0 - body_left, 3)
        annotated.append(
            replace(
                line,
                alignment=alignment,
                line_spacing=spacing,
                first_line_indent=first_line_indent,
            )
        )
        previous = line
    return tuple(annotated)


def _estimate_body_left(lines: Iterable[PdfTextLine]) -> float:
    """用出现次数最多的行首位置估计正文左边距。"""
    values = [
        round(line.x0, 1)
        for line in lines
        if not line.is_header_footer
        and abs(float(line.rotation or 0.0)) < 1.0
    ]
    if not values:
        return 0.0
    counts = Counter(values)
    most_common = max(counts.values())
    return min(
        value for value, count in counts.items() if count == most_common
    )


def _estimate_body_right(lines: Iterable[PdfTextLine]) -> float:
    """用出现次数最多的行尾位置估计正文右边界。"""
    values = [
        round(line.x1, 1)
        for line in lines
        if not line.is_header_footer
        and abs(float(line.rotation or 0.0)) < 1.0
    ]
    if not values:
        return 0.0
    counts = Counter(values)
    most_common = max(counts.values())
    # 同频时取最靠右的候选，避免短标题把正文区域估窄。
    return max(
        value for value, count in counts.items() if count == most_common
    )


def _mark_logos(
    pages: tuple[PdfPageLayout, ...],
) -> tuple[PdfPageLayout, ...]:
    """跨页重复出现的小图片视为 logo，并标记其页眉/页脚层。"""
    signatures: dict[tuple[int, int], int] = {}
    for page in pages:
        for image in page.images:
            signature = (round(len(image.data) / 64.0), hash(image.data[:96]))
            signatures[signature] = signatures.get(signature, 0) + 1
    updated: list[PdfPageLayout] = []
    for page in pages:
        new_images: list[PdfImageBlock] = []
        for image in page.images:
            signature = (round(len(image.data) / 64.0), hash(image.data[:96]))
            area_ratio = (image.width * image.height) / max(
                page.width * page.height, 1.0
            )
            is_logo = signatures.get(signature, 0) >= 2 and area_ratio <= 0.25
            layer = "body"
            if is_logo:
                if image.bbox[3] <= page.height * 0.22:
                    layer = "header"
                elif image.bbox[1] >= page.height * 0.78:
                    layer = "footer"
            new_images.append(replace(image, is_logo=is_logo, layer=layer))
        new_vectors: list[PdfVectorObject] = []
        for vector in page.vectors:
            layer = "body"
            if vector.bbox[3] <= page.height * 0.06:
                layer = "header"
            elif vector.bbox[1] >= page.height * 0.94:
                layer = "footer"
            new_vectors.append(replace(vector, layer=layer))
        new_lines: list[PdfTextLine] = []
        for line in page.lines:
            layer = "body"
            if line.is_header_footer:
                layer = (
                    "header" if line.center_y <= page.height / 2 else "footer"
                )
            new_lines.append(replace(line, layer=layer))
        updated.append(
            replace(
                page,
                images=tuple(new_images),
                vectors=tuple(new_vectors),
                lines=tuple(new_lines),
            )
        )
    return tuple(updated)


