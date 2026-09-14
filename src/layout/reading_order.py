"""阅读顺序、页眉页脚与 logo 判定。

依据行坐标识别多栏边界并排序，标记跨页重复的页眉页脚与重复出现的小图片。
"""

from __future__ import annotations

from bisect import bisect_right
from collections import Counter
from collections.abc import Iterable
from dataclasses import replace
import math
import re
from typing import Any

from .models import (
    PdfImageBlock,
    PdfPageLayout,
    PdfTable,
    PdfTextBlock,
    PdfTextLine,
    PdfVectorObject,
    _MIN_TABLE_WIDTH,
)
from .tables import _line_in_table


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
) -> tuple[float, ...]:
    """通过跨行稳定空白带检测多栏分栏边界。"""
    candidates = [
        line
        for line in lines
        if not line.is_header_footer
        and not _line_in_table(line, tables)
        and line.x1 - line.x0 < page_width * 0.65
    ]
    if len(candidates) < 6:
        return ()
    intervals = [(line.x0, line.x1) for line in candidates]
    edges = sorted({value for interval in intervals for value in interval})
    gaps: list[tuple[float, float]] = []
    for left, right in zip(edges, edges[1:]):
        # 分栏留白由整栏文字让出，十几点的间距已是明显分栏；
        # 更窄的间距属于行内字距，不作为分栏依据。
        if right - left < 12.0:
            continue
        midpoint = (left + right) / 2.0
        if not (page_width * 0.2 <= midpoint <= page_width * 0.8):
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
        if not boundaries or boundary - boundaries[-1] >= 24.0:
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
        if line.x1 - line.x0 >= page_width * 0.65
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
) -> int:
    if not boundaries:
        return 0
    if line.width >= page_width * 0.65 or any(
        line.x0 < boundary < line.x1 for boundary in boundaries
    ):
        return -1
    return bisect_right(boundaries, line.center_x)


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
) -> bool:
    if previous_column != current_column:
        return False
    if previous.is_header_footer != current.is_header_footer:
        return False
    if not _same_text_direction(previous, current):
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
    return gap <= max(8.0, scale * 1.75)


def _build_text_blocks(
    lines: tuple[PdfTextLine, ...],
    *,
    boundaries: tuple[float, ...],
    page_width: float,
) -> tuple[PdfTextBlock, ...]:
    """按方向、栏位和行间距把文本行聚合成几何版面块。"""
    blocks: list[PdfTextBlock] = []
    current: list[PdfTextLine] = []
    current_column = 0

    def flush() -> None:
        nonlocal current
        if current:
            blocks.append(_text_block_from_lines(current, current_column))
            current = []

    for line in lines:
        column = _line_column_index(line, boundaries, page_width)
        if not current:
            current = [line]
            current_column = column
            continue
        if not _can_extend_text_block(
            current[-1], line, current_column, column
        ):
            flush()
            current_column = column
        current.append(line)
    flush()
    return tuple(blocks)


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
    for line in ordered:
        left_bound, right_bound = _line_column_bounds(
            line, columns, page_width
        )
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
            if abs(line.center_x - column_center) <= max(3.0, (right_bound - left_bound) * 0.05):
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
        round(line.x0, 1) for line in lines if not line.is_header_footer
    ]
    if not values:
        return 0.0
    counts = Counter(values)
    most_common = max(counts.values())
    return min(
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


