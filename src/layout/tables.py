"""表格识别：矢量线表格、无边框表格与跨页续表。

有线表格由横竖线聚类出网格；无边框表格按列锚点与行分组识别；
最后统一处理跨页续表与表头重复。
"""

from __future__ import annotations
from typing import Any
from .models import (
    PdfTable,
    PdfTableCell,
    PdfTextLine,
    _MAX_TABLE_ROW_GAP,
    _MIN_BOUNDARY_COVERAGE,
    _MIN_CELL_LINE_LENGTH,
    _MIN_TABLE_WIDTH,
    _TableCandidate,
    _VerticalLine,
)
from .vectors import (
    _cluster_horizontal_lines,
    _cluster_vertical_lines,
    _extract_vector_lines,
    _has_horizontal_boundary,
    _has_vertical_boundary,
    _horizontal_coverage,
    _vertical_coverage,
)
import re
from collections import Counter
from dataclasses import replace
from .models import (
    PdfPageLayout,
    _COORDINATE_TOLERANCE,
    _HorizontalLine,
    PdfTextGlyph,
    PdfTextSpan,
)

def _table_boundary_lines(
    vertical: list[_VerticalLine],
    *,
    x0: float,
    x1: float,
    top: float,
    bottom: float,
) -> list[_VerticalLine]:
    return [
        line
        for line in vertical
        if x0 - _COORDINATE_TOLERANCE <= line.x <= x1 + _COORDINATE_TOLERANCE
        and _vertical_coverage([line], top=top, bottom=bottom) >= 0.9
    ]


def _table_column_boundaries(
    vertical: list[_VerticalLine],
    *,
    x0: float,
    x1: float,
    top: float,
    bottom: float,
) -> tuple[float, ...]:
    """从候选区域内的垂直线段提取列边界，保留局部边界以支持合并单元格。"""
    positions = [x0, x1]
    for line in vertical:
        if line.x < x0 - _COORDINATE_TOLERANCE or line.x > x1 + _COORDINATE_TOLERANCE:
            continue
        overlap = min(line.bottom, bottom) - max(line.top, top)
        if overlap >= _MIN_CELL_LINE_LENGTH:
            positions.append(line.x)

    clusters: list[list[float]] = []
    for position in sorted(positions):
        if (
            clusters
            and position - clusters[-1][-1] <= _COORDINATE_TOLERANCE
        ):
            clusters[-1].append(position)
        else:
            clusters.append([position])
    return tuple(sum(cluster) / len(cluster) for cluster in clusters)


def _extract_cell_spans(candidate: _TableCandidate) -> list[tuple[int, int, int, int]]:
    """根据局部边界缺失情况推断表格中的跨行和跨列单元格。"""
    column_boundaries = candidate.column_boundaries
    row_boundaries = candidate.row_boundaries
    row_count = max(len(row_boundaries) - 1, 0)
    column_count = max(len(column_boundaries) - 1, 0)
    occupied: set[tuple[int, int]] = set()
    spans: list[tuple[int, int, int, int]] = []

    for row_index in range(row_count):
        for column_index in range(column_count):
            if (row_index, column_index) in occupied:
                continue
            column_span = 1
            while column_index + column_span < column_count:
                divider_x = column_boundaries[column_index + column_span]
                if _has_vertical_boundary(
                    list(candidate.vertical_lines),
                    x=divider_x,
                    top=row_boundaries[row_index],
                    bottom=row_boundaries[row_index + 1],
                ):
                    break
                column_span += 1

            row_span = 1
            while row_index + row_span < row_count:
                next_row = row_index + row_span
                if any(
                    (next_row, column) in occupied
                    for column in range(column_index, column_index + column_span)
                ):
                    break
                if _has_horizontal_boundary(
                    list(candidate.horizontal_lines),
                    top=row_boundaries[next_row],
                    x0=column_boundaries[column_index],
                    x1=column_boundaries[column_index + column_span],
                ):
                    break
                row_span += 1

            span = (row_index, column_index, row_span, column_span)
            spans.append(span)
            for row in range(row_index, row_index + row_span):
                for column in range(column_index, column_index + column_span):
                    occupied.add((row, column))
    return spans


def _horizontal_run_tables(
    horizontal: list[_HorizontalLine],
    vertical: list[_VerticalLine],
) -> list[_TableCandidate]:
    candidates: list[_TableCandidate] = []
    index = 0
    while index < len(horizontal):
        run = [horizontal[index]]
        index += 1
        while index < len(horizontal):
            previous = run[-1]
            current = horizontal[index]
            shared_width = min(previous.x1, current.x1) - max(previous.x0, current.x0)
            if (
                current.top - previous.top > _MAX_TABLE_ROW_GAP
                or shared_width < _MIN_TABLE_WIDTH
            ):
                break
            prospective_run = [*run, current]
            prospective_x0 = min(item.x0 for item in prospective_run)
            prospective_x1 = max(item.x1 for item in prospective_run)
            if len(run) >= 1 and len(
                _table_boundary_lines(
                    vertical,
                    x0=prospective_x0,
                    x1=prospective_x1,
                    top=prospective_run[0].top,
                    bottom=prospective_run[-1].top,
                )
            ) < 2:
                break
            run.append(current)
            index += 1
        if len(run) < 3:
            continue

        table_top = run[0].top
        table_bottom = run[-1].top
        x0 = min(item.x0 for item in run)
        x1 = max(item.x1 for item in run)
        if x1 - x0 < _MIN_TABLE_WIDTH:
            continue

        column_boundaries = _table_column_boundaries(
            vertical,
            x0=x0,
            x1=x1,
            top=table_top,
            bottom=table_bottom,
        )
        if len(column_boundaries) < 2:
            continue
        x0 = column_boundaries[0]
        x1 = column_boundaries[-1]
        row_boundaries = tuple(item.top for item in run)
        candidates.append(
            _TableCandidate(
                bbox=(x0, table_top, x1, table_bottom),
                column_boundaries=column_boundaries,
                row_boundaries=row_boundaries,
                horizontal_lines=tuple(run),
                vertical_lines=tuple(vertical),
            )
        )
    return candidates


def _line_is_in_table(line: PdfTextLine, bbox: tuple[float, float, float, float]) -> bool:
    x0, top, x1, bottom = bbox
    return (
        x0 - 1.0 <= line.center_x <= x1 + 1.0
        and top - 1.0 <= line.center_y <= bottom + 1.0
    )


def _line_in_table(line: PdfTextLine, tables: tuple[PdfTable, ...]) -> bool:
    """判断文本行是否落在某个表格区域内。"""
    for table in tables:
        x0, top, x1, bottom = table.bbox
        if (
            x0 - 1.0 <= line.center_x <= x1 + 1.0
            and top - 1.0 <= line.center_y <= bottom + 1.0
        ):
            return True
    return False


def _line_in_tables(line: PdfTextLine, tables: tuple[PdfTable, ...]) -> bool:
    return any(_line_in_table(line, (table,)) for table in tables)


def _cell_spans_from_glyphs(
    glyphs: list[PdfTextGlyph],
) -> tuple[PdfTextSpan, ...]:
    """按字符样式重建单元格内的连续 span。"""
    if not glyphs:
        return ()
    spans: list[PdfTextSpan] = []
    current: list[PdfTextGlyph] = []
    for glyph in glyphs:
        if current and (
            glyph.font_name != current[0].font_name
            or glyph.pdf_font_name != current[0].pdf_font_name
            or glyph.bold != current[0].bold
            or glyph.italic != current[0].italic
            or glyph.color != current[0].color
        ):
            spans.append(_make_cell_span(current))
            current = []
        current.append(glyph)
    if current:
        spans.append(_make_cell_span(current))
    return tuple(spans)


def _glyph_row_groups(
    glyphs: tuple[PdfTextGlyph, ...],
    *,
    tolerance: float = 1.5,
) -> list[list[PdfTextGlyph]]:
    """按字形的垂直基线分组，空格归入最近的可见文字行。"""
    visible = [glyph for glyph in glyphs if glyph.bbox[3] > glyph.bbox[1]]
    if not visible:
        return []
    groups: list[list[PdfTextGlyph]] = []
    centers: list[float] = []
    for glyph in sorted(visible, key=lambda item: (item.bbox[1], item.bbox[0])):
        baseline = glyph.bbox[3]
        if centers and abs(baseline - centers[-1]) <= tolerance:
            groups[-1].append(glyph)
            centers[-1] = sum(
                item.bbox[3] for item in groups[-1]
            ) / len(groups[-1])
        else:
            groups.append([glyph])
            centers.append(baseline)
    result: list[list[PdfTextGlyph]] = []
    for glyph in glyphs:
        if glyph in visible:
            continue
        center = glyph.bbox[3]
        nearest = min(
            range(len(centers)),
            key=lambda index: abs(center - centers[index]),
        )
        groups[nearest].append(glyph)
    for group in groups:
        group.sort(key=lambda item: item.char_index if item.char_index >= 0 else 10**9)
    return groups


def _split_cell_by_glyph_rows(
    cell: PdfTableCell,
) -> list[tuple[list[PdfTextGlyph], tuple[float, float]]]:
    groups = _glyph_row_groups(cell.glyphs)
    if len(groups) < 2:
        return []
    return [
        (
            group,
            (
                min(glyph.bbox[1] for glyph in group),
                max(glyph.bbox[3] for glyph in group),
            ),
        )
        for group in groups
    ]


def _expand_uniform_internal_rows(table: PdfTable) -> PdfTable:
    """把所有单元格共有的内部文字基线展开为逻辑行。"""
    if not table.cells or table.row_count < 2:
        return table
    physical_rows: list[list[PdfTableCell]] = []
    for row_index in range(table.row_count):
        physical_rows.append(
            [cell for cell in table.cells if cell.row_index == row_index]
        )
    expansion: dict[int, list[list[tuple[list[PdfTextGlyph], tuple[float, float]]]]] = {}
    for row_index, cells in enumerate(physical_rows):
        if not cells or any(cell.row_span != 1 for cell in cells):
            continue
        split_cells = [_split_cell_by_glyph_rows(cell) for cell in cells]
        counts = Counter(len(groups) for groups in split_cells if groups)
        if not counts:
            continue
        group_count, support = counts.most_common(1)[0]
        if group_count < 2 or support < max(3, len(cells) // 2):
            continue
        if any(groups and len(groups) != group_count for groups in split_cells):
            continue
        expansion[row_index] = [
            groups if groups else [([], (0.0, 0.0))] * group_count
            for groups in split_cells
        ]
    if not expansion:
        return table

    new_boundaries: list[float] = [table.row_boundaries[0]]
    visible_boundaries: list[bool] = [True]
    row_offsets: dict[int, int] = {}
    logical_row_count = 0
    for row_index in range(table.row_count):
        row_offsets[row_index] = logical_row_count
        groups_for_row = expansion.get(row_index)
        if not groups_for_row:
            new_boundaries.append(table.row_boundaries[row_index + 1])
            visible_boundaries.append(True)
            logical_row_count += 1
            continue
        group_count = len(groups_for_row[0])
        row_top = table.row_boundaries[row_index]
        row_bottom = table.row_boundaries[row_index + 1]
        group_bounds = [
            (
                max(
                    (groups_for_row[cell_index][group_index][1][0]
                     for cell_index in range(len(groups_for_row))
                     if groups_for_row[cell_index][group_index][1] != (0.0, 0.0)),
                    default=row_top,
                ),
                min(
                    (groups_for_row[cell_index][group_index][1][1]
                     for cell_index in range(len(groups_for_row))
                     if groups_for_row[cell_index][group_index][1] != (0.0, 0.0)),
                    default=row_bottom,
                ),
            )
            for group_index in range(group_count)
        ]
        for group_index in range(group_count - 1):
            boundary = (group_bounds[group_index][1] + group_bounds[group_index + 1][0]) / 2.0
            new_boundaries.append(max(row_top, min(row_bottom, boundary)))
            visible_boundaries.append(False)
        new_boundaries.append(row_bottom)
        visible_boundaries.append(True)
        logical_row_count += group_count

    new_cells: list[PdfTableCell] = []
    new_rows = [
        ["" for _ in range(table.column_count)]
        for _ in range(logical_row_count)
    ]
    for row_index, cells in enumerate(physical_rows):
        groups_for_row = expansion.get(row_index)
        offset = row_offsets[row_index]
        if not groups_for_row:
            for cell in cells:
                new_cells.append(cell if cell.row_index == offset else replace(cell, row_index=offset))
                if cell.column_index < table.column_count:
                    new_rows[offset][cell.column_index] = cell.text
            continue
        group_count = len(groups_for_row[0])
        for cell_index, cell in enumerate(cells):
            for group_index in range(group_count):
                glyphs = groups_for_row[cell_index][group_index][0]
                text = "".join(glyph.text for glyph in glyphs).strip()
                if glyphs:
                    text_bbox = (
                        min(glyph.bbox[0] for glyph in glyphs),
                        min(glyph.bbox[1] for glyph in glyphs),
                        max(glyph.bbox[2] for glyph in glyphs),
                        max(glyph.bbox[3] for glyph in glyphs),
                    )
                    font_sizes = [glyph.font_size for glyph in glyphs if glyph.font_size > 0]
                    spans = _cell_spans_from_glyphs(glyphs)
                    source_char_ids = tuple(
                        dict.fromkeys(glyph.source_char_id for glyph in glyphs if glyph.source_char_id)
                    )
                else:
                    text_bbox = (0.0, 0.0, 0.0, 0.0)
                    font_sizes = []
                    spans = ()
                    source_char_ids = ()
                logical_index = offset + group_index
                new_cells.append(
                    replace(
                        cell,
                        row_index=logical_index,
                        bbox=(
                            cell.bbox[0],
                            new_boundaries[logical_index],
                            cell.bbox[2],
                            new_boundaries[logical_index + 1],
                        ),
                        text=text,
                        text_bbox=text_bbox,
                        font_size=(font_sizes[len(font_sizes) // 2] if font_sizes else 0.0),
                        spans=spans,
                        glyphs=tuple(glyphs),
                        source_char_ids=source_char_ids,
                    )
                )
                if cell.column_index < table.column_count:
                    new_rows[logical_index][cell.column_index] = text
    return replace(
        table,
        row_boundaries=tuple(new_boundaries),
        rows=tuple(tuple(row) for row in new_rows),
        cells=tuple(new_cells),
        visible_row_boundaries=tuple(visible_boundaries),
    )


def _make_cell_span(glyphs: list[PdfTextGlyph]) -> PdfTextSpan:
    representative = glyphs[0]
    return PdfTextSpan(
        text="".join(glyph.text for glyph in glyphs),
        bbox=(
            min(glyph.bbox[0] for glyph in glyphs),
            min(glyph.bbox[1] for glyph in glyphs),
            max(glyph.bbox[2] for glyph in glyphs),
            max(glyph.bbox[3] for glyph in glyphs),
        ),
        font_name=representative.font_name,
        pdf_font_name=representative.pdf_font_name,
        font_size=representative.font_size,
        color=representative.color,
        bold=representative.bold,
        italic=representative.italic,
        rotation=representative.rotation,
        z_order=min(glyph.z_order for glyph in glyphs),
        direction=representative.direction,
        glyphs=tuple(glyphs),
        source_char_ids=tuple(
            glyph.source_char_id for glyph in glyphs if glyph.source_char_id
        ),
        nominal_font_size=representative.nominal_font_size,
        font_size_source=representative.font_size_source,
    )


def _extract_table(
    lines: list[PdfTextLine],
    candidate: _TableCandidate,
) -> PdfTable:
    x0, top, x1, bottom = candidate.bbox
    column_boundaries = candidate.column_boundaries
    row_boundaries = candidate.row_boundaries
    cell_spans = _extract_cell_spans(candidate)
    cell_lines: dict[tuple[int, int], list[PdfTextLine]] = {
        (row_index, column_index): []
        for row_index, column_index, _, _ in cell_spans
    }
    cell_glyphs: dict[tuple[int, int], list[PdfTextGlyph]] = {
        (row_index, column_index): []
        for row_index, column_index, _, _ in cell_spans
    }
    for line in lines:
        if not _line_is_in_table(line, candidate.bbox):
            continue
        matching_cells: list[tuple[float, tuple[int, int]]] = []
        for row_index, column_index, row_span, column_span in cell_spans:
            cell_top = row_boundaries[row_index]
            cell_bottom = row_boundaries[row_index + row_span]
            cell_x0 = column_boundaries[column_index]
            cell_x1 = column_boundaries[column_index + column_span]
            overlap_x = min(line.x1, cell_x1) - max(line.x0, cell_x0)
            if overlap_x <= 0 and not cell_x0 <= line.center_x <= cell_x1:
                continue
            if not cell_top <= line.center_y <= cell_bottom:
                continue
            overlap_y = min(line.bottom, cell_bottom) - max(line.top, cell_top)
            score = max(overlap_x, 0.1) * max(overlap_y, 0.1)
            matching_cells.append((score, (row_index, column_index)))
        if not matching_cells:
            continue
        _, cell_key = max(matching_cells, key=lambda item: item[0])
        if not line.glyphs:
            cell_lines[cell_key].append(line)
            continue
        for glyph in line.glyphs:
            glyph_matches: list[tuple[float, tuple[int, int]]] = []
            glyph_center_x = (glyph.bbox[0] + glyph.bbox[2]) / 2.0
            glyph_center_y = (glyph.bbox[1] + glyph.bbox[3]) / 2.0
            for row_index, column_index, row_span, column_span in cell_spans:
                cell_top = row_boundaries[row_index]
                cell_bottom = row_boundaries[row_index + row_span]
                cell_x0 = column_boundaries[column_index]
                cell_x1 = column_boundaries[column_index + column_span]
                overlap_x = min(glyph.bbox[2], cell_x1) - max(
                    glyph.bbox[0], cell_x0
                )
                overlap_y = min(glyph.bbox[3], cell_bottom) - max(
                    glyph.bbox[1], cell_top
                )
                center_inside = (
                    cell_x0 <= glyph_center_x <= cell_x1
                    and cell_top <= glyph_center_y <= cell_bottom
                )
                if overlap_x <= 0 and not center_inside:
                    continue
                if overlap_y <= 0 and not center_inside:
                    continue
                score = max(overlap_x, 0.1) * max(overlap_y, 0.1)
                if center_inside:
                    score += 1_000.0
                glyph_matches.append((score, (row_index, column_index)))
            if glyph_matches:
                _, glyph_cell_key = max(
                    glyph_matches,
                    key=lambda item: item[0],
                )
                cell_glyphs[glyph_cell_key].append(glyph)

    rows = [
        ["" for _ in range(max(len(column_boundaries) - 1, 0))]
        for _ in range(max(len(row_boundaries) - 1, 0))
    ]
    cells: list[PdfTableCell] = []
    for row_index, column_index, row_span, column_span in cell_spans:
        values = sorted(
            cell_lines[(row_index, column_index)],
            key=lambda line: (line.top, line.x0),
        )
        glyphs = sorted(
            cell_glyphs[(row_index, column_index)],
            key=lambda glyph: (
                glyph.char_index if glyph.char_index >= 0 else 10**9,
                glyph.bbox[1],
                glyph.bbox[0],
            ),
        )
        if glyphs:
            text = "".join(glyph.text for glyph in glyphs).strip()
        else:
            text = "\n".join(line.text for line in values).strip()
        rows[row_index][column_index] = text
        text_bbox = (
            (
                min(glyph.bbox[0] for glyph in glyphs),
                min(glyph.bbox[1] for glyph in glyphs),
                max(glyph.bbox[2] for glyph in glyphs),
                max(glyph.bbox[3] for glyph in glyphs),
            )
            if glyphs
            else (
                (
                    min(line.x0 for line in values),
                    min(line.top for line in values),
                    max(line.x1 for line in values),
                    max(line.bottom for line in values),
                )
                if values
                else (0.0, 0.0, 0.0, 0.0)
            )
        )
        style_items = glyphs or [
            PdfTextGlyph(
                text=line.text,
                bbox=(line.x0, line.top, line.x1, line.bottom),
                font_name=line.font_name,
                pdf_font_name=line.pdf_font_name,
                font_size=line.font_size,
                color=line.color,
                bold=line.bold,
                italic=line.italic,
                rotation=line.rotation,
                direction=line.direction,
                z_order=line.z_order,
                source_char_id=line.source_char_ids[0]
                if line.source_char_ids
                else "",
            )
            for line in values
        ]
        font_sizes = sorted(
            glyph.font_size for glyph in style_items if glyph.font_size > 0
        )
        source_char_ids = tuple(
            dict.fromkeys(
                glyph.source_char_id
                for glyph in style_items
                if glyph.source_char_id
            )
        )
        spans = _cell_spans_from_glyphs(style_items)
        cells.append(
            PdfTableCell(
                row_index=row_index,
                column_index=column_index,
                row_span=row_span,
                column_span=column_span,
                bbox=(
                    column_boundaries[column_index],
                    row_boundaries[row_index],
                    column_boundaries[column_index + column_span],
                    row_boundaries[row_index + row_span],
                ),
                text=text,
                text_bbox=text_bbox,
                font_size=(
                    font_sizes[len(font_sizes) // 2] if font_sizes else 0.0
                ),
                bold=sum(glyph.bold for glyph in style_items) * 2
                >= max(len(style_items), 1),
                alignment=(
                    Counter(line.alignment for line in values).most_common(1)[0][0]
                    if values
                    else "left"
                ),
                spans=spans,
                glyphs=tuple(glyphs),
                source_char_ids=source_char_ids,
            )
        )

    return _expand_uniform_internal_rows(PdfTable(
        bbox=candidate.bbox,
        column_boundaries=column_boundaries,
        row_boundaries=row_boundaries,
        rows=tuple(tuple(row) for row in rows),
        cells=tuple(cells),
        source_char_ids=tuple(
            dict.fromkeys(
                char_id
                for cell in cells
                for char_id in cell.source_char_ids
            )
        ),
    ))


def _find_tables(
    page: Any,
    lines: list[PdfTextLine],
    page_height: float,
    raster_images: tuple[Any, ...] = (),
) -> tuple[PdfTable, ...]:
    horizontal, vertical = _extract_vector_lines(
        page,
        page_height,
        raster_images=raster_images,
    )
    candidates = _horizontal_run_tables(horizontal, vertical)
    tables = [
        _extract_table(lines, candidate)
        for candidate in candidates
    ]
    reliable: list[PdfTable] = []
    for table in tables:
        values = [value.strip() for row in table.rows for value in row if value.strip()]
        if not values:
            continue
        total_cells = max(table.row_count * table.column_count, 1)
        density = len(values) / total_cells
        if table.column_count >= 4 and density < 0.18:
            continue
        if (
            table.column_count >= 3
            and table.row_count >= 3
            and len(values) < table.row_count
        ):
            continue
        reliable.append(table)
    return tuple(reliable)


def _column_width_ratios(table: PdfTable) -> tuple[float, ...]:
    widths = [
        table.column_boundaries[index + 1] - table.column_boundaries[index]
        for index in range(table.column_count)
    ]
    total = sum(widths)
    if total <= 0:
        return ()
    return tuple(width / total for width in widths)


def _tables_can_continue(
    previous_page: PdfPageLayout,
    previous_table: PdfTable,
    current_page: PdfPageLayout,
    current_table: PdfTable,
) -> bool:
    """判断相邻页面的表格是否具有续表的几何特征。"""
    if previous_table.column_count != current_table.column_count:
        return False
    if not previous_table.rows or not current_table.rows:
        return False
    if previous_table.bbox[3] < previous_page.height * 0.7:
        return False
    if current_table.bbox[1] > current_page.height * 0.3:
        return False
    previous_x0, _, previous_x1, _ = previous_table.bbox
    current_x0, _, current_x1, _ = current_table.bbox
    previous_width = max(previous_x1 - previous_x0, 1.0)
    current_width = max(current_x1 - current_x0, 1.0)
    if abs(previous_x0 / previous_page.width - current_x0 / current_page.width) > 0.04:
        return False
    if abs(previous_x1 / previous_page.width - current_x1 / current_page.width) > 0.04:
        return False
    if abs(previous_width / previous_page.width - current_width / current_page.width) > 0.04:
        return False
    previous_ratios = _column_width_ratios(previous_table)
    current_ratios = _column_width_ratios(current_table)
    return bool(
        previous_ratios
        and len(previous_ratios) == len(current_ratios)
        and max(
            abs(previous - current)
            for previous, current in zip(previous_ratios, current_ratios)
        )
        <= 0.06
    )


def _mark_table_continuations(
    pages: tuple[PdfPageLayout, ...],
) -> tuple[PdfPageLayout, ...]:
    """为相邻页面的续表补充表头元数据。"""
    updated_pages = list(pages)
    for page_index in range(1, len(updated_pages)):
        previous_page = updated_pages[page_index - 1]
        current_page = updated_pages[page_index]
        if not previous_page.tables or not current_page.tables:
            continue
        previous_table = previous_page.tables[-1]
        current_table = current_page.tables[0]
        if not _tables_can_continue(
            previous_page,
            previous_table,
            current_page,
            current_table,
        ):
            continue
        header = (previous_table.rows[0],)
        current_header_is_present = current_table.rows[0] == previous_table.rows[0]
        updated_table = replace(
            current_table,
            continued_from_previous_page=True,
            header_row_count=1 if current_header_is_present else 0,
            continuation_header_rows=() if current_header_is_present else header,
        )
        updated_pages[page_index] = replace(
            current_page,
            tables=(updated_table, *current_page.tables[1:]),
        )
    return tuple(updated_pages)


def _cluster_positions(values: list[float], tolerance: float = 14.0) -> list[float]:
    clusters: list[list[float]] = []
    for value in sorted(values):
        if clusters and value - clusters[-1][-1] <= tolerance:
            clusters[-1].append(value)
        else:
            clusters.append([value])
    return [sum(cluster) / len(cluster) for cluster in clusters]


def _group_text_rows(
    lines: list[PdfTextLine],
    *,
    tolerance: float = 4.0,
) -> list[tuple[PdfTextLine, ...]]:
    rows: list[tuple[PdfTextLine, ...]] = []
    current: list[PdfTextLine] = []
    current_top: float | None = None
    for line in sorted(lines, key=lambda item: (item.top, item.x0)):
        if current and current_top is not None and abs(line.top - current_top) <= tolerance:
            current.append(line)
        else:
            if current:
                rows.append(tuple(sorted(current, key=lambda item: item.x0)))
            current = [line]
            current_top = line.top
    if current:
        rows.append(tuple(sorted(current, key=lambda item: item.x0)))
    return rows


def _is_numeric_table_cell(value: str) -> bool:
    text = value.strip()
    if not text or len(text) > 20:
        return False
    return bool(re.fullmatch(r"[\d.,%+\-/() ]+", text))


_TABLE_CAPTION = re.compile(r"^\s*(?:表|Table)\s*\d+")


def _match_borderless_row(
    row: tuple[PdfTextLine, ...],
    anchors: list[float],
) -> list[tuple[int, int, PdfTextLine]]:
    matched: list[tuple[int, int, PdfTextLine]] = []
    for line in row:
        start_index = min(
            range(len(anchors)),
            key=lambda index: abs(anchors[index] - line.x0),
        )
        if abs(anchors[start_index] - line.x0) > 14.0:
            continue
        end_index = start_index
        for index in range(start_index + 1, len(anchors)):
            if anchors[index] <= line.x1 + 8.0:
                end_index = index
            else:
                break
        matched.append((start_index, end_index, line))
    return matched


def _looks_like_code_table_line(value: str) -> bool:
    """无边框表格检测时排除代码行，避免把代码清单展开成窄列表格。"""
    text = value.strip()
    if not text:
        return False
    if text.startswith(("#", "//", "/*", "*/", "*")):
        return True
    if any(character in text for character in "{};"):
        return True
    if re.search(r"[A-Za-z_]\w*\s*->\s*[A-Za-z_]\w*", text):
        return True
    punctuation = sum(text.count(character) for character in "{};()=<>[]")
    return punctuation >= 3 and len(text) <= 80


def _find_borderless_tables(
    lines: list[PdfTextLine],
    *,
    page_width: float,
    page_height: float,
    existing_tables: tuple[PdfTable, ...],
    confirmed_regions: tuple[Any, ...] = (),
) -> tuple[PdfTable, ...]:
    """通过重复列锚点和多行结构一致性识别无边框表格。"""
    table_regions = tuple(
        region
        for region in confirmed_regions
        if str(getattr(region, "kind", "")).lower() in {"table", "table_region"}
    )
    if confirmed_regions and not table_regions:
        return ()
    candidates = [
        line
        for line in lines
        if not line.is_header_footer
        and not _line_in_tables(line, existing_tables)
        and line.x1 - line.x0 < page_width * 0.9
        and not _looks_like_code_table_line(line.text)
        and (
            not table_regions
            or any(
                _line_is_in_table(line, region.bbox)
                for region in table_regions
            )
        )
    ]
    rows = _group_text_rows(candidates)
    anchor_rows = [row for row in rows if len(row) >= 2]
    if len(anchor_rows) < 3:
        return ()

    anchor_positions = [
        round(line.x0 / 4.0) * 4.0
        for row in anchor_rows
        for line in row
    ]
    counts = Counter(anchor_positions)
    repeated = sorted(
        value for value, count in counts.items() if count >= 3
    )
    anchors = _cluster_positions(repeated)
    if len(anchors) < 2:
        return ()

    scored_rows: list[
        tuple[tuple[PdfTextLine, ...], list[tuple[int, int, PdfTextLine]]]
    ] = []
    for row in rows:
        matched = _match_borderless_row(row, anchors)
        spans_multiple_columns = any(
            end_index > start_index
            for start_index, end_index, _ in matched
        )
        if len(matched) >= 2 or spans_multiple_columns:
            scored_rows.append((row, matched))
    if len(scored_rows) < 3:
        return ()

    runs: list[
        list[tuple[tuple[PdfTextLine, ...], list[tuple[int, int, PdfTextLine]]]]
    ] = []
    current_run: list[
        tuple[tuple[PdfTextLine, ...], list[tuple[int, int, PdfTextLine]]]
    ] = []
    previous_top: float | None = None
    for item in scored_rows:
        top = min(line.top for line in item[0])
        if (
            current_run
            and previous_top is not None
            and top - previous_top > 60.0
        ):
            runs.append(current_run)
            current_run = []
        current_run.append(item)
        previous_top = top
    if current_run:
        runs.append(current_run)

    tables: list[PdfTable] = []
    for run in runs:
        if len(run) < 3:
            continue
        run_anchors = _cluster_positions(
            [
                round(anchor / 4.0) * 4.0
                for _, matched in run
                for start_index, end_index, _ in matched
                for anchor in anchors[start_index : end_index + 1]
            ]
        )
        column_count = len(run_anchors)
        if column_count < 2:
            continue
        first_row_top = min(
            line.top for _, matched in run for _, _, line in matched
        )
        has_caption = any(
            _TABLE_CAPTION.match(line.text)
            and line.bottom <= first_row_top + 4.0
            and first_row_top - line.bottom <= 50.0
            for line in candidates
        )
        numeric_rows = sum(
            any(_is_numeric_table_cell(line.text) for _, _, line in matched)
            for _, matched in run
        )
        if column_count > 8 and not has_caption:
            continue
        if len(run) > 40 and not has_caption:
            continue
        if not has_caption and numeric_rows < max(2, len(run) // 2):
            continue
        if column_count == 2:
            if len(run) < 4:
                continue
            if not has_caption and numeric_rows < max(2, len(run) // 2):
                continue

        run_rows: list[
            tuple[tuple[PdfTextLine, ...], list[tuple[int, int, PdfTextLine]]]
        ] = []
        for row, _ in run:
            run_rows.append((row, _match_borderless_row(row, run_anchors)))

        column_x0: list[float] = []
        column_x1: list[float] = []
        for column_index in range(column_count):
            values = [
                line
                for _, matched in run_rows
                for start_index, end_index, line in matched
                if start_index == end_index == column_index
            ]
            if not values:
                column_x0.append(run_anchors[column_index])
                column_x1.append(run_anchors[column_index] + 20.0)
                continue
            column_x0.append(min(line.x0 for line in values))
            column_x1.append(max(line.x1 for line in values))
        if column_x1[-1] - column_x0[0] < page_width * 0.35:
            continue
        if min(
            column_x1[index] - column_x0[index]
            for index in range(column_count)
        ) < 18.0:
            continue
        matched_texts = [
            line.text
            for _, matched in run
            for _, _, line in matched
        ]
        average_text_length = sum(len(text) for text in matched_texts) / max(
            len(matched_texts), 1
        )
        if average_text_length > 50.0:
            continue
        boundaries = [column_x0[0] - 8.0]
        for index in range(column_count - 1):
            boundaries.append((column_x1[index] + column_x0[index + 1]) / 2.0)
        boundaries.append(column_x1[-1] + 8.0)
        if any(
            boundaries[index] >= boundaries[index + 1]
            for index in range(len(boundaries) - 1)
        ):
            continue

        row_boundaries: list[float] = []
        table_rows: list[tuple[str, ...]] = []
        cells: list[PdfTableCell] = []
        for row_index, (row, matched) in enumerate(run_rows):
            row_top = min(line.top for line in row) - 2.0
            row_bottom = max(line.bottom for line in row) + 2.0
            row_boundaries.append(row_top)
            row_values = [""] * column_count
            cell_map: dict[int, tuple[int, list[str]]] = {}
            for start_index, end_index, line in matched:
                if start_index in cell_map:
                    old_end, texts = cell_map[start_index]
                    cell_map[start_index] = (
                        max(old_end, end_index),
                        [*texts, line.text],
                    )
                else:
                    cell_map[start_index] = (end_index, [line.text])
            covered_columns = {
                column
                for start_index, (end_index, _) in cell_map.items()
                for column in range(start_index, end_index + 1)
            }
            for column_index in range(column_count):
                if column_index in cell_map:
                    end_index, texts = cell_map[column_index]
                    row_values[column_index] = " ".join(texts)
                elif column_index in covered_columns:
                    continue
                else:
                    end_index = column_index
                cells.append(
                    PdfTableCell(
                        row_index=row_index,
                        column_index=column_index,
                        row_span=1,
                        column_span=max(1, end_index - column_index + 1),
                        bbox=(
                            boundaries[column_index],
                            row_top,
                            boundaries[min(end_index + 1, column_count)],
                            row_bottom,
                        ),
                        text=row_values[column_index],
                    )
                )
            table_rows.append(tuple(row_values))
        if len(row_boundaries) < 3:
            continue
        row_boundaries.append(
            max(line.bottom for _, matched in run_rows for _, _, line in matched) + 2.0
        )
        bbox = (
            min(line.x0 for _, matched in run_rows for _, _, line in matched),
            min(line.top for _, matched in run_rows for _, _, line in matched),
            max(line.x1 for _, matched in run_rows for _, _, line in matched),
            max(line.bottom for _, matched in run_rows for _, _, line in matched),
        )
        tables.append(
            PdfTable(
                bbox=bbox,
                column_boundaries=tuple(boundaries),
                row_boundaries=tuple(row_boundaries),
                rows=tuple(table_rows),
                cells=tuple(cells),
                header_row_count=1,
                has_borders=False,
            )
        )
    return tuple(tables)


