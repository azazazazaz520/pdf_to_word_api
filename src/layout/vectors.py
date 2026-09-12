"""矢量对象与线条的提取。

负责线条的聚类、覆盖率计算与表格边界线判定，以及矩形、多段路径等
矢量对象的抽取；表格模块依赖其中的边界线工具。
"""

from __future__ import annotations
from typing import Any
from .models import (
    PdfVectorObject,
    _COORDINATE_TOLERANCE,
    _MAX_LINE_THICKNESS,
    _MIN_HORIZONTAL_LINE_LENGTH,
    _MIN_VERTICAL_LINE_LENGTH,
    _HorizontalLine,
    _VerticalLine,
)
from .text import (
    _object_fill_color,
    _object_stroke_color,
    _stroke_width,
)
from ctypes import c_float, c_int
import pypdfium2.raw as pdfium_raw
from .models import _COORDINATE_TOLERANCE, _MIN_BOUNDARY_COVERAGE, _PAGEOBJ_PATH

def _merge_intervals(
    intervals: list[tuple[float, float]],
) -> list[tuple[float, float]]:
    if not intervals:
        return []
    merged: list[list[float]] = []
    for start, end in sorted(intervals):
        if not merged or start > merged[-1][1] + _COORDINATE_TOLERANCE:
            merged.append([start, end])
        else:
            merged[-1][1] = max(merged[-1][1], end)
    return [(start, end) for start, end in merged]


def _cluster_horizontal_lines(
    segments: list[_HorizontalLine],
) -> list[_HorizontalLine]:
    clusters: list[list[_HorizontalLine]] = []
    for segment in sorted(segments, key=lambda item: item.top):
        if (
            clusters
            and abs(clusters[-1][-1].top - segment.top)
            <= _COORDINATE_TOLERANCE
        ):
            clusters[-1].append(segment)
        else:
            clusters.append([segment])
    return [
        _HorizontalLine(
            top=sum(item.top for item in cluster) / len(cluster),
            x0=min(item.x0 for item in cluster),
            x1=max(item.x1 for item in cluster),
            segments=tuple((item.x0, item.x1) for item in cluster),
        )
        for cluster in clusters
    ]


def _cluster_vertical_lines(
    segments: list[_VerticalLine],
) -> list[_VerticalLine]:
    clusters: list[list[_VerticalLine]] = []
    for segment in sorted(segments, key=lambda item: item.x):
        if (
            clusters
            and abs(clusters[-1][-1].x - segment.x)
            <= _COORDINATE_TOLERANCE
        ):
            clusters[-1].append(segment)
        else:
            clusters.append([segment])
    merged: list[_VerticalLine] = []
    for cluster in clusters:
        intervals = _merge_intervals(
            [(item.top, item.bottom) for item in cluster]
        )
        for index, (top, bottom) in enumerate(intervals):
            merged.append(
                _VerticalLine(
                    x=sum(item.x for item in cluster) / len(cluster),
                    top=top,
                    bottom=bottom,
                )
            )
    return merged


def _extract_vector_lines(
    page: Any,
    page_height: float,
) -> tuple[list[_HorizontalLine], list[_VerticalLine]]:
    horizontal: list[_HorizontalLine] = []
    vertical: list[_VerticalLine] = []
    for obj in page.get_objects(filter=[_PAGEOBJ_PATH]):
        x0, y0, x1, y1 = (float(value) for value in obj.get_bounds())
        width = abs(x1 - x0)
        height = abs(y1 - y0)
        if height <= _MAX_LINE_THICKNESS and width >= _MIN_HORIZONTAL_LINE_LENGTH:
            horizontal.append(
                _HorizontalLine(
                    top=page_height - (y0 + y1) / 2,
                    x0=min(x0, x1),
                    x1=max(x0, x1),
                )
            )
        elif width <= _MAX_LINE_THICKNESS and height >= _MIN_VERTICAL_LINE_LENGTH:
            vertical.append(
                _VerticalLine(
                    x=(x0 + x1) / 2,
                    top=page_height - max(y0, y1),
                    bottom=page_height - min(y0, y1),
                )
            )
    return _cluster_horizontal_lines(horizontal), _cluster_vertical_lines(vertical)


def _vertical_coverage(
    lines: list[_VerticalLine],
    *,
    top: float,
    bottom: float,
) -> float:
    if bottom <= top:
        return 0.0
    intervals = _merge_intervals(
        [
            (max(line.top, top), min(line.bottom, bottom))
            for line in lines
            if line.bottom > top and line.top < bottom
        ]
    )
    covered = sum(max(end - start, 0.0) for start, end in intervals)
    return covered / (bottom - top)


def _horizontal_coverage(
    lines: list[_HorizontalLine],
    *,
    top: float,
    x0: float,
    x1: float,
) -> float:
    if x1 <= x0:
        return 0.0
    intervals: list[tuple[float, float]] = []
    for line in lines:
        if abs(line.top - top) > _COORDINATE_TOLERANCE:
            continue
        segments = line.segments or ((line.x0, line.x1),)
        intervals.extend(
            (
                max(segment_x0, x0),
                min(segment_x1, x1),
            )
            for segment_x0, segment_x1 in segments
            if segment_x1 > x0 and segment_x0 < x1
        )
    merged = _merge_intervals(intervals)
    covered = sum(max(end - start, 0.0) for start, end in merged)
    return covered / (x1 - x0)


def _has_vertical_boundary(
    lines: list[_VerticalLine],
    *,
    x: float,
    top: float,
    bottom: float,
) -> bool:
    return _vertical_coverage(
        [line for line in lines if abs(line.x - x) <= _COORDINATE_TOLERANCE],
        top=top,
        bottom=bottom,
    ) >= _MIN_BOUNDARY_COVERAGE


def _has_horizontal_boundary(
    lines: list[_HorizontalLine],
    *,
    top: float,
    x0: float,
    x1: float,
) -> bool:
    return (
        _horizontal_coverage(lines, top=top, x0=x0, x1=x1)
        >= _MIN_BOUNDARY_COVERAGE
    )


def _path_segments(
    page_object: Any,
    page_height: float,
) -> tuple[tuple[tuple[str, float, float], ...], bool]:
    """读取路径对象的分段，坐标转换为左上角原点。"""
    try:
        count = int(pdfium_raw.FPDFPath_CountSegments(page_object))
    except Exception:
        return (), False
    labels = {
        pdfium_raw.FPDF_SEGMENT_LINETO: "L",
        pdfium_raw.FPDF_SEGMENT_BEZIERTO: "C",
        pdfium_raw.FPDF_SEGMENT_MOVETO: "M",
    }
    segments: list[tuple[str, float, float]] = []
    closed = False
    for index in range(count):
        try:
            segment = pdfium_raw.FPDFPath_GetPathSegment(page_object, index)
            point_x, point_y = c_float(), c_float()
            if not pdfium_raw.FPDFPathSegment_GetPoint(
                segment, point_x, point_y
            ):
                continue
            try:
                if pdfium_raw.FPDFPathSegment_GetClose(segment):
                    closed = True
            except Exception:
                pass
            kind = int(pdfium_raw.FPDFPathSegment_GetType(segment))
            label = labels.get(kind, "M")
            segments.append(
                (
                    label,
                    float(point_x.value),
                    page_height - float(point_y.value),
                )
            )
        except Exception:
            continue
    return tuple(segments), closed


def _path_draw_mode(page_object: Any) -> tuple[bool, bool]:
    """返回路径的填充/描边绘制模式。"""
    fill_mode = c_int()
    stroke_flag = c_int()
    try:
        if not pdfium_raw.FPDFPath_GetDrawMode(
            page_object, fill_mode, stroke_flag
        ):
            return True, True
    except Exception:
        return True, True
    return bool(fill_mode.value), bool(stroke_flag.value)


def _has_dash(page_object: Any) -> bool:
    try:
        return int(pdfium_raw.FPDFPageObj_GetDashCount(page_object)) > 0
    except Exception:
        return False


def _extract_vector_objects(
    page: Any,
    page_height: float,
) -> tuple[PdfVectorObject, ...]:
    """提取页面中的线条、矩形和路径对象，保留 z-order 与样式。"""
    try:
        objects = list(page.get_objects())
    except Exception:
        return ()
    vectors: list[PdfVectorObject] = []
    for object_index, page_object in enumerate(objects):
        try:
            if page_object.type != _PAGEOBJ_PATH:
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
            width = bbox[2] - bbox[0]
            height = bbox[3] - bbox[1]
            fill_mode, stroke_flag = _path_draw_mode(page_object)
            stroke_color = (
                _object_stroke_color(page_object) if stroke_flag else None
            )
            fill_color = _object_fill_color(page_object) if fill_mode else None
            if (
                stroke_color is None
                and fill_color is None
                and width < 0.5
                and height < 0.5
            ):
                continue
            if width < 0.4 and height < 0.4:
                continue
            segments, closed = _path_segments(page_object, page_height)
            labels = {segment[0] for segment in segments}
            line_like = (
                not closed
                and len(segments) == 2
                and labels <= {"M", "L"}
            )
            rect_like = (
                closed
                and 3 <= len(segments) <= 6
                and labels <= {"M", "L"}
            )
            if line_like:
                kind = "line"
            elif rect_like:
                kind = "rect"
            else:
                kind = "path"
            complex_path = len(segments) > 48 or (
                kind == "path" and len(segments) > 24
            )
            vectors.append(
                PdfVectorObject(
                    kind=kind,
                    bbox=bbox,
                    stroke_color=stroke_color,
                    fill_color=fill_color,
                    stroke_width=_stroke_width(page_object),
                    segments=segments,
                    closed=closed,
                    complex=complex_path,
                    dashed=_has_dash(page_object),
                    filled=bool(fill_mode),
                    stroked=bool(stroke_flag),
                    z_order=object_index,
                )
            )
        finally:
            close_object = getattr(page_object, "close", None)
            if close_object is not None:
                close_object()
    return tuple(vectors)


