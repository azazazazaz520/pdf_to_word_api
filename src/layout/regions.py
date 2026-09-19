"""页面区域适配器、坐标换算和局部文字合并。

区域模块只负责几何契约和内容归属，不创建新的 OCR 模型生命周期。这样可以
在已有文字层、版面模型和局部识别之间复用同一套页面坐标。
"""

from __future__ import annotations

from dataclasses import dataclass, replace
import math
import re
from typing import Any, Iterable

from .models import (
    PdfImageBlock,
    PdfLayoutRegion,
    PdfTable,
    PdfTextLine,
    PdfVectorObject,
)


@dataclass(frozen=True)
class PageCoordinateTransform:
    """把模型像素坐标转换到页面左上角点坐标。"""

    pixel_width: float
    pixel_height: float
    page_width: float
    page_height: float
    rotation: int = 0
    crop_x: float = 0.0
    crop_y: float = 0.0

    def pixel_to_page(self, x: float, y: float) -> tuple[float, float]:
        rotation = int(self.rotation) % 360
        if rotation in {90, 270}:
            display_width, display_height = self.page_height, self.page_width
        else:
            display_width, display_height = self.page_width, self.page_height
        display_x = float(x) * display_width / max(self.pixel_width, 1.0)
        display_y = float(y) * display_height / max(self.pixel_height, 1.0)
        if rotation == 90:
            page_x, page_y = display_y, self.page_height - display_x
        elif rotation == 180:
            page_x, page_y = self.page_width - display_x, self.page_height - display_y
        elif rotation == 270:
            page_x, page_y = self.page_width - display_y, display_x
        else:
            page_x, page_y = display_x, display_y
        return page_x + self.crop_x, page_y + self.crop_y

    def bbox_to_page(
        self,
        bbox: Iterable[float],
    ) -> tuple[float, float, float, float]:
        x0, y0, x1, y1 = (float(value) for value in tuple(bbox)[:4])
        points = (
            self.pixel_to_page(x0, y0),
            self.pixel_to_page(x0, y1),
            self.pixel_to_page(x1, y0),
            self.pixel_to_page(x1, y1),
        )
        return (
            min(point[0] for point in points),
            min(point[1] for point in points),
            max(point[0] for point in points),
            max(point[1] for point in points),
        )


def _bbox_area(bbox: tuple[float, float, float, float]) -> float:
    return max(bbox[2] - bbox[0], 0.0) * max(bbox[3] - bbox[1], 0.0)


def _intersection_area(
    left: tuple[float, float, float, float],
    right: tuple[float, float, float, float],
) -> float:
    return max(min(left[2], right[2]) - max(left[0], right[0]), 0.0) * max(
        min(left[3], right[3]) - max(left[1], right[1]),
        0.0,
    )


def _line_source_ids(line: PdfTextLine) -> tuple[str, ...]:
    return tuple(
        dict.fromkeys(
            char_id
            for char_id in getattr(line, "source_char_ids", ())
            if char_id
        )
    )


def _source_ids_in_bbox(
    lines: Iterable[PdfTextLine],
    bbox: tuple[float, float, float, float],
) -> tuple[str, ...]:
    source_ids: list[str] = []
    for line in lines:
        if _intersection_area(
            (line.x0, line.top, line.x1, line.bottom),
            bbox,
        ) <= 0:
            continue
        source_ids.extend(_line_source_ids(line))
    return tuple(dict.fromkeys(source_ids))


def build_geometry_regions(
    *,
    page_number: int,
    page_width: float,
    page_height: float,
    lines: Iterable[PdfTextLine],
    tables: Iterable[PdfTable] = (),
    images: Iterable[PdfImageBlock] = (),
    vectors: Iterable[PdfVectorObject] = (),
    columns: tuple[float, ...] = (),
) -> tuple[PdfLayoutRegion, ...]:
    """根据现有页面对象建立稳定的几何区域。"""
    collected_lines = tuple(lines)
    regions: list[PdfLayoutRegion] = []
    for index, table in enumerate(tables):
        regions.append(
            PdfLayoutRegion(
                region_id=f"table-{index}",
                page_number=page_number,
                kind="table",
                bbox=table.bbox,
                confidence=0.95,
                source="geometry-table",
                source_char_ids=tuple(table.source_char_ids)
                or _source_ids_in_bbox(collected_lines, table.bbox),
            )
        )
    for index, image in enumerate(images):
        regions.append(
            PdfLayoutRegion(
                region_id=f"image-{index}",
                page_number=page_number,
                kind="image",
                bbox=image.bbox,
                confidence=0.9,
                source="geometry-image",
                source_char_ids=_source_ids_in_bbox(collected_lines, image.bbox),
            )
        )
    for index, vector in enumerate(vectors):
        if vector.width < 8.0 and vector.height < 8.0:
            continue
        regions.append(
            PdfLayoutRegion(
                region_id=f"vector-{index}",
                page_number=page_number,
                kind="figure",
                bbox=vector.bbox,
                confidence=0.75,
                source="geometry-vector",
                source_char_ids=_source_ids_in_bbox(collected_lines, vector.bbox),
            )
        )

    body_lines = [
        line
        for line in collected_lines
        if not any(
            _intersection_area(
                (line.x0, line.top, line.x1, line.bottom), table.bbox
            ) > 0
            for table in tables
        )
    ]
    if body_lines:
        boundaries = tuple(sorted(float(value) for value in columns))
        groups: dict[int, list[PdfTextLine]] = {}
        for line in body_lines:
            if line.width >= page_width * 0.65 or not boundaries:
                group_id = -1
            else:
                group_id = sum(line.center_x > boundary for boundary in boundaries)
            groups.setdefault(group_id, []).append(line)
        for index, group in enumerate(
            sorted(groups.items(), key=lambda item: (item[0] == -1, item[0]))
        ):
            group_id, group_lines = group
            bbox = (
                min(line.x0 for line in group_lines),
                min(line.top for line in group_lines),
                max(line.x1 for line in group_lines),
                max(line.bottom for line in group_lines),
            )
            regions.append(
                PdfLayoutRegion(
                    region_id="body-full" if group_id == -1 else f"body-column-{group_id}",
                    page_number=page_number,
                    kind="body",
                    bbox=bbox,
                    confidence=0.8,
                    source="geometry-text",
                    column_id=None if group_id == -1 else str(group_id),
                    source_char_ids=_source_ids_in_bbox(collected_lines, bbox),
                )
            )
    return tuple(regions)


def region_for_bbox(
    bbox: tuple[float, float, float, float],
    regions: Iterable[PdfLayoutRegion],
) -> PdfLayoutRegion | None:
    """返回覆盖面积最大的区域，表格和图片在相同时优先。"""
    candidates: list[tuple[float, int, PdfLayoutRegion]] = []
    priority = {"table": 3, "image": 2, "figure": 1, "body": 0}
    for region in regions:
        overlap = _intersection_area(bbox, region.bbox)
        if overlap <= 0:
            continue
        candidates.append((overlap / max(_bbox_area(bbox), 1.0), priority.get(region.kind, 0), region))
    return max(candidates, key=lambda item: (item[0], item[1]))[2] if candidates else None


def normalize_model_regions(
    detections: Iterable[Any],
    *,
    page_number: int,
    transform: PageCoordinateTransform,
    lines: Iterable[PdfTextLine] = (),
    include_source_char_ids: bool = True,
) -> tuple[PdfLayoutRegion, ...]:
    """把版面模型输出适配成稳定的页面区域。"""
    regions: list[PdfLayoutRegion] = []
    for index, detection in enumerate(detections):
        if isinstance(detection, dict):
            raw_bbox = (
                detection.get("bbox")
                or detection.get("box")
                or detection.get("coordinate")
            )
            kind = str(detection.get("kind") or detection.get("label") or "body").lower()
            confidence = float(detection.get("confidence", detection.get("score", 0.0)) or 0.0)
            order_hint = detection.get("order_hint")
        else:
            raw_bbox = getattr(detection, "bbox", None)
            kind = str(getattr(detection, "kind", "body"))
            confidence = float(getattr(detection, "confidence", 0.0) or 0.0)
            order_hint = getattr(detection, "order_hint", None)
        if raw_bbox is None or len(raw_bbox) < 4:
            continue
        bbox = transform.bbox_to_page(raw_bbox)
        regions.append(
            PdfLayoutRegion(
                region_id=f"model-{page_number}-{index}",
                page_number=page_number,
                kind=kind,
                bbox=bbox,
                confidence=max(0.0, min(1.0, confidence)),
                source="layout-model",
                order_hint=int(order_hint) if order_hint is not None else None,
                source_char_ids=(
                    _source_ids_in_bbox(lines, bbox)
                    if include_source_char_ids
                    else ()
                ),
            )
        )
    return tuple(regions)


def _model_region_columns(
    regions: Iterable[PdfLayoutRegion],
    *,
    page_width: float,
) -> tuple[float, ...]:
    """从模型正文区域的中心位置推断局部栏界。"""
    centers = sorted(
        (region.bbox[0] + region.bbox[2]) / 2.0
        for region in regions
        if region.kind not in {"table", "image", "figure", "chart"}
        and region.bbox[2] - region.bbox[0] < page_width * 0.72
    )
    if len(centers) < 2:
        return ()
    clusters: list[list[float]] = []
    for center in centers:
        if clusters and center - clusters[-1][-1] <= page_width * 0.12:
            clusters[-1].append(center)
        else:
            clusters.append([center])
    if len(clusters) < 2:
        return ()
    centers = [sum(cluster) / len(cluster) for cluster in clusters]
    return tuple(
        (left + right) / 2.0 for left, right in zip(centers, centers[1:])
    )


def assign_model_regions_to_lines(
    lines: Iterable[PdfTextLine],
    *,
    model_regions: Iterable[PdfLayoutRegion],
    page_width: float,
    tables: Iterable[PdfTable] = (),
) -> tuple[tuple[PdfTextLine, ...], tuple[float, ...]]:
    """将模型正文区域映射到原生行，并返回模型推断的栏界。"""
    collected_lines = tuple(lines)
    collected_regions = tuple(model_regions)
    columns = _model_region_columns(collected_regions, page_width=page_width)
    table_boxes = tuple(table.bbox for table in tables)
    assigned: list[PdfTextLine] = []
    for line in collected_lines:
        line_bbox = (line.x0, line.top, line.x1, line.bottom)
        if any(_intersection_area(line_bbox, bbox) > 0 for bbox in table_boxes):
            assigned.append(line)
            continue
        candidates = [
            region
            for region in collected_regions
            if region.kind not in {"table", "image", "figure", "chart"}
            and _intersection_area(line_bbox, region.bbox) > 0
        ]
        if not candidates:
            assigned.append(line)
            continue
        region = max(
            candidates,
            key=lambda item: _intersection_area(line_bbox, item.bbox)
            / max(_bbox_area(line_bbox), 1.0),
        )
        if region.bbox[2] - region.bbox[0] >= page_width * 0.72:
            region_id = "full-width"
            column_id = None
        else:
            column_index = sum(line.center_x > boundary for boundary in columns)
            region_id = f"column-{column_index}"
            column_id = str(column_index)
        assigned.append(
            replace(line, region_id=region_id, column_id=column_id)
        )
    return tuple(assigned), columns


def merge_region_sets(
    geometry_regions: Iterable[PdfLayoutRegion],
    model_regions: Iterable[PdfLayoutRegion],
) -> tuple[PdfLayoutRegion, ...]:
    """合并模型和几何区域，避免同一内容被重复承担。"""
    result = list(geometry_regions)
    for model in model_regions:
        overlaps = [
            existing
            for existing in result
            if _intersection_area(existing.bbox, model.bbox)
            / max(min(_bbox_area(existing.bbox), _bbox_area(model.bbox)), 1.0)
            >= 0.65
        ]
        if overlaps:
            for existing in overlaps:
                result.remove(existing)
            source_ids = tuple(
                dict.fromkeys(
                    char_id
                    for existing in overlaps
                    for char_id in (*existing.source_char_ids, *model.source_char_ids)
                )
            )
            model = PdfLayoutRegion(
                **{
                    **model.__dict__,
                    "source_char_ids": source_ids,
                }
            )
        result.append(model)
    return tuple(sorted(result, key=lambda region: (region.bbox[1], region.bbox[0], region.region_id)))


def merge_local_ocr_lines(
    native_lines: Iterable[PdfTextLine],
    ocr_lines: Iterable[PdfTextLine],
    *,
    region: PdfLayoutRegion,
    garbled_ratio: float = 0.25,
) -> tuple[tuple[PdfTextLine, ...], dict[str, Any]]:
    """按区域合并 OCR 候选，保留可用原生文字并记录来源。"""
    native = [
        line
        for line in native_lines
        if _intersection_area((line.x0, line.top, line.x1, line.bottom), region.bbox) > 0
    ]
    candidates = [
        line
        for line in ocr_lines
        if _intersection_area((line.x0, line.top, line.x1, line.bottom), region.bbox) > 0
    ]
    native_text = "".join(line.text for line in native)
    bad_chars = sum(
        character in "\ufffd\ufffe" or ord(character) < 32
        for character in native_text
    )
    native_is_usable = bool(native_text.strip()) and bad_chars / max(len(native_text), 1) <= garbled_ratio
    if native_is_usable or not candidates:
        return tuple(native), {
            "source": "native",
            "native_line_count": len(native),
            "ocr_line_count": 0,
            "deduplicated": 0,
            "reason": "native_available" if native_is_usable else "ocr_empty",
        }
    deduplicated = 0
    merged: list[PdfTextLine] = list(native)
    for candidate in candidates:
        normalized = re.sub(r"\s+", "", candidate.text).lower()
        duplicate = any(
            normalized
            and normalized == re.sub(r"\s+", "", line.text).lower()
            and _intersection_area(
                (candidate.x0, candidate.top, candidate.x1, candidate.bottom),
                (line.x0, line.top, line.x1, line.bottom),
            ) > 0
            for line in merged
        )
        if duplicate:
            deduplicated += 1
            continue
        merged.append(candidate)
    return tuple(sorted(merged, key=lambda line: (line.top, line.x0))), {
        "source": "local_ocr",
        "native_line_count": len(native),
        "ocr_line_count": len(candidates),
        "deduplicated": deduplicated,
        "reason": "native_missing_or_garbled",
    }
