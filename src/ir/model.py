from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class IRWarning:
    """统一的页级或文档级质量警告。"""

    code: str
    message: str
    page: int | None = None
    severity: str = "warning"
    meta: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        payload = {
            "code": self.code,
            "message": self.message,
            "page": self.page,
            "severity": self.severity,
        }
        if self.meta:
            payload["meta"] = dict(self.meta)
        return payload


@dataclass(frozen=True)
class IRTextSpan:
    """行内一段同字体文本，用于精确还原混排字体。"""

    text: str
    bbox: tuple[float, float, float, float]
    font_name: str = ""
    pdf_font_name: str = ""
    font_size: float = 0.0
    color: tuple[int, int, int] | None = None
    bold: bool = False
    italic: bool = False
    rotation: float = 0.0
    z_order: int = 0
    substituted: bool = False
    fallback_reason: str = ""

    @property
    def width(self) -> float:
        return max(self.bbox[2] - self.bbox[0], 0.0)

    def to_dict(self) -> dict[str, Any]:
        return {
            "text": self.text,
            "bbox": [round(value, 3) for value in self.bbox],
            "font_name": self.font_name,
            "pdf_font_name": self.pdf_font_name,
            "font_size": round(self.font_size, 3),
            "color": list(self.color) if self.color else None,
            "bold": self.bold,
            "italic": self.italic,
            "rotation": round(self.rotation, 3),
            "z_order": self.z_order,
            "font_substituted": self.substituted,
            "font_fallback_reason": self.fallback_reason,
        }


@dataclass(frozen=True)
class IRTextLine:
    """高保真导出的行级定位信息。

    ``bbox`` 使用 PDF 页面坐标（左上角原点，单位为 point），
    导出时直接映射到 OOXML 的绝对定位偏移量。
    """

    text: str
    bbox: tuple[float, float, float, float]
    font_name: str = ""
    font_size: float = 0.0
    color: tuple[int, int, int] | None = None
    bold: bool = False
    italic: bool = False
    alignment: str = "left"
    line_spacing: float = 0.0
    first_line_indent: float = 0.0
    rotation: float = 0.0
    z_order: int = 0
    confidence: float | None = None
    spans: tuple[IRTextSpan, ...] = ()
    pdf_font_name: str = ""
    substituted: bool = False
    fallback_reason: str = ""

    @property
    def width(self) -> float:
        return max(self.bbox[2] - self.bbox[0], 0.0)

    @property
    def height(self) -> float:
        return max(self.bbox[3] - self.bbox[1], 0.0)

    def to_dict(self) -> dict[str, Any]:
        return {
            "text": self.text,
            "bbox": [round(value, 3) for value in self.bbox],
            "font_name": self.font_name,
            "font_size": round(self.font_size, 3),
            "color": list(self.color) if self.color else None,
            "bold": self.bold,
            "italic": self.italic,
            "alignment": self.alignment,
            "line_spacing": round(self.line_spacing, 4),
            "first_line_indent": round(self.first_line_indent, 3),
            "rotation": round(self.rotation, 3),
            "z_order": self.z_order,
            "confidence": self.confidence,
            "pdf_font_name": self.pdf_font_name,
            "font_substituted": self.substituted,
            "font_fallback_reason": self.fallback_reason,
            "spans": [span.to_dict() for span in self.spans],
        }


_TEXT_KINDS = frozenset(
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
        "formula",
        "caption",
        "quote",
        "code",
    }
)
_MEDIA_KINDS = frozenset({"image", "page_image"})
_TABLE_KINDS = frozenset({"table", "html_table"})
_HEADING_KINDS = frozenset({"title", "heading1", "heading2", "heading3"})
_LIST_KINDS = frozenset({"ordered", "bullet", "plugin"})
_VECTOR_KINDS = frozenset({"vector"})
_LAYER_HEADER = "header"
_LAYER_FOOTER = "footer"


@dataclass
class IRBlock:
    """文档 IR 的基本块，覆盖文字、表格、图片、公式和矢量对象。

    ``bbox``/字体/对齐等字段用于高保真绝对定位导出；
    ``lines`` 保存行级定位信息，缺失时回退到块级 bbox。
    """

    kind: str
    page: int
    text: str = ""
    role: str = ""
    level: int = 0
    bbox: tuple[float, float, float, float] | None = None
    source: str = "text"
    confidence: float | None = None
    table: Any = None
    image_bytes: bytes | None = None
    image_width: float = 0.0
    image_height: float = 0.0
    image_alt: str = ""
    meta: dict[str, Any] = field(default_factory=dict)
    warnings: tuple[IRWarning, ...] = ()
    font_name: str = ""
    font_size: float = 0.0
    color: tuple[int, int, int] | None = None
    bold: bool = False
    italic: bool = False
    alignment: str = ""
    line_spacing: float = 0.0
    first_line_indent: float = 0.0
    z_order: int = 0
    rotation: float = 0.0
    layer: str = "body"
    lines: tuple[IRTextLine, ...] = ()
    vector: Any = None
    latex: str = ""
    fallback_image: bytes | None = None
    fallback_reason: str = ""

    @property
    def is_text(self) -> bool:
        return self.kind in _TEXT_KINDS

    @property
    def is_media(self) -> bool:
        return self.kind in _MEDIA_KINDS

    @property
    def is_table(self) -> bool:
        return self.kind in _TABLE_KINDS

    @property
    def is_heading(self) -> bool:
        return self.kind in _HEADING_KINDS

    @property
    def is_list(self) -> bool:
        return self.kind in _LIST_KINDS

    @property
    def is_vector(self) -> bool:
        return self.kind in _VECTOR_KINDS

    @property
    def has_fallback(self) -> bool:
        return bool(self.fallback_reason) or self.fallback_image is not None

    @property
    def width(self) -> float:
        if not self.bbox:
            return 0.0
        return max(self.bbox[2] - self.bbox[0], 0.0)

    @property
    def height(self) -> float:
        if not self.bbox:
            return 0.0
        return max(self.bbox[3] - self.bbox[1], 0.0)

    @property
    def area(self) -> float:
        return self.width * self.height

    def line_dicts(self) -> list[dict[str, Any]]:
        return [line.to_dict() for line in self.lines]

    def to_dict(self) -> dict[str, Any]:
        return {
            "kind": self.kind,
            "page": self.page,
            "text": self.text,
            "role": self.role,
            "level": self.level,
            "bbox": [round(value, 3) for value in self.bbox] if self.bbox else None,
            "source": self.source,
            "confidence": self.confidence,
            "font_name": self.font_name,
            "font_size": round(self.font_size, 3),
            "color": list(self.color) if self.color else None,
            "bold": self.bold,
            "italic": self.italic,
            "alignment": self.alignment,
            "line_spacing": round(self.line_spacing, 4),
            "first_line_indent": round(self.first_line_indent, 3),
            "z_order": self.z_order,
            "rotation": round(self.rotation, 3),
            "layer": self.layer,
            "line_count": len(self.lines),
            "fallback_reason": self.fallback_reason,
            "has_fallback_image": self.fallback_image is not None,
        }


@dataclass
class IRPage:
    """单页 IR。一个 IRPage 对应 Word 输出中的一个 section。"""

    page_number: int
    width: float
    height: float
    route: str
    blocks: list[IRBlock] = field(default_factory=list)
    confidence: float | None = None
    warnings: list[IRWarning] = field(default_factory=list)
    editable: bool = True
    reconstruction_confidence: float | None = None
    fidelity: dict[str, Any] = field(default_factory=dict)
    header_footer_native: bool = False

    def add_warning(self, warning: IRWarning) -> None:
        self.warnings.append(warning)

    @property
    def table_count(self) -> int:
        return sum(block.is_table for block in self.blocks)

    @property
    def media_count(self) -> int:
        return sum(block.is_media for block in self.blocks)

    @property
    def heading_count(self) -> int:
        return sum(block.is_heading for block in self.blocks)

    @property
    def list_count(self) -> int:
        return sum(block.is_list for block in self.blocks)

    @property
    def formula_count(self) -> int:
        return sum(block.kind == "formula" for block in self.blocks)

    @property
    def code_count(self) -> int:
        return sum(block.kind == "code" for block in self.blocks)

    @property
    def vector_count(self) -> int:
        return sum(block.is_vector for block in self.blocks)

    @property
    def header_blocks(self) -> list[IRBlock]:
        return [block for block in self.blocks if block.layer == _LAYER_HEADER]

    @property
    def footer_blocks(self) -> list[IRBlock]:
        return [block for block in self.blocks if block.layer == _LAYER_FOOTER]

    @property
    def body_blocks(self) -> list[IRBlock]:
        return [
            block
            for block in self.blocks
            if block.layer not in {_LAYER_HEADER, _LAYER_FOOTER}
        ]

    @property
    def fallback_blocks(self) -> list[IRBlock]:
        return [block for block in self.blocks if block.has_fallback]

    @property
    def fallback_regions(self) -> list[dict[str, Any]]:
        regions: list[dict[str, Any]] = []
        for block in self.blocks:
            if not block.has_fallback:
                continue
            regions.append(
                {
                    "kind": block.kind,
                    "layer": block.layer,
                    "bbox": (
                        [round(value, 3) for value in block.bbox]
                        if block.bbox
                        else None
                    ),
                    "reason": block.fallback_reason or "explicit_fallback_image",
                }
            )
        return regions


@dataclass
class IRDocument:
    """统一文档 IR，是质量报告和 DOCX 导出的唯一输入。"""

    pages: list[IRPage] = field(default_factory=list)
    title: str = ""
    warnings: list[IRWarning] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def source_page_count(self) -> int:
        return len(self.pages)

    @property
    def route_summary(self) -> dict[str, int]:
        summary: dict[str, int] = {}
        for page in self.pages:
            summary[page.route] = summary.get(page.route, 0) + 1
        return summary

    @property
    def table_count(self) -> int:
        return sum(page.table_count for page in self.pages)

    @property
    def media_count(self) -> int:
        return sum(page.media_count for page in self.pages)

    @property
    def formula_count(self) -> int:
        return sum(page.formula_count for page in self.pages)

    @property
    def code_count(self) -> int:
        return sum(page.code_count for page in self.pages)

    @property
    def heading_count(self) -> int:
        return sum(page.heading_count for page in self.pages)

    @property
    def list_count(self) -> int:
        return sum(page.list_count for page in self.pages)

    @property
    def vector_count(self) -> int:
        return sum(page.vector_count for page in self.pages)

    @property
    def editable_page_count(self) -> int:
        return sum(page.editable for page in self.pages)

    @property
    def text_line_count(self) -> int:
        return sum(len(block.lines) for page in self.pages for block in page.blocks)

    def font_requests(self) -> list[dict[str, Any]]:
        """收集 IR 中出现的字体（用于字体嵌入计划）。"""
        requests: list[dict[str, Any]] = []
        seen: set[tuple[str, str, bool, bool]] = set()
        for page in self.pages:
            for block in page.blocks:
                for line in block.lines:
                    candidates = line.spans or (line,)
                    for item in candidates:
                        raw = getattr(item, "pdf_font_name", "") or ""
                        name = getattr(item, "font_name", "") or ""
                        bold = bool(getattr(item, "bold", False))
                        italic = bool(getattr(item, "italic", False))
                        key = (raw, name, bold, italic)
                        if key in seen:
                            continue
                        seen.add(key)
                        requests.append(
                            {
                                "pdf_font_name": raw,
                                "font_name": name,
                                "bold": bold,
                                "italic": italic,
                            }
                        )
        return requests

    def font_usage(self) -> dict[str, Any]:
        """统计 PDF 原始字体 -> Word 字体的使用与替换情况。"""
        from ..fonts.resolver import summarize_font_usage

        entries: list[dict[str, Any]] = []
        for page in self.pages:
            for block in page.blocks:
                for line in block.lines:
                    if line.spans:
                        for span in line.spans:
                            entries.append(
                                {
                                    "font_name": span.font_name,
                                    "pdf_font_name": span.pdf_font_name,
                                    "bold": span.bold,
                                    "italic": span.italic,
                                    "font_substituted": span.substituted,
                                    "font_fallback_reason": span.fallback_reason,
                                }
                            )
                        continue
                    entries.append(
                        {
                            "font_name": line.font_name,
                            "pdf_font_name": line.pdf_font_name,
                            "bold": line.bold,
                            "italic": line.italic,
                            "font_substituted": line.substituted,
                            "font_fallback_reason": line.fallback_reason,
                        }
                    )
        return summarize_font_usage(entries)

    def fidelity_report(self, *, compact: bool = False) -> dict[str, Any]:
        """汇总逐页重建置信度和图片兜底区域。"""
        pages: list[dict[str, Any]] = []
        fallback_region_count = 0
        fallback_pages: list[int] = []
        confidences: list[float] = []
        total_bbox_error = 0.0
        total_font_error = 0.0
        error_samples = 0
        for page in self.pages:
            page_fidelity = dict(page.fidelity or {})
            regions = list(page_fidelity.get("fallback_regions") or [])
            if not regions:
                regions = page.fallback_regions
            confidence = page_fidelity.get("rebuild_confidence")
            if confidence is None:
                confidence = page.reconstruction_confidence
            if isinstance(confidence, (int, float)):
                confidences.append(float(confidence))
            if page_fidelity.get("bbox_error_max") is not None:
                total_bbox_error += float(page_fidelity["bbox_error_max"])
            if page_fidelity.get("font_size_error_max") is not None:
                total_font_error += float(page_fidelity["font_size_error_max"])
            error_samples += 1
            fallback_region_count += len(regions)
            if regions:
                fallback_pages.append(page.page_number)
            entry = {
                "page": page.page_number,
                "route": page.route,
                "page_size": [round(page.width, 2), round(page.height, 2)],
                "block_count": len(page.blocks),
                "text_line_count": sum(len(block.lines) for block in page.blocks),
                "fallback_block_count": len(page.fallback_blocks),
                "fallback_regions": regions,
                "rebuild_confidence": (
                    round(float(confidence), 4)
                    if isinstance(confidence, (int, float))
                    else None
                ),
                "header_footer_native": page.header_footer_native,
                "editable": page.editable,
            }
            for key in (
                "native_block_count",
                "image_fallback_count",
                "page_image_fallback",
                "bbox_error_max",
                "font_size_error_max",
                "placements",
            ):
                if key in page_fidelity:
                    entry[key] = page_fidelity[key]
            pages.append(entry)
        ordered_confidence = sorted(confidences)
        report = {
            "mode": self.metadata.get("export_mode"),
            "page_count": self.source_page_count,
            "fallback_region_count": fallback_region_count,
            "fallback_pages": sorted(set(fallback_pages)),
            "page_image_fallback_pages": sorted(
                page.page_number
                for page in self.pages
                if (page.fidelity or {}).get("page_image_fallback")
            ),
            "mean_rebuild_confidence": (
                round(sum(ordered_confidence) / len(ordered_confidence), 4)
                if ordered_confidence
                else None
            ),
            "min_rebuild_confidence": (
                round(ordered_confidence[0], 4) if ordered_confidence else None
            ),
            "max_bbox_error": (
                round(max(
                    float((page.fidelity or {}).get("bbox_error_max") or 0.0)
                    for page in self.pages
                ), 3)
                if self.pages
                else None
            ),
            "max_font_size_error": (
                round(max(
                    float((page.fidelity or {}).get("font_size_error_max") or 0.0)
                    for page in self.pages
                ), 3)
                if self.pages
                else None
            ),
            "mean_bbox_error": (
                round(total_bbox_error / error_samples, 3) if error_samples else None
            ),
            "mean_font_size_error": (
                round(total_font_error / error_samples, 3) if error_samples else None
            ),
            "unreconstructable_regions": _summarize_regions(pages),
            "pages": pages,
        }
        if compact and len(pages) > 200:
            report["pages"] = [
                entry
                for entry in pages
                if entry["fallback_regions"]
                or not entry["editable"]
                or (entry["rebuild_confidence"] or 1.0) < 0.98
            ]
            report["pages_compacted"] = True
        return report

    def quality_report(self, *, compact: bool = False) -> dict[str, Any]:
        """生成质量报告；compact 模式用于长文档，避免接口体积过大。"""
        collected_warnings: list[dict[str, Any]] = [
            item.to_dict() for item in self.warnings
        ]
        needs_review_pages: set[int] = set()
        page_results: list[dict[str, Any]] = []
        for page in self.pages:
            page_warnings = [item.to_dict() for item in page.warnings]
            collected_warnings.extend(page_warnings)
            if not page.editable:
                needs_review_pages.add(page.page_number)
            for warning in page.warnings:
                if warning.severity in {"warning", "error"}:
                    needs_review_pages.add(page.page_number)
            page_results.append(
                {
                    "page": page.page_number,
                    "route": page.route,
                    "confidence": page.confidence,
                    "editable": page.editable,
                    "block_count": len(page.blocks),
                    "text_line_count": sum(len(block.lines) for block in page.blocks),
                    "table_count": page.table_count,
                    "media_count": page.media_count,
                    "heading_count": page.heading_count,
                    "list_count": page.list_count,
                    "formula_count": page.formula_count,
                    "code_count": page.code_count,
                    "vector_count": page.vector_count,
                    "fallback_block_count": len(page.fallback_blocks),
                    "rebuild_confidence": page.reconstruction_confidence,
                    "warnings": page_warnings,
                }
            )
        warnings = _aggregate_info_warnings(collected_warnings)
        compacted = False
        if compact and len(page_results) > 200:
            selected: list[dict[str, Any]] = []
            for item in page_results:
                meaningful_warnings = [
                    warning
                    for warning in item["warnings"]
                    if warning.get("severity") in {"warning", "error"}
                ]
                if not item["editable"] or meaningful_warnings:
                    compacted_item = dict(item)
                    compacted_item["warnings"] = meaningful_warnings
                    selected.append(compacted_item)
            page_results = selected
            compacted = True
        intentional_blank_pages = [
            page.page_number for page in self.pages if page.route == "blank"
        ]
        effective_source_page_count = self.source_page_count - len(
            intentional_blank_pages
        )
        fidelity = None
        if any(page.fidelity for page in self.pages) or self.metadata.get(
            "fidelity_report"
        ):
            fidelity = self.fidelity_report(compact=compact)
        report = {
            "route_summary": self.route_summary,
            "source_page_count": self.source_page_count,
            "effective_source_page_count": effective_source_page_count,
            "intentional_blank_pages": intentional_blank_pages,
            "rendered_page_count": effective_source_page_count,
            "page_delta": 0,
            "table_count": self.table_count,
            "media_count": self.media_count,
            "formula_count": self.formula_count,
            "code_count": self.code_count,
            "heading_count": self.heading_count,
            "list_count": self.list_count,
            "vector_count": self.vector_count,
            "text_line_count": self.text_line_count,
            "editable_page_count": self.editable_page_count,
            "visual_only_page_count": sum(not page.editable for page in self.pages),
            "needs_review_pages": sorted(needs_review_pages),
            "warnings": warnings,
            "page_results": page_results,
            "page_results_compacted": compacted,
            "page_results_total": self.source_page_count,
            "engine": self.metadata.get("engine"),
            "model_version": self.metadata.get("model_version"),
            "route_mode": self.metadata.get("route_mode"),
            "export_mode": self.metadata.get("export_mode"),
            "fidelity": fidelity,
            "fonts": self.font_usage(),
        }
        return report


def _summarize_regions(pages: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """把逐页兜底区域压缩成按页统计的列表，便于接口消费。"""
    summaries: list[dict[str, Any]] = []
    for entry in pages:
        regions = entry.get("fallback_regions") or []
        if not regions:
            continue
        summaries.append(
            {
                "page": entry.get("page"),
                "count": len(regions),
                "regions": regions[:50],
            }
        )
    return summaries


def make_warning(
    code: str,
    message: str,
    *,
    page: int | None = None,
    severity: str = "warning",
    meta: dict[str, Any] | None = None,
) -> IRWarning:
    """快捷构造质量警告。"""
    return IRWarning(
        code=code,
        message=message,
        page=page,
        severity=severity,
        meta=dict(meta or {}),
    )


def _format_page_ranges(pages: list[int]) -> str:
    """把页码列表压缩成 1-3,5,8-10 形式。"""
    ordered = sorted(set(page for page in pages if page is not None))
    if not ordered:
        return ""
    ranges: list[str] = []
    start = previous = ordered[0]
    for page in ordered[1:]:
        if page == previous + 1:
            previous = page
            continue
        ranges.append(str(start) if start == previous else f"{start}-{previous}")
        start = previous = page
    ranges.append(str(start) if start == previous else f"{start}-{previous}")
    return ",".join(ranges)


def _aggregate_info_warnings(
    warnings: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """聚合页眉页脚等高频 info 警告，减少长文档报告体积。"""
    aggregated: list[dict[str, Any]] = []
    grouped: dict[tuple[str, str, str], list[int]] = {}
    for warning in warnings:
        code = str(warning.get("code") or "")
        severity = str(warning.get("severity") or "warning")
        message = str(warning.get("message") or "")
        page = warning.get("page")
        if code in {"header_footer_filtered", "fidelity_region_fallback"} and isinstance(page, int):
            grouped.setdefault((code, severity, message), []).append(page)
            continue
        aggregated.append(warning)
    for (code, severity, message), pages in grouped.items():
        aggregated.append(
            {
                "code": code,
                "message": message,
                "severity": severity,
                "page": None,
                "pages": _format_page_ranges(pages),
                "page_count": len(set(pages)),
            }
        )
    return aggregated
