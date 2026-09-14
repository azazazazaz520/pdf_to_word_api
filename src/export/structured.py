"""结构化流式 DOCX 导出。

普通正文按逻辑块写入 Word 段落；表格写成 ``w:tbl``，公式优先写成
OMML，图片写成内嵌图片。该路径依赖 Document IR 的阅读顺序，避免把每个
字符或每个文本片段都变成绝对定位对象。
"""

from __future__ import annotations

from html.parser import HTMLParser
from io import BytesIO
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from ..fonts.embedding import FontPlan, attach_embedded_fonts
from ..ir.model import IRBlock, IRDocument, IRPage, IRTextLine
from ..ooxml_positioning import detach_header_footer
from ..page_render import (
    _add_bookmark,
    _add_formula_omml_element,
    _add_toc_field,
    _bookmark_name,
)
from .document_setup import (
    StageCallback,
    _notify_stage,
    _remove_initial_empty_paragraph,
    _set_document_styles,
    _set_section_page,
    render_pdf_region,
)
from .table_render import _add_pdf_table


STRUCTURED_MODES = frozenset({"structured", "flow"})
FLOW_MARGIN_POINTS = 46.8
DENSE_VECTOR_LIMIT = 64
FLOW_TEXT_KINDS = frozenset(
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
        "caption",
        "quote",
        "code",
    }
)
_HEADING_LEVELS = {"title": 1, "heading1": 1, "heading2": 2, "heading3": 3}
_ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


@dataclass(frozen=True)
class _HtmlTableCell:
    text: str
    row_span: int = 1
    column_span: int = 1
    header: bool = False


class _HtmlTableParser(HTMLParser):
    """提取 OCR 表格 HTML 中的行和单元格文本。"""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.rows: list[list[_HtmlTableCell]] = []
        self._row: list[_HtmlTableCell] | None = None
        self._cell: list[str] | None = None
        self._cell_row_span = 1
        self._cell_column_span = 1
        self._cell_header = False

    def handle_starttag(
        self,
        tag: str,
        attrs: list[tuple[str, str | None]],
    ) -> None:
        if tag.lower() == "tr":
            self._row = []
        elif tag.lower() in {"td", "th"} and self._row is not None:
            self._cell = []
            attributes = {
                str(name).lower(): str(value or "")
                for name, value in attrs
            }
            self._cell_row_span = self._positive_span(attributes.get("rowspan"))
            self._cell_column_span = self._positive_span(
                attributes.get("colspan")
            )
            self._cell_header = tag.lower() == "th"

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in {"td", "th"}:
            if self._row is not None and self._cell is not None:
                value = " ".join("".join(self._cell).split())
                self._row.append(
                    _HtmlTableCell(
                        text=value,
                        row_span=self._cell_row_span,
                        column_span=self._cell_column_span,
                        header=self._cell_header,
                    )
                )
            self._cell = None
        elif tag == "tr" and self._row is not None:
            if self._row:
                self.rows.append(self._row)
            self._row = None

    def handle_data(self, data: str) -> None:
        if self._cell is not None:
            self._cell.append(data)

    @staticmethod
    def _positive_span(value: str | None) -> int:
        try:
            return max(int(value or "1"), 1)
        except ValueError:
            return 1


def _html_table_rows(value: str) -> list[list[_HtmlTableCell]]:
    parser = _HtmlTableParser()
    parser.feed(value)
    return parser.rows


def _set_run_style(
    run: Any,
    *,
    font_name: str = "",
    font_size: float = 0.0,
    bold: bool = False,
    italic: bool = False,
    color: tuple[int, int, int] | None = None,
) -> None:
    run.font.name = font_name or "Microsoft YaHei"
    if font_size > 0:
        run.font.size = Pt(font_size)
    run.bold = bool(bold)
    run.italic = bool(italic)
    if color is not None and len(color) >= 3:
        run.font.color.rgb = RGBColor(
            int(color[0]) & 0xFF,
            int(color[1]) & 0xFF,
            int(color[2]) & 0xFF,
        )
    properties = run._element.get_or_add_rPr()
    fonts = properties.rFonts
    if fonts is not None:
        fonts.set(qn("w:eastAsia"), font_name or "Microsoft YaHei")


def _planned_font(
    font_plan: FontPlan | None,
    *,
    raw_name: str,
    fallback_name: str,
    bold: bool,
    italic: bool,
) -> tuple[str, bool, bool]:
    if font_plan is None:
        return fallback_name or "Microsoft YaHei", bool(bold), bool(italic)
    name = font_plan.word_name_for(raw_name, fallback_name)
    planned_bold, planned_italic = font_plan.style_for(
        raw_name,
        bool(bold),
        bool(italic),
    )
    return name or fallback_name or "Microsoft YaHei", planned_bold, planned_italic


def _configure_paragraph(paragraph: Any, block: IRBlock | None = None) -> None:
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    if block is None:
        return
    alignment = str(block.alignment or "").lower()
    if alignment in _ALIGNMENTS:
        paragraph.alignment = _ALIGNMENTS[alignment]
    if block.line_spacing > 0:
        paragraph.paragraph_format.line_spacing = Pt(block.line_spacing)
    if block.first_line_indent > 0:
        paragraph.paragraph_format.first_line_indent = Pt(block.first_line_indent)


def _line_spans_match(line: IRTextLine, block: IRBlock) -> bool:
    spans = list(line.spans or ())
    return bool(
        spans
        and len(block.lines) == 1
        and block.text == line.text
        and "".join(span.text for span in spans) == line.text
    )


def _add_text_runs(
    paragraph: Any,
    block: IRBlock,
    *,
    font_plan: FontPlan | None = None,
) -> None:
    if len(block.lines) == 1 and _line_spans_match(block.lines[0], block):
        for span in block.lines[0].spans:
            font_name, bold, italic = _planned_font(
                font_plan,
                raw_name=str(getattr(span, "pdf_font_name", "") or ""),
                fallback_name=str(getattr(span, "font_name", "") or ""),
                bold=bool(getattr(span, "bold", False)),
                italic=bool(getattr(span, "italic", False)),
            )
            run = paragraph.add_run(span.text)
            _set_run_style(
                run,
                font_name=font_name,
                font_size=float(getattr(span, "font_size", 0.0) or 0.0),
                bold=bold,
                italic=italic,
                color=getattr(span, "color", None),
            )
        return
    text = block.text or "\n".join(line.text for line in block.lines)
    font_name, bold, italic = _planned_font(
        font_plan,
        raw_name=str(getattr(block, "pdf_font_name", "") or ""),
        fallback_name=str(getattr(block, "font_name", "") or ""),
        bold=bool(getattr(block, "bold", False)),
        italic=bool(getattr(block, "italic", False)),
    )
    run = paragraph.add_run(text)
    _set_run_style(
        run,
        font_name=font_name,
        font_size=float(block.font_size or 0.0),
        bold=bold,
        italic=italic,
        color=block.color,
    )


def _add_text_block(
    document: Any,
    block: IRBlock,
    *,
    font_plan: FontPlan | None = None,
) -> Any:
    level = _HEADING_LEVELS.get(block.kind)
    if level is not None:
        paragraph = document.add_heading(level=level)
        _configure_paragraph(paragraph, block)
        _add_text_runs(paragraph, block, font_plan=font_plan)
        return paragraph
    if block.kind == "ordered":
        paragraph = document.add_paragraph(style="List Number")
    elif block.kind in {"bullet", "plugin"}:
        paragraph = document.add_paragraph(style="List Bullet")
    elif block.kind == "code":
        paragraph = document.add_paragraph(style="Code Block")
    else:
        paragraph = document.add_paragraph()
    _configure_paragraph(paragraph, block)
    _add_text_runs(paragraph, block, font_plan=font_plan)
    return paragraph


def _image_size_points(
    block: IRBlock,
    page: IRPage,
    *,
    margin_points: float,
) -> tuple[float, float]:
    if block.bbox is not None:
        width = max(float(block.bbox[2]) - float(block.bbox[0]), 1.0)
        height = max(float(block.bbox[3]) - float(block.bbox[1]), 1.0)
    else:
        width = max(float(block.image_width or 0.0), 1.0)
        height = max(float(block.image_height or 0.0), 1.0)
    max_width = max(page.width - margin_points * 2.0, 24.0)
    max_height = max(page.height - margin_points * 2.0, 24.0)
    scale = min(1.0, max_width / width, max_height / height)
    return width * scale, height * scale


def _image_alignment(block: IRBlock, page: IRPage) -> int:
    if block.bbox is None:
        return WD_ALIGN_PARAGRAPH.CENTER
    center = (float(block.bbox[0]) + float(block.bbox[2])) / 2.0
    if abs(center - page.width / 2.0) <= max(page.width * 0.12, 24.0):
        return WD_ALIGN_PARAGRAPH.CENTER
    if center > page.width * 0.58:
        return WD_ALIGN_PARAGRAPH.RIGHT
    return WD_ALIGN_PARAGRAPH.LEFT


def _add_inline_image(
    container: Any,
    block: IRBlock,
    page: IRPage,
    *,
    margin_points: float,
    paragraph: Any | None = None,
) -> Any | None:
    if not block.image_bytes:
        return None
    created_paragraph = paragraph is None
    target = paragraph or container.add_paragraph()
    target.alignment = _image_alignment(block, page)
    _configure_paragraph(target)
    width, height = _image_size_points(
        block,
        page,
        margin_points=margin_points,
    )
    try:
        inline = target.add_run().add_picture(
            BytesIO(block.image_bytes),
            width=Inches(width / 72.0),
            height=Inches(height / 72.0),
        )
    except Exception:
        if created_paragraph:
            element = target._element
            parent = element.getparent()
            if parent is not None:
                parent.remove(element)
        return None
    description = str(block.image_alt or block.source or "PDF 图片")
    inline._inline.docPr.set("descr", description)
    return target


def _render_block_region(
    source_pdf: Path | None,
    page: IRPage,
    block: IRBlock,
    *,
    max_pixels: int,
) -> bytes | None:
    if block.fallback_image:
        return block.fallback_image
    if source_pdf is None or block.bbox is None or not source_pdf.is_file():
        return None
    try:
        return render_pdf_region(
            source_pdf,
            page.page_number - 1,
            block.bbox,
            max_pixels=max_pixels,
            jpeg_quality=88,
            scale=3.0,
        )
    except Exception:
        return None


def _vector_union_bbox(page: IRPage) -> tuple[float, float, float, float] | None:
    boxes = [
        block.bbox
        for block in page.body_blocks
        if block.kind == "vector" and block.bbox is not None
    ]
    if not boxes:
        return None
    return (
        max(min(float(box[0]) for box in boxes), 0.0),
        max(min(float(box[1]) for box in boxes), 0.0),
        min(max(float(box[2]) for box in boxes), page.width),
        min(max(float(box[3]) for box in boxes), page.height),
    )


def _add_formula_block(
    document: Any,
    page: IRPage,
    block: IRBlock,
    *,
    source_pdf: Path | None,
    margin_points: float,
    max_fallback_pixels: int,
    report: dict[str, Any],
) -> Any:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _configure_paragraph(paragraph)
    formula_text = str(block.text or block.latex or "").strip()
    if formula_text and _add_formula_omml_element(paragraph, formula_text):
        report["formula_omml_count"] += 1
        return paragraph

    image_bytes = _render_block_region(
        source_pdf,
        page,
        block,
        max_pixels=max_fallback_pixels,
    )
    if image_bytes:
        image_block = IRBlock(
            kind="image",
            page=block.page,
            bbox=block.bbox,
            image_bytes=image_bytes,
            image_width=block.width,
            image_height=block.height,
            image_alt="公式图片兜底",
            source="formula_fallback",
        )
        image_paragraph = _add_inline_image(
            document,
            image_block,
            page,
            margin_points=margin_points,
            paragraph=paragraph,
        )
        if image_paragraph is not None:
            report["formula_image_fallback_count"] += 1
            return paragraph

    run = paragraph.add_run(formula_text)
    _set_run_style(
        run,
        font_name="Cambria Math",
        font_size=float(block.font_size or 10.0),
        italic=True,
    )
    report["formula_text_fallback_count"] += 1
    return paragraph


def _add_html_table(document: Any, value: str) -> Any | None:
    rows = _html_table_rows(value)
    if not rows:
        return None

    placements: list[tuple[int, int, _HtmlTableCell]] = []
    occupied: set[tuple[int, int]] = set()
    column_count = 0
    for row_index, values in enumerate(rows):
        column_index = 0
        for cell in values:
            while any(
                (row_index + row_offset, column_index + column_offset)
                in occupied
                for row_offset in range(cell.row_span)
                for column_offset in range(cell.column_span)
            ):
                column_index += 1
            placements.append((row_index, column_index, cell))
            for row_offset in range(cell.row_span):
                for column_offset in range(cell.column_span):
                    occupied.add(
                        (row_index + row_offset, column_index + column_offset)
                    )
            column_index += cell.column_span
            column_count = max(column_count, column_index)
    if column_count < 1:
        return None

    table = document.add_table(rows=len(rows), cols=column_count)
    table.style = "Table Grid"
    table.autofit = True
    for row_index, column_index, source_cell in placements:
        target_cell = table.cell(row_index, column_index)
        if source_cell.row_span > 1 or source_cell.column_span > 1:
            target_cell = target_cell.merge(
                table.cell(
                    row_index + source_cell.row_span - 1,
                    column_index + source_cell.column_span - 1,
                )
            )
        target_cell.text = source_cell.text
        for paragraph in target_cell.paragraphs:
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            if source_cell.header:
                for run in paragraph.runs:
                    run.bold = True
    return table


def _remove_tables_added_after(document: Any, table_count: int) -> None:
    for table in list(document.tables)[table_count:]:
        element = table._element
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)


def _add_table_block(
    document: Any,
    page: IRPage,
    block: IRBlock,
    *,
    source_pdf: Path | None,
    margin_points: float,
    max_fallback_pixels: int,
    report: dict[str, Any],
) -> Any | None:
    table = None
    existing_table_count = len(document.tables)
    try:
        if block.kind == "table" and block.table is not None:
            table = _add_pdf_table(
                document,
                block.table,
                bordered=bool(getattr(block.table, "has_borders", True)),
                exact_widths=False,
                cell_vertical_alignment="top",
            )
        elif block.kind == "html_table":
            table = _add_html_table(document, block.text)
    except Exception as error:
        _remove_tables_added_after(document, existing_table_count)
        report["table_native_error_count"] += 1
        report["table_native_error_types"].append(type(error).__name__)
    if table is not None:
        report["table_count"] += 1
        if block.kind == "html_table":
            report["html_table_count"] += 1
        return table

    image_bytes = _render_block_region(
        source_pdf,
        page,
        block,
        max_pixels=max_fallback_pixels,
    )
    if image_bytes:
        image_block = IRBlock(
            kind="image",
            page=block.page,
            bbox=block.bbox,
            image_bytes=image_bytes,
            image_width=block.width,
            image_height=block.height,
            image_alt="表格图片兜底",
            source="table_fallback",
        )
        image_paragraph = _add_inline_image(
            document,
            image_block,
            page,
            margin_points=margin_points,
        )
        if image_paragraph is not None:
            report["table_image_fallback_count"] += 1
            return None
    report["skipped_block_count"] += 1
    return None


def _add_flow_header_footer(
    section: Any,
    page: IRPage,
    *,
    font_plan: FontPlan | None,
    report: dict[str, Any],
) -> None:
    for layer, blocks in (
        ("header", page.header_blocks),
        ("footer", page.footer_blocks),
    ):
        if not blocks:
            continue
        container = section.header if layer == "header" else section.footer
        container.is_linked_to_previous = False
        paragraphs = list(container.paragraphs)
        first = paragraphs[0] if paragraphs else container.add_paragraph()
        first.text = ""
        for index, block in enumerate(sorted(blocks, key=lambda item: item.z_order)):
            paragraph = first if index == 0 else container.add_paragraph()
            if block.kind in FLOW_TEXT_KINDS:
                _configure_paragraph(paragraph, block)
                _add_text_runs(paragraph, block, font_plan=font_plan)
                report["header_footer_count"] += 1
            elif block.kind == "image":
                if _add_inline_image(
                    container,
                    block,
                    page,
                    margin_points=FLOW_MARGIN_POINTS,
                    paragraph=paragraph,
                ):
                    report["header_footer_count"] += 1


def _record_page_report(
    report: dict[str, Any],
    page: IRPage,
    *,
    start_index: int,
) -> None:
    placements = report["placements"][start_index:]
    report["pages"].append(
        {
            "page": page.page_number,
            "route": page.route,
            "text_paragraph_count": sum(
                item["kind"] in FLOW_TEXT_KINDS and item["status"] == "native"
                for item in placements
            ),
            "table_count": sum(
                item["kind"] in {"table", "html_table"}
                and item["status"] == "native"
                for item in placements
            ),
            "image_count": sum(
                item["kind"] in {"image", "page_image"}
                and item["status"] == "native"
                for item in placements
            ),
            "placements": placements,
        }
    )


def _record_placement(
    report: dict[str, Any],
    block: IRBlock,
    *,
    status: str,
    reason: str = "",
) -> None:
    report["placements"].append(
        {
            "kind": block.kind,
            "page": block.page,
            "status": status,
            "reason": reason,
        }
    )


def _add_page_block(
    document: Any,
    page: IRPage,
    block: IRBlock,
    *,
    source_pdf: Path | None,
    font_plan: FontPlan | None,
    margin_points: float,
    max_fallback_pixels: int,
    report: dict[str, Any],
) -> Any | None:
    if block.kind in FLOW_TEXT_KINDS:
        paragraph = _add_text_block(document, block, font_plan=font_plan)
        report["text_paragraph_count"] += 1
        _record_placement(report, block, status="native", reason="flow_paragraph")
        return paragraph
    if block.kind in {"table", "html_table"}:
        fallback_count = report["table_image_fallback_count"]
        table = _add_table_block(
            document,
            page,
            block,
            source_pdf=source_pdf,
            margin_points=margin_points,
            max_fallback_pixels=max_fallback_pixels,
            report=report,
        )
        _record_placement(
            report,
            block,
            status=(
                "native"
                if table is not None
                else (
                    "image_fallback"
                    if report["table_image_fallback_count"] > fallback_count
                    else "skipped"
                )
            ),
            reason=(
                "flow_table"
                if table is not None
                else (
                    "table_image_fallback"
                    if report["table_image_fallback_count"] > fallback_count
                    else "table_unavailable"
                )
            ),
        )
        return None
    if block.kind == "formula":
        omml_count = report["formula_omml_count"]
        image_fallback_count = report["formula_image_fallback_count"]
        text_fallback_count = report["formula_text_fallback_count"]
        paragraph = _add_formula_block(
            document,
            page,
            block,
            source_pdf=source_pdf,
            margin_points=margin_points,
            max_fallback_pixels=max_fallback_pixels,
            report=report,
        )
        report["formula_count"] += 1
        _record_placement(
            report,
            block,
            status=(
                "native"
                if report["formula_omml_count"] > omml_count
                else (
                    "image_fallback"
                    if report["formula_image_fallback_count"] > image_fallback_count
                    else "text_fallback"
                )
            ),
            reason=(
                "flow_formula_omml"
                if report["formula_omml_count"] > omml_count
                else (
                    "formula_image_fallback"
                    if report["formula_image_fallback_count"] > image_fallback_count
                    else "formula_text_fallback"
                )
            ),
        )
        return paragraph
    if block.kind in {"image", "page_image"}:
        paragraph = _add_inline_image(
            document,
            block,
            page,
            margin_points=margin_points,
        )
        if paragraph is None:
            report["skipped_block_count"] += 1
            _record_placement(report, block, status="skipped", reason="image_data_missing")
            return None
        report["image_count"] += 1
        if block.kind == "page_image":
            report["page_image_count"] += 1
        _record_placement(report, block, status="native", reason="inline_picture")
        return paragraph
    if block.kind == "vector":
        image_bytes = _render_block_region(
            source_pdf,
            page,
            block,
            max_pixels=max_fallback_pixels,
        )
        if image_bytes:
            image_block = IRBlock(
                kind="image",
                page=block.page,
                bbox=block.bbox,
                image_bytes=image_bytes,
                image_width=block.width,
                image_height=block.height,
                image_alt="PDF 矢量区域",
                source="vector_fallback",
            )
            image_paragraph = _add_inline_image(
                document,
                image_block,
                page,
                margin_points=margin_points,
            )
            if image_paragraph is not None:
                report["vector_image_fallback_count"] += 1
                _record_placement(
                    report,
                    block,
                    status="image_fallback",
                    reason="vector_rasterized",
                )
                return None
        report["skipped_block_count"] += 1
        _record_placement(report, block, status="skipped", reason="vector_unavailable")
        return None
    report["skipped_block_count"] += 1
    _record_placement(report, block, status="skipped", reason="unsupported_kind")
    return None


def export_structured_docx(
    ir: IRDocument,
    output_path: Path,
    *,
    title: str | None = None,
    source_pdf: Path | None = None,
    include_toc: bool = False,
    include_bookmarks: bool = True,
    margin_points: float = FLOW_MARGIN_POINTS,
    max_fallback_pixels: int = 2_000_000,
    font_plan: FontPlan | None = None,
    stage_callback: StageCallback | None = None,
) -> dict[str, Any]:
    """按 IR 的逻辑块顺序生成可编辑、可流式重排的 DOCX。"""
    if margin_points < 0:
        raise ValueError("流式导出页边距不能为负数")
    if max_fallback_pixels < 1:
        raise ValueError("结构化导出图片像素上限必须大于 0")
    if source_pdf is not None:
        source_pdf = Path(source_pdf)

    document = Document()
    _set_document_styles(document)
    document.core_properties.title = title or ir.title
    _remove_initial_empty_paragraph(document)
    if include_toc:
        document.add_heading("目录", level=1)
        _add_toc_field(document)
        document.add_page_break()

    outline_by_page: dict[int, list[dict[str, Any]]] = {}
    if include_bookmarks:
        for entry in ir.metadata.get("outline", []):
            page_number = entry.get("page")
            if isinstance(page_number, int):
                outline_by_page.setdefault(page_number, []).append(entry)

    report: dict[str, Any] = {
        "mode": "structured",
        "page_count": len(ir.pages),
        "text_paragraph_count": 0,
        "table_count": 0,
        "html_table_count": 0,
        "formula_count": 0,
        "formula_omml_count": 0,
        "formula_image_fallback_count": 0,
        "formula_text_fallback_count": 0,
        "image_count": 0,
        "page_image_count": 0,
        "vector_image_fallback_count": 0,
        "vector_compacted_page_count": 0,
        "vector_skipped_count": 0,
        "table_image_fallback_count": 0,
        "table_native_error_count": 0,
        "table_native_error_types": [],
        "header_footer_count": 0,
        "skipped_block_count": 0,
        "placements": [],
        "pages": [],
    }
    _notify_stage(stage_callback, "structured_export_started", page_count=len(ir.pages))
    bookmark_id = 0

    for page_index, page in enumerate(ir.pages):
        section = (
            document.sections[0]
            if page_index == 0
            else document.add_section(WD_SECTION.NEW_PAGE)
        )
        _set_section_page(
            section,
            max(page.width / 72.0, 0.01),
            max(page.height / 72.0, 0.01),
            margin_points / 72.0,
        )
        if page.header_blocks or page.footer_blocks:
            _add_flow_header_footer(
                section,
                page,
                font_plan=font_plan,
                report=report,
            )
        else:
            detach_header_footer(section)

        start_index = len(report["placements"])
        first_paragraph = None
        vector_blocks = [
            block for block in page.body_blocks if block.kind == "vector"
        ]
        dense_vector_bbox = (
            _vector_union_bbox(page)
            if len(vector_blocks) > DENSE_VECTOR_LIMIT
            else None
        )
        dense_vector_image = None
        if dense_vector_bbox is not None:
            dense_vector_block = IRBlock(
                kind="vector",
                page=page.page_number,
                bbox=dense_vector_bbox,
                source="dense_vector_region",
            )
            dense_vector_image = _render_block_region(
                source_pdf,
                page,
                dense_vector_block,
                max_pixels=max_fallback_pixels,
            )
            report["vector_compacted_page_count"] += 1
        dense_vector_inserted = False
        for block in page.body_blocks:
            if dense_vector_bbox is not None and block.kind == "vector":
                if dense_vector_image is not None and not dense_vector_inserted:
                    image_block = IRBlock(
                        kind="image",
                        page=block.page,
                        bbox=dense_vector_bbox,
                        image_bytes=dense_vector_image,
                        image_width=dense_vector_bbox[2] - dense_vector_bbox[0],
                        image_height=dense_vector_bbox[3] - dense_vector_bbox[1],
                        image_alt="密集矢量区域",
                        source="dense_vector_fallback",
                    )
                    paragraph = _add_inline_image(
                        document,
                        image_block,
                        page,
                        margin_points=margin_points,
                    )
                    if paragraph is not None:
                        dense_vector_inserted = True
                        report["vector_image_fallback_count"] += 1
                        _record_placement(
                            report,
                            block,
                            status="image_fallback",
                            reason="dense_vector_layer_rasterized",
                        )
                        if first_paragraph is None:
                            first_paragraph = paragraph
                        continue
                report["vector_skipped_count"] += 1
                report["skipped_block_count"] += 1
                _record_placement(
                    report,
                    block,
                    status="skipped",
                    reason=(
                        "dense_vector_layer_compacted"
                        if dense_vector_image is not None
                        else "dense_vector_layer_unavailable"
                    ),
                )
                continue
            paragraph = _add_page_block(
                document,
                page,
                block,
                source_pdf=source_pdf,
                font_plan=font_plan,
                margin_points=margin_points,
                max_fallback_pixels=max_fallback_pixels,
                report=report,
            )
            if first_paragraph is None and paragraph is not None:
                first_paragraph = paragraph

        if first_paragraph is None:
            first_paragraph = document.add_paragraph()
            _configure_paragraph(first_paragraph)
        if include_bookmarks:
            for entry in outline_by_page.get(page.page_number, []):
                bookmark_id += 1
                _add_bookmark(
                    first_paragraph,
                    name=_bookmark_name(
                        str(entry.get("title") or ""),
                        bookmark_id,
                    ),
                    bookmark_id=bookmark_id,
                )
        _record_page_report(report, page, start_index=start_index)
        _notify_stage(
            stage_callback,
            "structured_page_completed",
            page=page.page_number,
            page_count=len(ir.pages),
        )

    if font_plan is not None and font_plan.programs:
        report["embedded_font_count"] = len(font_plan.programs)
        attach_embedded_fonts(document, font_plan.programs)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    _notify_stage(
        stage_callback,
        "structured_export_completed",
        page_count=len(ir.pages),
        text_paragraph_count=report["text_paragraph_count"],
        table_count=report["table_count"],
        formula_count=report["formula_count"],
        image_count=report["image_count"],
    )
    return report


__all__ = ["STRUCTURED_MODES", "export_structured_docx"]
