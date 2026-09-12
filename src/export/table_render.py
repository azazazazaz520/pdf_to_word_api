"""把版面提取出的 PdfTable 渲染为 Word 固定布局表格。

表格边框、列宽、行高与单元格内边距在此处对齐源 PDF；三条导出路径共用。
"""


from typing import Any
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from ..layout.models import PdfTable


def _set_table_row_properties(row: Any, *, repeat_header: bool) -> None:
    row_properties = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    row_properties.append(cant_split)
    if repeat_header:
        table_header = OxmlElement("w:tblHeader")
        table_header.set(qn("w:val"), "true")
        row_properties.append(table_header)


def _set_table_cell_margins(
    table: Any,
    *,
    tight: bool = False,
    padding_points: tuple[float, float] | None = None,
) -> None:
    """把 Word 表格单元格上下左右边距压缩到接近源 PDF。"""
    table_properties = table._tbl.tblPr
    for existing in table_properties.findall(qn("w:tblCellMar")):
        table_properties.remove(existing)
    cell_margins = OxmlElement("w:tblCellMar")
    if padding_points is not None:
        left = max(int(round(float(padding_points[0]) * 20)), 0)
        top = max(int(round(float(padding_points[1]) * 20)), 0)
        values = (
            ("w:top", top),
            ("w:left", left),
            ("w:bottom", 0),
            ("w:right", left),
        )
    else:
        values = (
            (("w:top", 0), ("w:left", 10), ("w:bottom", 0), ("w:right", 10))
            if tight
            else (("w:top", 20), ("w:left", 40), ("w:bottom", 20), ("w:right", 40))
        )
    for tag, value in values:
        element = OxmlElement(tag)
        element.set(qn("w:w"), str(value))
        element.set(qn("w:type"), "dxa")
        cell_margins.append(element)
    table_properties.append(cell_margins)


def _set_fixed_table_layout(table: Any) -> None:
    """将 Word 表格设置为固定布局，避免自动调整破坏 PDF 列宽比例。"""
    table_properties = table._tbl.tblPr
    layout = table_properties.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        table_properties.append(layout)
    layout.set(qn("w:type"), "fixed")


def _set_pdf_cell_content(
    cell: Any,
    text: str,
    *,
    width: float,
    is_header: bool,
    font_size: float = 0.0,
    bold: bool = False,
    alignment: str = "",
    vertical_alignment: str = "center",
) -> None:
    """设置 PDF 表格单元格的宽度、对齐方式和文本格式。"""
    cell.width = Inches(width)
    cell.vertical_alignment = (
        WD_CELL_VERTICAL_ALIGNMENT.TOP
        if vertical_alignment == "top"
        else WD_CELL_VERTICAL_ALIGNMENT.CENTER
    )
    cell.text = text
    alignments = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.space_before = Pt(0)
        if alignment in alignments:
            paragraph.alignment = alignments[alignment]
        if is_header:
            for run in paragraph.runs:
                run.bold = True
        if font_size and font_size > 0:
            for run in paragraph.runs:
                run.font.size = Pt(font_size)
        if bold:
            for run in paragraph.runs:
                run.bold = True


def _cell_lookup(
    pdf_table: PdfTable,
    row_index: int,
    column_index: int,
) -> Any:
    for cell in pdf_table.cells:
        if (
            cell.row_index == row_index
            and cell.column_index == column_index
        ):
            return cell
    return None


def _cell_font_size(
    pdf_table: PdfTable,
    row_index: int,
    column_index: int,
) -> float:
    cell = _cell_lookup(pdf_table, row_index, column_index)
    return float(getattr(cell, "font_size", 0.0) or 0.0)


def _cell_alignment(
    pdf_table: PdfTable,
    row_index: int,
    column_index: int,
) -> str:
    cell = _cell_lookup(pdf_table, row_index, column_index)
    return str(getattr(cell, "alignment", "") or "")


def _add_pdf_table(
    document: Document,
    pdf_table: PdfTable,
    *,
    bordered: bool = True,
    exact_widths: bool = False,
    cell_vertical_alignment: str = "center",
    tight_cell_margins: bool = False,
    cell_padding_points: tuple[float, float] | None = None,
) -> Any:
    if not pdf_table.rows or pdf_table.column_count < 1:
        return None
    table = document.add_table(rows=0, cols=pdf_table.column_count)
    if bordered:
        table.style = "Table Grid"
    table.autofit = False
    _set_fixed_table_layout(table)
    _set_table_cell_margins(
        table,
        tight=tight_cell_margins,
        padding_points=cell_padding_points,
    )
    widths = [
        max(
            (
                pdf_table.column_boundaries[index + 1]
                - pdf_table.column_boundaries[index]
            )
            / 72.0,
            0.02 if exact_widths else 0.2,
        )
        for index in range(pdf_table.column_count)
    ]
    if exact_widths:
        scale = 1.0
    else:
        section = document.sections[-1]
        available_width = max(
            section.page_width.inches
            - section.left_margin.inches
            - section.right_margin.inches,
            0.2,
        )
        scale = min(1.0, available_width / max(sum(widths), 0.2))
    widths = [width * scale for width in widths]
    continuation_header_rows = list(pdf_table.continuation_header_rows)
    source_rows = list(pdf_table.rows)
    row_offset = len(continuation_header_rows)
    all_rows = continuation_header_rows + source_rows
    header_row_count = row_offset + pdf_table.header_row_count
    row_heights = pdf_table.row_heights
    default_header_height = row_heights[0] if row_heights else 18.0
    all_row_heights = [
        *([default_header_height] * row_offset),
        *row_heights,
    ]
    # Word 会把单元格上内边距计入行高，这里换算成纯文本行高。
    top_padding = (
        max(float(cell_padding_points[1]), 0.0)
        if cell_padding_points
        else 0.0
    )
    all_row_heights = [
        max(height - top_padding, 6.0) for height in all_row_heights
    ]
    for row_index in range(len(all_rows)):
        row = table.add_row()
        _set_table_row_properties(
            row,
            repeat_header=row_index < header_row_count,
        )
        if row_index < len(all_row_heights):
            row.height = Pt(max(all_row_heights[row_index] * scale, 6.0))
            row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        for column_index, cell in enumerate(row.cells):
            cell.width = Inches(widths[column_index])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    for row_index, values in enumerate(continuation_header_rows):
        for column_index, cell in enumerate(table.rows[row_index].cells):
            _set_pdf_cell_content(
                cell,
                values[column_index] if column_index < len(values) else "",
                width=widths[column_index],
                is_header=True,
                font_size=_cell_font_size(pdf_table, row_index, column_index),
                alignment=_cell_alignment(pdf_table, row_index, column_index),
                vertical_alignment=cell_vertical_alignment,
            )

    if pdf_table.cells:
        for pdf_cell in pdf_table.cells:
            source_row_index = pdf_cell.row_index + row_offset
            cell = table.cell(source_row_index, pdf_cell.column_index)
            if pdf_cell.row_span > 1 or pdf_cell.column_span > 1:
                cell = cell.merge(
                    table.cell(
                        source_row_index + pdf_cell.row_span - 1,
                        pdf_cell.column_index + pdf_cell.column_span - 1,
                    )
                )
            _set_pdf_cell_content(
                cell,
                pdf_cell.text,
                width=sum(
                    widths[
                        pdf_cell.column_index : pdf_cell.column_index
                        + pdf_cell.column_span
                    ]
                ),
                is_header=pdf_cell.row_index < pdf_table.header_row_count,
                font_size=pdf_cell.font_size,
                bold=pdf_cell.bold,
                alignment=pdf_cell.alignment,
                vertical_alignment=cell_vertical_alignment,
            )
    else:
        for row_index, values in enumerate(all_rows):
            for column_index, cell in enumerate(table.rows[row_index].cells):
                _set_pdf_cell_content(
                    cell,
                    values[column_index] if column_index < len(values) else "",
                    width=widths[column_index],
                    is_header=row_index < header_row_count,
                    font_size=_cell_font_size(pdf_table, row_index, column_index),
                    alignment=_cell_alignment(pdf_table, row_index, column_index),
                    vertical_alignment=cell_vertical_alignment,
                )
    for column_index, column in enumerate(table.columns):
        column.width = Inches(widths[column_index])
    return table


