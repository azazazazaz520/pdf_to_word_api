"""版面提取的数据契约：页面、文本行、表格、矢量与图片的表示。

本模块不依赖同包其他模块，是拆包后的共同依赖层。
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any
from collections import Counter
from ..pdf_routing import normalize_page_text

_PAGEOBJ_PATH = 2


_PAGEOBJ_TEXT = 1


_PAGEOBJ_IMAGE = 3


_MAX_LINE_THICKNESS = 2.0


_MIN_HORIZONTAL_LINE_LENGTH = 20.0


_MIN_VERTICAL_LINE_LENGTH = 10.0


_COORDINATE_TOLERANCE = 2.5


_MAX_TABLE_ROW_GAP = 60.0


_MIN_TABLE_WIDTH = 60.0


_MIN_CELL_LINE_LENGTH = 8.0


_MIN_BOUNDARY_COVERAGE = 0.6


@dataclass(frozen=True)
class PdfTextGlyph:
    """一个可定位的 PDF 字符。

    ``bbox`` 使用页面左上角为原点的坐标，``direction`` 同样使用页面坐标
    （x 向右、y 向下）。字符级数据是行、段落和版面块聚合时的唯一几何来源。
    """

    text: str
    bbox: tuple[float, float, float, float]
    font_name: str = ""
    pdf_font_name: str = ""
    font_size: float = 0.0
    color: tuple[int, int, int] = (0, 0, 0)
    bold: bool = False
    italic: bool = False
    rotation: float = 0.0
    direction: tuple[float, float] = (1.0, 0.0)
    z_order: int = 0
    char_index: int = -1
    object_index: int = -1
    substituted: bool = False
    fallback_reason: str = ""

    @property
    def x0(self) -> float:
        return self.bbox[0]

    @property
    def top(self) -> float:
        return self.bbox[1]

    @property
    def x1(self) -> float:
        return self.bbox[2]

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2.0

    @property
    def center_y(self) -> float:
        return (self.top + self.bottom) / 2.0

    @property
    def width(self) -> float:
        return max(self.x1 - self.x0, 0.0)

    @property
    def height(self) -> float:
        return max(self.bottom - self.top, 0.0)

    @property
    def bottom(self) -> float:
        return self.bbox[3]


@dataclass(frozen=True)
class PdfTextSpan:
    """同一行内具有相同字体/字号/颜色的连续文本片段。"""

    text: str
    bbox: tuple[float, float, float, float]
    font_name: str = ""
    pdf_font_name: str = ""
    font_size: float = 0.0
    color: tuple[int, int, int] = (0, 0, 0)
    bold: bool = False
    italic: bool = False
    rotation: float = 0.0
    z_order: int = 0
    substituted: bool = False
    fallback_reason: str = ""
    direction: tuple[float, float] = (1.0, 0.0)
    glyphs: tuple[PdfTextGlyph, ...] = ()

    @property
    def x0(self) -> float:
        return self.bbox[0]

    @property
    def top(self) -> float:
        return self.bbox[1]

    @property
    def x1(self) -> float:
        return self.bbox[2]

    @property
    def bottom(self) -> float:
        return self.bbox[3]


@dataclass(frozen=True)
class PdfTextLine:
    """保存可用于阅读顺序和列表识别的 PDF 文本行。

    字符级字形、方向和基础字体信息始终保留；对齐、行距和 z-order 等
    高保真字段由版面提取阶段补充。
    """

    text: str
    x0: float
    top: float
    x1: float
    bottom: float
    font_size: float
    is_header_footer: bool = False
    font_name: str = ""
    color: tuple[int, int, int] = (0, 0, 0)
    bold: bool = False
    italic: bool = False
    alignment: str = "left"
    line_spacing: float = 0.0
    first_line_indent: float = 0.0
    rotation: float = 0.0
    z_order: int = 0
    layer: str = "body"
    confidence: float | None = None
    spans: tuple[PdfTextSpan, ...] = ()
    pdf_font_name: str = ""
    font_substituted: bool = False
    font_fallback_reason: str = ""
    direction: tuple[float, float] = (1.0, 0.0)
    glyphs: tuple[PdfTextGlyph, ...] = ()

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def center_y(self) -> float:
        return (self.top + self.bottom) / 2

    @property
    def width(self) -> float:
        return max(self.x1 - self.x0, 0.0)

    @property
    def height(self) -> float:
        return max(self.bottom - self.top, 0.0)


@dataclass(frozen=True)
class PdfTextBlock:
    """由几何相邻文本行组成的版面文本块。

    ``reading_order`` 只表示页面内顺序；块的类型和关系信息在建立后统一
    传给 IR，导出阶段不再从原始坐标重新猜测顺序。
    """

    lines: tuple[PdfTextLine, ...]
    bbox: tuple[float, float, float, float]
    column_index: int = -1
    direction: tuple[float, float] = (1.0, 0.0)
    rotation: float = 0.0
    is_header_footer: bool = False
    block_id: str = ""
    block_type: str = "TEXT"
    confidence: float = 1.0
    reading_order: int = -1
    parent_id: str | None = None
    region_id: str = ""
    needs_review: bool = False

    @property
    def text(self) -> str:
        return "\n".join(line.text for line in self.lines)

    @property
    def x0(self) -> float:
        return self.bbox[0]

    @property
    def top(self) -> float:
        return self.bbox[1]

    @property
    def x1(self) -> float:
        return self.bbox[2]

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2.0

    @property
    def center_y(self) -> float:
        return (self.top + self.bottom) / 2.0

    @property
    def width(self) -> float:
        return max(self.x1 - self.x0, 0.0)

    @property
    def height(self) -> float:
        return max(self.bottom - self.top, 0.0)

    @property
    def bottom(self) -> float:
        return self.bbox[3]


@dataclass(frozen=True)
class PdfImageBlock:
    """保存页面内嵌图片的位置和原始图像数据。"""

    name: str
    bbox: tuple[float, float, float, float]
    width: float
    height: float
    data: bytes
    mime_type: str = "image/png"
    source: str = "embedded"
    z_order: int = 0
    is_logo: bool = False
    layer: str = "body"


@dataclass(frozen=True)
class PdfVectorObject:
    """保存页面矢量对象（线条、矩形、路径）的几何和样式。"""

    kind: str
    bbox: tuple[float, float, float, float]
    stroke_color: tuple[int, int, int] | None = None
    fill_color: tuple[int, int, int] | None = None
    stroke_width: float = 0.0
    segments: tuple[tuple[str, float, float], ...] = ()
    closed: bool = False
    complex: bool = False
    dashed: bool = False
    filled: bool = True
    stroked: bool = True
    z_order: int = 0
    layer: str = "body"

    @property
    def width(self) -> float:
        return max(self.bbox[2] - self.bbox[0], 0.0)

    @property
    def height(self) -> float:
        return max(self.bbox[3] - self.bbox[1], 0.0)


@dataclass(frozen=True)
class PdfTableCell:
    """保存表格单元格的网格位置、跨行跨列信息和文本。"""

    row_index: int
    column_index: int
    row_span: int
    column_span: int
    bbox: tuple[float, float, float, float]
    text: str
    text_bbox: tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)
    border_widths: tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)
    border_color: tuple[int, int, int] = (0, 0, 0)
    font_size: float = 0.0
    bold: bool = False
    alignment: str = "left"


@dataclass(frozen=True)
class PdfTable:
    """保存 PDF 几何表格的边界、列边界、行边界和单元格。"""

    bbox: tuple[float, float, float, float]
    column_boundaries: tuple[float, ...]
    row_boundaries: tuple[float, ...]
    rows: tuple[tuple[str, ...], ...]
    cells: tuple[PdfTableCell, ...] = ()
    header_row_count: int = 1
    continued_from_previous_page: bool = False
    continuation_header_rows: tuple[tuple[str, ...], ...] = ()
    border_width: float = 0.0
    border_color: tuple[int, int, int] = (0, 0, 0)
    has_borders: bool = True
    z_order: int = 0

    @property
    def column_count(self) -> int:
        return max(len(self.column_boundaries) - 1, 0)

    @property
    def row_count(self) -> int:
        return max(len(self.row_boundaries) - 1, 0)

    @property
    def row_heights(self) -> tuple[float, ...]:
        """返回每一行在 PDF 中的高度，单位为 point。"""
        return tuple(
            max(self.row_boundaries[index + 1] - self.row_boundaries[index], 0.0)
            for index in range(self.row_count)
        )


@dataclass(frozen=True)
class PdfContentBlock:
    """页面中参与阅读顺序的统一内容块。"""

    block_id: str
    block_type: str
    bbox: tuple[float, float, float, float]
    text_block: PdfTextBlock | None = None
    table: PdfTable | None = None
    image: PdfImageBlock | None = None
    vector: PdfVectorObject | None = None
    source: str = ""
    confidence: float = 1.0
    reading_order: int = -1
    parent_id: str | None = None
    region_id: str = ""
    layer: str = "body"
    needs_review: bool = False

    @property
    def top(self) -> float:
        return self.bbox[1]

    @property
    def bottom(self) -> float:
        return self.bbox[3]

    @property
    def x0(self) -> float:
        return self.bbox[0]

    @property
    def x1(self) -> float:
        return self.bbox[2]

    @property
    def width(self) -> float:
        return max(self.x1 - self.x0, 0.0)

    @property
    def height(self) -> float:
        return max(self.bottom - self.top, 0.0)


@dataclass(frozen=True)
class PdfPageLayout:
    """保存单页的尺寸、文本行、几何文本块、表格和矢量对象。"""

    width: float
    height: float
    lines: tuple[PdfTextLine, ...]
    tables: tuple[PdfTable, ...]
    images: tuple[PdfImageBlock, ...] = ()
    columns: tuple[float, ...] = ()
    vectors: tuple[PdfVectorObject, ...] = ()
    has_page_background: bool = False
    text_blocks: tuple[PdfTextBlock, ...] = ()
    content_blocks: tuple[PdfContentBlock, ...] = ()
    reading_order_confidence: float = 1.0
    reading_order_warnings: tuple[str, ...] = ()


@dataclass(frozen=True)
class PdfDocumentLayout:
    """保存 PDF 全文的布局中间模型。"""

    pages: tuple[PdfPageLayout, ...]

    @property
    def page_count(self) -> int:
        return len(self.pages)

    @property
    def table_count(self) -> int:
        return sum(len(page.tables) for page in self.pages)


@dataclass(frozen=True)
class _HorizontalLine:
    top: float
    x0: float
    x1: float
    segments: tuple[tuple[float, float], ...] = ()


@dataclass(frozen=True)
class _VerticalLine:
    x: float
    top: float
    bottom: float


@dataclass(frozen=True)
class _TableCandidate:
    """保存表格候选区域及其可用于识别合并单元格的线段。"""

    bbox: tuple[float, float, float, float]
    column_boundaries: tuple[float, ...]
    row_boundaries: tuple[float, ...]
    horizontal_lines: tuple[_HorizontalLine, ...]
    vertical_lines: tuple[_VerticalLine, ...]


@dataclass
class _TextCharacter:
    text: str
    x0: float
    top: float
    x1: float
    bottom: float
    font_size: float
    is_header_footer: bool = False
    font_name: str = ""
    color: tuple[int, int, int] = (0, 0, 0)
    bold: bool = False
    italic: bool = False
    rotation: float = 0.0
    z_order: int = 0
    pdf_font_name: str = ""
    font_substituted: bool = False
    font_fallback_reason: str = ""
    direction: tuple[float, float] = (1.0, 0.0)
    char_index: int = -1
    object_index: int = -1

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def center_y(self) -> float:
        return (self.top + self.bottom) / 2


@dataclass
class _TextObjectStyle:
    """一个 PDF 文本对象的字体、颜色和层级信息。"""

    bbox: tuple[float, float, float, float]
    font_name: str
    font_size: float
    color: tuple[int, int, int]
    bold: bool
    italic: bool
    z_order: int
    rotation: float = 0.0
    pdf_font_name: str = ""
    substituted: bool = False
    fallback_reason: str = ""
    object_key: int = 0
    object_index: int = -1

    @property
    def area(self) -> float:
        return max(self.bbox[2] - self.bbox[0], 0.0) * max(
            self.bbox[3] - self.bbox[1], 0.0
        )


def _compact_text(value: str) -> str:
    return normalize_page_text(value.replace("\r", "\n")).replace("\n", " ")


def _dominant_value(values: list[Any], default: Any) -> Any:
    if not values:
        return default
    counts = Counter(values)
    return counts.most_common(1)[0][0]


