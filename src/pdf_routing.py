from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path

import pypdfium2 as pdfium

_URL_PATTERN = re.compile(r"https?://\S+", re.IGNORECASE)
_SYMBOL_FONT_CHAR_MAPPING = str.maketrans(
    {
        # Wingdings 私有区项目符号，字形为实心圆点与空心方块
        "\uf06c": "•",
        "\uf071": "❑",
        "\uf0b7": "•",
    }
)
_REPLACEMENT_CHARACTERS = frozenset({"\ufffd", "\u0000"})
_PAGE_IMAGE_OBJECT_TYPE = 3


@dataclass(frozen=True)
class PdfPageTextAnalysis:
    """保存单页文本层质量和页面图像信号。"""

    text: str
    text_char_count: int
    garbled_char_count: int
    url_count: int
    image_count: int
    full_page_image_count: int
    max_image_pixels: int
    quality_score: float

    @property
    def garbled_char_ratio(self) -> float:
        return self.garbled_char_count / max(self.text_char_count, 1)

    @property
    def url_char_ratio(self) -> float:
        url_char_count = sum(len(match.group(0)) for match in _URL_PATTERN.finditer(self.text))
        return url_char_count / max(self.text_char_count, 1)

    def is_high_quality(self, *, min_page_chars: int) -> bool:
        return (
            self.text_char_count >= min_page_chars
            and self.quality_score >= 0.75
        )


@dataclass(frozen=True)
class PdfTextAnalysis:
    """保存 PDF 文本层检测结果和逐页提取文本。"""

    page_texts: list[str]
    pages: list[PdfPageTextAnalysis]
    usable_page_count: int
    high_quality_page_count: int
    text_char_count: int
    full_page_image_page_count: int
    garbled_char_count: int

    @property
    def page_count(self) -> int:
        return len(self.page_texts)

    @property
    def usable_page_ratio(self) -> float:
        return self.usable_page_count / max(self.page_count, 1)

    @property
    def high_quality_page_ratio(self) -> float:
        return self.high_quality_page_count / max(self.page_count, 1)

    def has_usable_text_layer(
        self,
        *,
        min_page_chars: int,
        min_page_ratio: float,
    ) -> bool:
        return (
            self.usable_page_count > 0
            and self.text_char_count >= min_page_chars
            and self.usable_page_ratio >= min_page_ratio
        )

    def has_complete_text_layer(
        self,
        *,
        min_page_chars: int,
        min_high_quality_ratio: float,
    ) -> bool:
        high_quality_page_count = sum(
            page.is_high_quality(min_page_chars=min_page_chars) for page in self.pages
        )
        return (
            self.page_count > 0
            and high_quality_page_count > 0
            and high_quality_page_count / max(self.page_count, 1)
            >= min_high_quality_ratio
        )


def normalize_page_text(value: str | None) -> str:
    """归一化 PDF 字体字符，清理空字符和行内多余空格并保留换行。"""
    if not value:
        return ""
    value = "".join(
        unicodedata.normalize("NFKC", character)
        if "MATHEMATICAL" in unicodedata.name(character, "")
        else character
        for character in value
    ).translate(_SYMBOL_FONT_CHAR_MAPPING)
    lines = []
    for line in value.replace("\x00", "").splitlines():
        normalized = re.sub(r"[ \t]+", " ", line).strip()
        if normalized:
            lines.append(normalized)
    return "\n".join(lines)


def _read_page_text(value: str | None) -> str:
    """归一化页面文本用于文本层判定，兼容字符按 Unicode 标准映射折叠。"""
    if not value:
        return ""
    return normalize_page_text(unicodedata.normalize("NFKC", value))


def count_text_characters(value: str) -> int:
    """统计去除空白后的文本字符数量。"""
    return len(re.sub(r"\s+", "", value))


def _is_suspicious_character(character: str) -> bool:
    """判断字符是否为无法确认语义的异常字形。"""
    return (
        character in _REPLACEMENT_CHARACTERS
        or unicodedata.category(character) == "Co"
    )


def _count_garbled_characters(value: str) -> int:
    return sum(_is_suspicious_character(character) for character in value)


def _page_image_sizes(page: object) -> list[tuple[int, int]]:
    """返回页面内所有图像对象的固有像素尺寸。"""
    sizes = []
    try:
        objects = list(page.get_objects())
    except Exception:
        return sizes
    for item in objects:
        if item.type != _PAGE_IMAGE_OBJECT_TYPE:
            continue
        try:
            width, height = item.get_px_size()
        except Exception:
            continue
        sizes.append((int(width), int(height)))
    return sizes


def _is_full_page_image(
    width: int,
    height: int,
    *,
    page_width: float,
    page_height: float,
    min_pixels: int,
) -> bool:
    if width * height < min_pixels:
        return False
    page_ratio = page_width / max(page_height, 0.01)
    image_ratio = width / max(height, 1)
    return abs(image_ratio - page_ratio) / max(page_ratio, 0.01) <= 0.03


def _analyze_page(
    page: object,
    *,
    min_page_chars: int,
    full_page_image_min_pixels: int,
    garbled_char_ratio_threshold: float,
) -> PdfPageTextAnalysis:
    text = _read_page_text(page.get_textpage().get_text_range())
    text_char_count = count_text_characters(text)
    garbled_char_count = _count_garbled_characters(text)
    url_count = len(_URL_PATTERN.findall(text))
    page_width, page_height = page.get_size()
    image_sizes = _page_image_sizes(page)
    image_count = len(image_sizes)
    full_page_image_count = 0
    max_image_pixels = 0
    for width, height in image_sizes:
        max_image_pixels = max(max_image_pixels, width * height)
        if _is_full_page_image(
            width,
            height,
            page_width=float(page_width),
            page_height=float(page_height),
            min_pixels=full_page_image_min_pixels,
        ):
            full_page_image_count += 1

    quality_score = 1.0
    if text_char_count < min_page_chars:
        quality_score -= 0.55
    if full_page_image_count:
        quality_score -= 0.45
    if garbled_char_count:
        garbled_char_ratio = garbled_char_count / max(text_char_count, 1)
        quality_score -= 0.35 * min(1.0, garbled_char_ratio / garbled_char_ratio_threshold)
    url_char_count = sum(len(match.group(0)) for match in _URL_PATTERN.finditer(text))
    if url_char_count / max(text_char_count, 1) > 0.25:
        quality_score -= 0.15
    if image_count and text_char_count < max(min_page_chars * 5, 100):
        quality_score -= 0.15

    return PdfPageTextAnalysis(
        text=text,
        text_char_count=text_char_count,
        garbled_char_count=garbled_char_count,
        url_count=url_count,
        image_count=image_count,
        full_page_image_count=full_page_image_count,
        max_image_pixels=max_image_pixels,
        quality_score=round(max(0.0, min(1.0, quality_score)), 3),
    )


def analyze_pdf_text(
    path: Path,
    *,
    min_page_chars: int,
    full_page_image_min_pixels: int = 300_000,
    garbled_char_ratio_threshold: float = 0.05,
) -> PdfTextAnalysis:
    """提取各页文本并评估文本层完整性与页面图像信号。"""
    if min_page_chars < 1:
        raise ValueError("文本层最小字符数必须大于 0")
    if full_page_image_min_pixels < 1:
        raise ValueError("全页图像最小像素数必须大于 0")
    if not 0 <= garbled_char_ratio_threshold <= 1:
        raise ValueError("异常字形比例阈值必须在 0 到 1 之间")

    document = pdfium.PdfDocument(str(path))
    page_texts: list[str] = []
    pages: list[PdfPageTextAnalysis] = []
    usable_page_count = 0
    high_quality_page_count = 0
    text_char_count = 0
    full_page_image_page_count = 0
    garbled_char_count = 0
    try:
        for page in document:
            page_analysis = _analyze_page(
                page,
                min_page_chars=min_page_chars,
                full_page_image_min_pixels=full_page_image_min_pixels,
                garbled_char_ratio_threshold=garbled_char_ratio_threshold,
            )
            pages.append(page_analysis)
            page_texts.append(page_analysis.text)
            text_char_count += page_analysis.text_char_count
            garbled_char_count += page_analysis.garbled_char_count
            if page_analysis.text_char_count >= min_page_chars:
                usable_page_count += 1
            if page_analysis.is_high_quality(min_page_chars=min_page_chars):
                high_quality_page_count += 1
            if page_analysis.full_page_image_count:
                full_page_image_page_count += 1
    finally:
        document.close()

    return PdfTextAnalysis(
        page_texts=page_texts,
        pages=pages,
        usable_page_count=usable_page_count,
        high_quality_page_count=high_quality_page_count,
        text_char_count=text_char_count,
        full_page_image_page_count=full_page_image_page_count,
        garbled_char_count=garbled_char_count,
    )
