"""源文档内容的识别与写块：角色判定、行分组、可编辑段落。

流式导出与 IR 组装共用；输入为 PdfTextLine 与 IRBlock，输出为 Word 段落、
列表、代码块与表格。保真导出不使用本模块。
"""

from __future__ import annotations

import re
from collections import Counter
from collections.abc import Iterable
from typing import Any
from ..layout.models import PdfTable, PdfTextLine
from .text_utils import _FORMULA_PREFIX, _FORMULA_SYMBOL_PREFIX, _field, _looks_like_code_line


_ORDERED_ITEM = re.compile(
    r"^\s*(?:(?:\d+[.．、)](?!\d))|(?:[（(]\s*\d+\s*[)）])|(?:[①-⑳]))\s*(.+)$"
)
_BULLET_ITEM = re.compile(r"^\s*[•·●▪◦]\s*(.+)$")
_PLUGIN_ITEM = re.compile(r"^\s*(plugin_[a-zA-Z0-9_]+)\s*[:：]\s*(.+)$")
_SECTION_HEADING = re.compile(
    r"^\s*(?:[一二三四五六七八九十百]+、|第\d+章|摘要(?:\s|$)|参考文献(?:\s|（|\(|$))"
)
_SUBSECTION_HEADING = re.compile(r"^\s*\d+\.\d+(?:\s|$)")
_NUMERIC_SECTION_HEADING = re.compile(r"^\s*\d+\.(?!\d)\s+\S+")
_LIST_ROLES = frozenset({"ordered", "bullet", "plugin"})
_LINE_END_HYPHENS = frozenset(
    {"-", "\u00ad", "\u2010", "\u2011", "\u2012", "\u2013", "\u2014"}
)


def _ordered_blocks(result: Any) -> Iterable[Any]:
    blocks = list(_field(result, "parsing_res_list", []) or [])
    indexed = list(enumerate(blocks))

    def sort_key(item: tuple[int, Any]) -> tuple[int, int]:
        index, block = item
        order = _field(block, "block_order")
        if isinstance(order, (int, float)):
            return (0, int(order))
        bbox = _field(block, "block_bbox") or [0, index]
        return (1, int(bbox[1]) if len(bbox) > 1 else index)

    for _, block in sorted(indexed, key=sort_key):
        yield block


def _line_is_in_pdf_table(
    line: PdfTextLine,
    pdf_table: PdfTable,
) -> bool:
    x0, top, x1, bottom = pdf_table.bbox
    return (
        x0 - 1.0 <= line.center_x <= x1 + 1.0
        and top - 1.0 <= line.center_y <= bottom + 1.0
    )


def _typical_body_left(lines: Iterable[PdfTextLine]) -> float:
    """用出现次数最多的行首位置作为正文左边距，避免标题或表格干扰缩进判断。"""
    values = [
        round(line.x0, 1)
        for line in lines
        if not line.is_header_footer
    ]
    if not values:
        return 0.0
    counts = Counter(values)
    max_count = max(counts.values())
    candidates = [
        value for value, count in counts.items() if count == max_count
    ]
    return min(candidates)


def _is_heading_sized(line: PdfTextLine, body_font_size: float) -> bool:
    if body_font_size <= 0:
        return True
    return line.font_size >= body_font_size + 1.8


def _is_short_heading_text(value: str) -> bool:
    text = value.strip()
    if len(text) > 42:
        return False
    return not bool(re.search(r"[。；;，,：:]$", text))


def _layout_line_role(
    line: PdfTextLine,
    *,
    body_left: float,
    is_first_line: bool,
    body_font_size: float = 0.0,
) -> str:
    if _SECTION_HEADING.match(line.text):
        return "heading1"
    heading_sized = _is_heading_sized(line, body_font_size)
    if (
        _SUBSECTION_HEADING.match(line.text)
        and (heading_sized or len(line.text.strip()) <= 24)
        and _is_short_heading_text(line.text)
    ):
        return "heading2"
    if (
        _NUMERIC_SECTION_HEADING.match(line.text)
        and line.x0 <= body_left + 12
        and (heading_sized or len(line.text.strip()) <= 24)
        and _is_short_heading_text(line.text)
    ):
        return "heading1"
    if _ORDERED_ITEM.match(line.text):
        return "ordered"
    if re.match(r"^\s*\.\s+", line.text) and line.x0 >= body_left + 12:
        return "ordered"
    if _BULLET_ITEM.match(line.text):
        return "bullet"
    if _PLUGIN_ITEM.match(line.text):
        return "plugin"
    if _looks_like_code_line(line.text):
        return "code"
    if _is_formula_line(line.text):
        return "formula"
    if is_first_line:
        return "title"
    return "body"


def _layout_lines_to_text(
    lines: Iterable[PdfTextLine],
    *,
    body_left: float,
    is_document_start: bool,
) -> tuple[str, list[str]]:
    line_list = list(lines)
    font_sizes = sorted(line.font_size for line in line_list if line.font_size > 0)
    body_font_size = font_sizes[len(font_sizes) // 2] if font_sizes else 0.0
    code_flags = [_looks_like_code_line(line.text) for line in line_list]
    indented = [
        line.x0 >= body_left + 22 and not code_flags[index]
        for index, line in enumerate(line_list)
    ]
    prepared: list[str] = []
    roles: list[str] = []
    previous_role: str | None = None
    previous_line: PdfTextLine | None = None
    for index, line in enumerate(line_list):
        role = _layout_line_role(
            line,
            body_left=body_left,
            is_first_line=is_document_start and index == 0,
            body_font_size=body_font_size,
        )
        if (
            role == "body"
            and previous_role in {"heading1", "heading2"}
            and previous_line is not None
            and abs(line.top - previous_line.top) <= 4.0
            and line.font_size >= previous_line.font_size - 1.0
            and len(line.text) <= 32
            and not _is_sentence_terminal(line.text)
        ):
            role = previous_role
        if role == "body" and indented[index]:
            previous_indented = index > 0 and indented[index - 1]
            next_indented = index + 1 < len(line_list) and indented[index + 1]
            if previous_indented or next_indented:
                role = "bullet"
        if (
            role == "bullet"
            and previous_role == "ordered"
            and previous_line is not None
            and previous_line.text.endswith(("：", ":"))
            and line.x0 <= previous_line.x0 + 14
        ):
            role = "body"
        elif (
            role == "bullet"
            and previous_role == "bullet"
            and previous_line is not None
            and not _is_sentence_terminal(previous_line.text)
            and not _BULLET_ITEM.match(line.text)
        ):
            role = "body"
        text = line.text
        if role == "bullet" and not _BULLET_ITEM.match(text):
            text = f"• {text}"
        elif role == "ordered" and re.match(r"^\s*\.\s+", text):
            text = re.sub(r"^\s*\.\s+", "1. ", text, count=1)
        prepared.append(text)
        roles.append(role)
        previous_role = role
        previous_line = line
    return "\n".join(prepared), roles


def _is_formula_line(line: str) -> bool:
    """识别数学公式；代码行、赋值语句和带分号的代码不会进入公式路线。"""
    text = line.strip()
    if not text or _looks_like_code_line(text):
        return False
    if re.search(r"[{};]", text):
        return False
    if _FORMULA_PREFIX.match(text) or _FORMULA_SYMBOL_PREFIX.match(text):
        return True
    if not re.match(r"^\s*(?:[A-Za-zα-ωΑ-Ω]|\d)", text):
        return False
    if not re.search(r"[=≤≥≈]", text):
        return False
    if re.search(r"[∑∏∫√∞∂∇∀∃α-ωΑ-Ω₀-₉⁰-⁹]", text):
        return True
    if re.search(r"[A-Za-zα-ωΑ-Ω]\s*[=≤≥≈]\s*[^=]+[+\-*/^]", text):
        return True
    return bool(re.fullmatch(r"[\sA-Za-z0-9_+\-*/^().,=]+", text))


def _text_line_role(
    line: str,
    *,
    is_first_line: bool,
    recognize_numeric_headings: bool = False,
) -> str:
    if _SECTION_HEADING.match(line):
        return "heading1"
    if _SUBSECTION_HEADING.match(line):
        return "heading2"
    if recognize_numeric_headings and _NUMERIC_SECTION_HEADING.match(line):
        return "heading1"
    if _ORDERED_ITEM.match(line):
        return "ordered"
    if _BULLET_ITEM.match(line):
        return "bullet"
    if _PLUGIN_ITEM.match(line):
        return "plugin"
    if _looks_like_code_line(line):
        return "code"
    if _is_formula_line(line):
        return "formula"
    if is_first_line:
        return "title"
    return "body"


def _is_sentence_terminal(line: str) -> bool:
    return bool(re.search(r"[。！？!?；;：:.]$", line))


def _needs_word_space(previous: str, current: str) -> bool:
    if not previous or not current:
        return False
    if previous[-1] in _LINE_END_HYPHENS:
        return False
    if previous[-1] in "+-=*/≤≥≈":
        return True
    return previous[-1].isascii() and current[0].isascii() and (
        previous[-1].isalnum() and current[0].isalnum()
    )


def _join_wrapped_lines(lines: list[str]) -> str:
    if not lines:
        return ""
    value = lines[0]
    for line in lines[1:]:
        if _needs_word_space(value, line):
            value += " "
        value += line
    return value


def _group_text_lines(
    content: str,
    *,
    is_document_start: bool = False,
    recognize_numeric_headings: bool = False,
    line_roles: list[str] | None = None,
    spans_out: list[tuple[int, ...]] | None = None,
) -> list[tuple[str, str]]:
    """将 PDF 物理换行合并为逻辑段落，并识别标题、列表和公式。

    ``spans_out`` 会按输出顺序记录每个逻辑段落对应的原始行下标，
    高保真导出用它把段落映射回行级坐标。
    """
    lines = [line.strip() for line in content.splitlines() if line.strip()]
    blocks: list[tuple[str, str]] = []
    current_role: str | None = None
    current_lines: list[str] = []
    current_indices: list[int] = []

    def flush() -> None:
        nonlocal current_role, current_lines, current_indices
        if current_role is not None and current_lines:
            if current_role == "code":
                text = "\n".join(current_lines)
            else:
                text = _join_wrapped_lines(current_lines)
            if current_role == "ordered":
                match = _ORDERED_ITEM.match(text)
                text = match.group(1) if match else text
            elif current_role == "bullet":
                match = _BULLET_ITEM.match(text)
                text = match.group(1) if match else text
            elif current_role == "plugin":
                match = _PLUGIN_ITEM.match(text)
                if match:
                    text = f"{match.group(1)}：{match.group(2)}"
            blocks.append((current_role, text))
            if spans_out is not None:
                spans_out.append(tuple(current_indices))
        current_role = None
        current_lines = []
        current_indices = []

    for index, line in enumerate(lines):
        if line_roles is not None and index < len(line_roles):
            role = line_roles[index]
        else:
            role = _text_line_role(
                line,
                is_first_line=is_document_start and index == 0,
                recognize_numeric_headings=recognize_numeric_headings,
            )
        if role == "code":
            if current_role != "code":
                flush()
                current_role = "code"
                current_lines = []
                current_indices = []
            current_lines.append(line)
            current_indices.append(index)
            continue
        if current_role == "code":
            flush()
        if role in {"title", "heading1", "heading2", "formula"}:
            flush()
            blocks.append((role, line))
            if spans_out is not None:
                spans_out.append((index,))
            continue

        if role in _LIST_ROLES:
            flush()
            current_role = role
            current_lines = [line]
            current_indices = [index]
            if _is_sentence_terminal(line) and not line.endswith(("：", ":")):
                flush()
            continue

        if current_role in _LIST_ROLES:
            current_lines.append(line)
            current_indices.append(index)
            if _is_sentence_terminal(line) and not line.endswith(("：", ":")):
                flush()
            continue

        if current_role != "body":
            current_role = "body"
            current_lines = []
            current_indices = []
        current_lines.append(line)
        current_indices.append(index)
        if _is_sentence_terminal(line):
            flush()

    flush()
    return blocks


