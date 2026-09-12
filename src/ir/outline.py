from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from pypdf import PdfReader


def extract_pdf_outline(source_pdf: Path) -> list[dict[str, Any]]:
    """提取 PDF outline，返回标题、页码和层级。"""
    reader = PdfReader(str(source_pdf))
    entries: list[dict[str, Any]] = []

    def walk(items: Any, level: int) -> None:
        if not items:
            return
        for item in items:
            if isinstance(item, list):
                walk(item, level + 1)
                continue
            title = getattr(item, "title", None)
            if not title:
                continue
            try:
                page_number = reader.get_destination_page_number(item) + 1
            except Exception:
                page_number = None
            entries.append(
                {
                    "title": str(title).strip(),
                    "page": page_number,
                    "level": level,
                }
            )

    walk(reader.outline, 0)
    return entries

def _normalize_outline_text(value: str) -> str:
    return re.sub(r"[\s\W_]+", "", value).casefold()


def apply_outline_to_headings(
    document: Any,
    outline: list[dict[str, Any]],
) -> int:
    """把 PDF outline 层级匹配到正文标题块。"""
    if not outline:
        return 0
    by_page: dict[int, list[dict[str, Any]]] = {}
    for entry in outline:
        page_number = entry.get("page")
        if isinstance(page_number, int):
            by_page.setdefault(page_number, []).append(entry)
    matched = 0
    for page in getattr(document, "pages", []):
        page_entries = by_page.get(page.page_number, [])
        if not page_entries:
            continue
        for block in page.blocks:
            if not block.is_heading:
                continue
            block_text = _normalize_outline_text(block.text)
            if not block_text:
                continue
            best: dict[str, Any] | None = None
            best_score = 0
            for entry in page_entries:
                title_text = _normalize_outline_text(str(entry.get("title") or ""))
                if not title_text:
                    continue
                if block_text == title_text:
                    score = len(title_text) + 100
                elif block_text in title_text or title_text in block_text:
                    score = min(len(block_text), len(title_text))
                else:
                    continue
                if score > best_score:
                    best = entry
                    best_score = score
            if best is None:
                continue
            level = int(best.get("level") or 0) + 1
            level = max(1, min(3, level))
            block.kind = f"heading{level}"
            block.role = block.kind
            block.level = level
            block.meta["outline_title"] = best.get("title")
            matched += 1
    return matched
