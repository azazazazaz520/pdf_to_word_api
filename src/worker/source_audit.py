"""从 PDFium 原始字符数组生成独立的源字符审计摘要。"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable


def _raw_unicode(text_page: Any, index: int) -> str:
    """按 PDFium 内部字符索引读取字符，不经过布局层文本聚合。"""
    from pypdfium2 import raw as pdfium_raw

    getter = getattr(pdfium_raw, "FPDFText_GetUnicode", None)
    if getter is not None:
        try:
            codepoint = int(getter(text_page.raw, index))
        except Exception:
            codepoint = -1
        if 0 < codepoint <= 0x10FFFF and codepoint != 0xFFFE:
            return chr(codepoint)
        return ""
    try:
        value = text_page.get_text_range(index, 1)
    except Exception:
        return ""
    return value[0] if value else ""


def _character_kind(value: str) -> str:
    if not value:
        return "invalid"
    codepoint = ord(value)
    if codepoint in {0x09, 0x0A, 0x0D}:
        return "layout_control"
    if codepoint < 0x20 and codepoint not in {0x09, 0x0A, 0x0D}:
        return "control"
    if value.isspace():
        return "space"
    if value.isprintable():
        return "visible"
    return "non_printable"


def audit_pdfium_source_characters(
    source_pdf: Path,
    *,
    source_char_ids: Iterable[str],
    sample_limit: int = 20,
) -> dict[str, Any]:
    """将布局使用的源 ID 与独立 PDFium 字符数组进行交叉核对。"""
    source_ids = {str(value) for value in source_char_ids if str(value)}
    raw_ids: set[str] = set()
    kinds: dict[str, str] = {}
    details: dict[str, dict[str, Any]] = {}
    samples: list[dict[str, Any]] = []
    raw_character_count = 0
    raw_visible_count = 0
    raw_space_count = 0
    raw_control_count = 0
    invalid_character_count = 0
    try:
        import pypdfium2 as pdfium

        document = pdfium.PdfDocument(str(source_pdf))
        try:
            for page_number, page in enumerate(document, start=1):
                text_page = page.get_textpage()
                try:
                    _, page_height = (float(value) for value in page.get_size())
                    count = int(text_page.count_chars())
                    for index in range(max(count, 0)):
                        character = _raw_unicode(text_page, index)
                        kind = _character_kind(character)
                        if kind == "invalid":
                            invalid_character_count += 1
                            continue
                        character_id = f"{page_number}:{index}"
                        raw_ids.add(character_id)
                        kinds[character_id] = kind
                        try:
                            x0, y0, x1, y1 = (
                                float(value)
                                for value in text_page.get_charbox(index)
                            )
                            bbox = [
                                x0,
                                page_height - max(y0, y1),
                                x1,
                                page_height - min(y0, y1),
                            ]
                        except Exception:
                            bbox = None
                        details[character_id] = {
                            "text": character,
                            "bbox": bbox,
                        }
                        raw_character_count += 1
                        if kind == "visible":
                            raw_visible_count += 1
                        elif kind == "space":
                            raw_space_count += 1
                        elif kind in {"control", "layout_control", "non_printable"}:
                            raw_control_count += 1
                finally:
                    text_page.close()
                page.close()
        finally:
            document.close()
    except Exception as error:
        return {
            "status": "unverified",
            "independent_raw_pdfium_character_list": False,
            "provenance": "pypdfium2.PdfDocument.textpage.FPDFText_GetUnicode",
            "error": f"{type(error).__name__}: {error}",
            "source_character_count": len(source_ids),
        }

    missing_source_ids = sorted(source_ids - raw_ids)
    unassigned_ids = raw_ids - source_ids
    unassigned_visible_ids = sorted(
        character_id
        for character_id in unassigned_ids
        if kinds.get(character_id) == "visible"
    )
    unassigned_space_ids = sorted(
        character_id
        for character_id in unassigned_ids
        if kinds.get(character_id) == "space"
    )
    unassigned_control_ids = sorted(
        character_id
        for character_id in unassigned_ids
        if kinds.get(character_id)
        in {"control", "layout_control", "non_printable"}
    )
    for character_id in (
        *missing_source_ids[:sample_limit],
        *unassigned_visible_ids[:sample_limit],
        *unassigned_space_ids[:sample_limit],
    ):
        if len(samples) >= sample_limit:
            break
        samples.append(
            {
                "id": character_id,
                "kind": kinds.get(character_id, "missing_source_id"),
                "text": details.get(character_id, {}).get("text"),
                "bbox": details.get(character_id, {}).get("bbox"),
            }
        )

    if not source_ids:
        status = "unverified"
    elif missing_source_ids:
        status = "failed"
    elif unassigned_visible_ids:
        status = "stage_only"
    else:
        status = "stage_only" if unassigned_ids else "passed"
    return {
        "status": status,
        "independent_raw_pdfium_character_list": True,
        "provenance": "pypdfium2.PdfDocument.textpage.FPDFText_GetUnicode",
        "source_character_count": len(source_ids),
        "raw_character_count": raw_character_count,
        "raw_visible_character_count": raw_visible_count,
        "raw_space_count": raw_space_count,
        "raw_control_count": raw_control_count,
        "invalid_character_count": invalid_character_count,
        "missing_source_character_count": len(missing_source_ids),
        "missing_source_character_ids": missing_source_ids[:sample_limit],
        "unassigned_visible_character_count": len(unassigned_visible_ids),
        "unassigned_space_count": len(unassigned_space_ids),
        "unassigned_control_count": len(unassigned_control_ids),
        "unassigned_samples": samples,
    }
