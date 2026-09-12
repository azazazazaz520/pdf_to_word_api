"""页面渲染与保真模块共用的 OOXML 写入工具。

* ``render_page_image``：公开的单页渲染接口，逐页路由用它生成页面保真图像；
* ``_add_formula_omml_element``、``_bookmark_name``、``_add_bookmark``、
  ``_add_toc_field``：由 ``export/fidelity.py`` 在运行期按需导入。

保真导出的实现位于 ``export/fidelity.py``，本模块不含导出逻辑。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from .export.document_setup import DEFAULT_PAGE_IMAGE_JPEG_QUALITY, DEFAULT_PAGE_IMAGE_MAX_PIXELS, _render_page_image
from .formula_omml import formula_text_to_omml


def render_page_image(
    source_pdf: Path,
    page_index: int,
    *,
    max_pixels: int = DEFAULT_PAGE_IMAGE_MAX_PIXELS,
    jpeg_quality: int = DEFAULT_PAGE_IMAGE_JPEG_QUALITY,
) -> bytes:
    """公开的单页渲染接口，供逐页路由生成页面保真图像。"""
    if not source_pdf.is_file():
        raise FileNotFoundError(f"页面渲染所需的 PDF 不存在：{source_pdf}")
    if max_pixels < 1:
        raise ValueError("页面图像像素上限必须大于 0")
    import pypdfium2 as pdfium

    document = pdfium.PdfDocument(str(source_pdf))
    try:
        return _render_page_image(
            document,
            page_index,
            max_pixels=max_pixels,
            jpeg_quality=jpeg_quality,
        )
    finally:
        document.close()


# ============================================================================
# IR 流式导出：按 Document IR 的块顺序重排，入口为 export_ir_to_docx
# ============================================================================


def _add_formula_omml_element(paragraph: Any, text: str) -> bool:
    """尝试把公式写成 OMML 并挂到指定段落，失败时返回 False。"""
    try:
        from .formula_omml import formula_text_to_omml
    except Exception:
        return False
    success, omml, _ = formula_text_to_omml(text)
    if not success or not omml:
        return False
    try:
        from docx.oxml import parse_xml

        element = parse_xml(omml.encode("utf-8"))
        paragraph._p.append(element)
        return True
    except Exception:
        return False


def _bookmark_name(title: str, bookmark_id: int) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]", "_", title).strip("_")
    if not cleaned:
        cleaned = f"Bookmark_{bookmark_id}"
    return cleaned[:40]


def _add_bookmark(paragraph: Any, *, name: str, bookmark_id: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def _add_toc_field(document: Document) -> None:
    paragraph = document.add_paragraph()
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = r'TOC \o "1-3" \h \z \u'
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "右键更新域以生成目录"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for element in (begin, instruction, separate, placeholder, end):
        run._r.append(element)




