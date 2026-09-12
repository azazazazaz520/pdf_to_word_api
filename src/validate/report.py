"""渲染回读总入口：串起渲染、视觉对比与文本比对，产出验收数据。

pdf_worker 与服务通过 validate_docx_rendering 获取逐页 SSIM、覆盖率与
布局误差，据此生成 fidelity_acceptance 结论。
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any

from . import render, similarity, text_compare

_SIBLINGS = (render, similarity, text_compare)


def _resolve(name: str) -> Any:
    """解析被调用函数，优先采用对外门面上可能被替换的版本。

    `src/docx_render_validation.py` 是包内实现的对外门面，测试会以该门面为
    补丁目标（如 patch("src.docx_render_validation.render_docx_to_pdf")）。
    这里先看门面中的同名属性，使补丁生效；未被打补丁时回落到入口文件自身。
    """
    import sys

    facade = sys.modules.get("src.docx_render_validation")
    if facade is not None:
        candidate = getattr(facade, name, None)
        if candidate is not None:
            return candidate
    for module in _SIBLINGS:
        candidate = getattr(module, name, None)
        if candidate is not None:
            return candidate
    raise AttributeError(name)


def validate_docx_rendering(
    docx_path: Path,
    *,
    source_page_count: int,
    work_dir: Path,
    timeout_seconds: float = 180.0,
    source_pdf: Path | None = None,
    compare_ssim: bool = False,
    compare_text: bool = False,
    compare_coverage: bool = False,
    ssim_dpi: float = 110.0,
    max_compare_pages: int | None = None,
    ssim_threshold: float = 0.98,
    expected_blank_pages: list[int] | None = None,
    page_indices: list[int] | None = None,
) -> dict[str, Any]:
    """渲染 DOCX 并返回页数变化、空白页和高保真对比指标。"""
    docx_path = Path(docx_path)
    work_dir = Path(work_dir)
    pdf_path = (work_dir / f"{docx_path.stem}.rendered.pdf").resolve()
    started = time.perf_counter()
    _resolve("render_docx_to_pdf")(
        docx_path,
        pdf_path,
        timeout_seconds=timeout_seconds,
    )
    result = _resolve("inspect_rendered_pdf")(
        pdf_path,
        source_page_count=source_page_count,
    )
    expected = sorted(set(expected_blank_pages or []))
    unexpected_blank_pages = [
        page for page in result["blank_pages"] if page not in expected
    ]
    result["expected_blank_pages"] = expected
    result["unexpected_blank_pages"] = unexpected_blank_pages
    result["pdf_path"] = str(pdf_path)
    if source_pdf is not None and source_pdf.is_file():
        if compare_ssim:
            try:
                result["ssim"] = _resolve("compare_pdf_pages")(
                    source_pdf,
                    pdf_path,
                    dpi=ssim_dpi,
                    max_pages=max_compare_pages,
                    threshold=ssim_threshold,
                    page_indices=page_indices,
                )
            except Exception as error:
                result["ssim"] = {
                    "status": "failed",
                    "error": f"{type(error).__name__}: {error}",
                }
        if compare_coverage:
            try:
                result["text_coverage"] = _resolve("compare_text_coverage")(
                    source_pdf,
                    pdf_path,
                )
            except Exception as error:
                result["text_coverage"] = {
                    "status": "failed",
                    "error": f"{type(error).__name__}: {error}",
                }
        if compare_text:
            try:
                result["text_layout"] = _resolve("compare_text_layout")(
                    source_pdf,
                    pdf_path,
                    max_pages=max_compare_pages,
                    page_indices=page_indices,
                )
            except Exception as error:
                result["text_layout"] = {
                    "status": "failed",
                    "error": f"{type(error).__name__}: {error}",
                }
    result["elapsed_seconds"] = round(time.perf_counter() - started, 3)
    return result
