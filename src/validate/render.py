"""Word 渲染回读与页面检查。

通过 Word COM 把 DOCX 导出为 PDF，并读取渲染结果的页数与空白页；
另提供按 DPI 渲染 PDF 页面为图像的能力，供视觉对比使用。
"""

from __future__ import annotations

import os
import subprocess
import time
from pathlib import Path
from typing import Any

from pypdf import PdfReader


class RenderValidationError(RuntimeError):
    """DOCX 渲染回读失败。"""


_WORD_RENDER_SCRIPT = r'''
$ErrorActionPreference = "Stop"
$docx = $env:PRISM_RENDER_DOCX
$pdf = $env:PRISM_RENDER_PDF
try {
    $word = New-Object -ComObject Word.Application
}
catch {
    Write-Error ("WORD_APP_FAILED: " + $_.Exception.Message)
    exit 10
}
$word.Visible = $false
$word.DisplayAlerts = 0
$word.ScreenUpdating = $false
try {
    try {
        $doc = $word.Documents.Open($docx, $false, $true)
    }
    catch {
        Write-Error ("WORD_OPEN_FAILED: " + $_.Exception.Message)
        exit 11
    }
    try {
        $doc.ExportAsFixedFormat($pdf, 17)
    }
    catch {
        Write-Error ("WORD_EXPORT_FAILED: " + $_.Exception.Message)
        exit 12
    }
    finally {
        $doc.Close($false)
    }
}
finally {
    $word.Quit()
}
'''


def render_docx_to_pdf(
    docx_path: Path,
    pdf_path: Path,
    *,
    timeout_seconds: float = 180.0,
) -> None:
    """使用 Microsoft Word COM 将 DOCX 渲染为 PDF。"""
    if os.name != "nt":
        raise RenderValidationError("当前环境不是 Windows，无法使用 Word COM 渲染")
    if timeout_seconds <= 0:
        raise ValueError("渲染超时时间必须大于 0")
    docx_path = Path(docx_path).resolve()
    pdf_path = Path(pdf_path).resolve()
    if not docx_path.is_file():
        raise RenderValidationError(f"待渲染的 DOCX 不存在：{docx_path}")
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    if pdf_path.exists():
        pdf_path.unlink()

    environment = os.environ.copy()
    environment["PRISM_RENDER_DOCX"] = str(docx_path)
    environment["PRISM_RENDER_PDF"] = str(pdf_path)
    try:
        completed = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                _WORD_RENDER_SCRIPT,
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_seconds,
            env=environment,
            check=False,
        )
    except subprocess.TimeoutExpired as error:
        raise RenderValidationError("Word 渲染超时") from error
    if completed.returncode != 0 or not pdf_path.is_file():
        detail = (completed.stderr or completed.stdout or "").strip()
        raise RenderValidationError(
            "Word 渲染失败："
            f"returncode={completed.returncode}, "
            f"{detail or '未能生成 PDF'}"
        )


def inspect_rendered_pdf(
    pdf_path: Path,
    *,
    source_page_count: int,
) -> dict[str, Any]:
    """统计渲染 PDF 页数和空白页。"""
    pdf_path = Path(pdf_path)
    if not pdf_path.is_file():
        raise RenderValidationError(f"渲染 PDF 不存在：{pdf_path}")
    reader = PdfReader(str(pdf_path))
    rendered_page_count = len(reader.pages)
    blank_pages: list[int] = []
    non_rendering_operators = {
        b"q", b"Q", b"BT", b"ET", b"T*", b"Td", b"TD", b"Tm",
        b"Tc", b"Tw", b"Tz", b"TL", b"Ts", b"Tr", b"Tf",
        b"rg", b"RG", b"g", b"G", b"cm", b"w", b"J", b"j",
        b"M", b"d", b"ri", b"i", b"gs",
    }
    for page_number, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        if text.strip():
            continue
        try:
            contents = page.get_contents()
            operations = list(contents.operations) if contents is not None else []
        except Exception:
            operations = []
        if not any(
            operator not in non_rendering_operators
            for _, operator in operations
        ):
            blank_pages.append(page_number)
    return {
        "status": "succeeded",
        "source_page_count": int(source_page_count),
        "rendered_page_count": rendered_page_count,
        "page_delta": rendered_page_count - int(source_page_count),
        "blank_pages": blank_pages,
        "pdf_path": str(pdf_path),
    }


def _load_numpy() -> Any:
    try:
        import numpy  # noqa: PLC0415

        return numpy
    except Exception:
        return None


def render_pdf_pages(
    pdf_path: Path,
    *,
    dpi: float = 110.0,
    page_indices: list[int] | None = None,
) -> list[Any]:
    """把 PDF 渲染为灰度 numpy 图像列表，用于 SSIM 对比。"""
    numpy = _load_numpy()
    if numpy is None:
        raise RenderValidationError("SSIM 对比需要安装 numpy")
    import pypdfium2 as pdfium

    document = pdfium.PdfDocument(str(pdf_path))
    images: list[Any] = []
    try:
        indices = (
            list(page_indices)
            if page_indices is not None
            else list(range(len(document)))
        )
        for index in indices:
            if index < 0 or index >= len(document):
                continue
            page = document[index]
            bitmap = None
            try:
                scale = max(0.2, min(float(dpi) / 72.0, 4.0))
                bitmap = page.render(scale=scale)
                image = bitmap.to_pil().convert("L")
                try:
                    images.append(numpy.asarray(image, dtype=numpy.float32))
                finally:
                    image.close()
            finally:
                close_bitmap = getattr(bitmap, "close", None)
                if close_bitmap is not None:
                    close_bitmap()
                page.close()
    finally:
        document.close()
    return images
