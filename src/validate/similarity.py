"""逐页视觉相似度对比。

对源 PDF 与渲染结果的对应页面计算 SSIM，并汇总均值、最低值与未达标页。
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any

from .render import (
    RenderValidationError,
    _load_numpy,
    render_pdf_pages,
)


def _gaussian_blur(image: Any, *, sigma: float = 1.5, radius: int = 5) -> Any:
    numpy = _load_numpy()
    if numpy is None:
        raise RenderValidationError("SSIM 对比需要安装 numpy")
    try:
        from scipy.ndimage import gaussian_filter  # noqa: PLC0415

        return gaussian_filter(image, sigma=sigma, mode="reflect")
    except Exception:
        pass
    try:
        import cv2  # noqa: PLC0415

        return cv2.GaussianBlur(image, (0, 0), sigma)
    except Exception:
        pass
    size = radius * 2 + 1
    kernel = numpy.exp(
        -((numpy.arange(size) - radius) ** 2) / (2 * sigma * sigma)
    )
    kernel /= kernel.sum()
    padded = numpy.pad(image, radius, mode="reflect")
    horizontal = numpy.apply_along_axis(
        lambda row: numpy.convolve(row, kernel, mode="valid"), 1, padded
    )
    transposed = numpy.pad(horizontal, ((radius, radius), (0, 0)), mode="reflect")
    return numpy.apply_along_axis(
        lambda column: numpy.convolve(column, kernel, mode="valid"),
        0,
        transposed,
    )


def _resize_pair(left: Any, right: Any) -> tuple[Any, Any]:
    if left.shape == right.shape:
        return left, right
    from PIL import Image  # noqa: PLC0415

    height = max(left.shape[0], right.shape[0])
    width = max(left.shape[1], right.shape[1])
    resized: list[Any] = []
    for image in (left, right):
        if image.shape[0] == height and image.shape[1] == width:
            resized.append(image)
            continue
        pil = Image.fromarray(image.astype("uint8"), mode="L")
        try:
            resized.append(
                _load_numpy().asarray(
                    pil.resize((width, height), Image.Resampling.BILINEAR),
                    dtype=_load_numpy().float32,
                )
            )
        finally:
            pil.close()
    return resized[0], resized[1]


def ssim_score(left: Any, right: Any, *, data_range: float = 255.0) -> float:
    """计算两张灰度图的 SSIM（Wang 等 2004 定义）。"""
    numpy = _load_numpy()
    if numpy is None:
        raise RenderValidationError("SSIM 对比需要安装 numpy")
    left, right = _resize_pair(
        left.astype(numpy.float32), right.astype(numpy.float32)
    )
    if left.size == 0:
        return 0.0
    mu_left = _gaussian_blur(left)
    mu_right = _gaussian_blur(right)
    mu_left_sq = mu_left * mu_left
    mu_right_sq = mu_right * mu_right
    mu_product = mu_left * mu_right
    sigma_left = _gaussian_blur(left * left) - mu_left_sq
    sigma_right = _gaussian_blur(right * right) - mu_right_sq
    sigma_product = _gaussian_blur(left * right) - mu_product
    c1 = (0.01 * data_range) ** 2
    c2 = (0.03 * data_range) ** 2
    numerator = (2 * mu_product + c1) * (2 * sigma_product + c2)
    denominator = (mu_left_sq + mu_right_sq + c1) * (
        sigma_left + sigma_right + c2
    )
    ssim_map = numerator / denominator
    return float(ssim_map.mean())


def compare_pdf_pages(
    source_pdf: Path,
    rendered_pdf: Path,
    *,
    dpi: float = 110.0,
    max_pages: int | None = None,
    threshold: float = 0.98,
    page_indices: list[int] | None = None,
) -> dict[str, Any]:
    """逐页比较源 PDF 与渲染 PDF 的 SSIM。

    ``page_indices`` 使用 0 基页码；提供时只比较这些页，用于
    "只复检自动兜底页" 的快速复验。
    """
    started = time.perf_counter()
    selected = (
        [int(index) for index in page_indices]
        if page_indices is not None
        else None
    )
    source_images = render_pdf_pages(source_pdf, dpi=dpi, page_indices=selected)
    rendered_images = render_pdf_pages(
        rendered_pdf, dpi=dpi, page_indices=selected
    )
    page_count = min(len(source_images), len(rendered_images))
    if max_pages is not None:
        page_count = min(page_count, max(0, int(max_pages)))
    per_page: list[dict[str, Any]] = []
    scores: list[float] = []
    for index in range(page_count):
        score = ssim_score(source_images[index], rendered_images[index])
        scores.append(score)
        per_page.append(
            {
                "page": (
                    selected[index] + 1
                    if selected is not None and index < len(selected)
                    else index + 1
                ),
                "ssim": round(score, 4),
            }
        )
    below_threshold = [
        entry["page"] for entry in per_page if entry["ssim"] < threshold
    ]
    return {
        "status": "succeeded",
        "source_page_count": len(source_images),
        "rendered_page_count": len(rendered_images),
        "compared_page_count": page_count,
        "selected_page_numbers": (
            [entry["page"] for entry in per_page] if selected is not None else None
        ),
        "threshold": threshold,
        "mean_ssim": round(sum(scores) / len(scores), 4) if scores else None,
        "min_ssim": round(min(scores), 4) if scores else None,
        "pages_below_threshold": below_threshold,
        "per_page_ssim": per_page,
        "elapsed_seconds": round(time.perf_counter() - started, 3),
    }
