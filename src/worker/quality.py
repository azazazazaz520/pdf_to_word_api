"""质量报告与验收判定：告警、门禁、渲染结果合并、保真验收、整页兜底。

质量门禁与验收结论在此处生成，供 worker 编排层与对外接口读取。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from ..ir.model import IRBlock, IRDocument, IRWarning
from ..export.document_setup import render_page_image_png
from ..page_render import render_page_image

# 整页贴图的触发依据：该页渲染后文字覆盖率不足，即内容确实丢失
_PAGE_IMAGE_FALLBACK_REASON = "auto_fallback_text_coverage_below_threshold"


def _append_quality_warning(
    quality: dict[str, Any],
    *,
    code: str,
    message: str,
    page: int | None,
    severity: str = "warning",
) -> None:
    warning = {
        "code": code,
        "message": message,
        "page": page,
        "severity": severity,
    }
    warnings = quality.setdefault("warnings", [])
    if warning in warnings:
        return
    warnings.append(warning)
    if severity not in {"warning", "error"}:
        return
    review_pages = set(quality.get("needs_review_pages", []))
    if page is not None:
        review_pages.add(page)
    quality["needs_review_pages"] = sorted(review_pages)


def _evaluate_quality_gate(
    quality: dict[str, Any],
    *,
    enabled: bool,
    page_delta_warn_ratio: float,
    page_delta_warn_absolute: int,
) -> dict[str, Any]:
    if not enabled:
        return {"status": "disabled", "checks": []}
    render = quality.get("render_validation") or {}
    if not render:
        return {"status": "not_run", "checks": []}
    if render.get("status") != "succeeded":
        return {
            "status": "error",
            "checks": [
                {
                    "name": "render_validation",
                    "status": "failed",
                    "detail": render.get("error") or "渲染回读未完成",
                }
            ],
        }
    source_page_count = int(
        render.get("source_page_count")
        or quality.get("effective_source_page_count")
        or quality.get("source_page_count")
        or 0
    )
    page_delta = int(render.get("page_delta") or 0)
    blank_pages = list(
        render.get("unexpected_blank_pages")
        if render.get("unexpected_blank_pages") is not None
        else (render.get("blank_pages") or [])
    )
    delta_threshold = max(
        page_delta_warn_absolute,
        int(source_page_count * page_delta_warn_ratio),
    )
    checks = [
        {
            "name": "render_page_delta",
            "status": "passed" if abs(page_delta) <= delta_threshold else "failed",
            "detail": (
                f"page_delta={page_delta}, threshold={delta_threshold}"
            ),
        },
        {
            "name": "blank_pages",
            "status": "passed" if not blank_pages else "failed",
            "detail": f"blank_pages={blank_pages}",
        },
    ]
    status = "passed" if all(
        check["status"] == "passed" for check in checks
    ) else "failed"
    return {
        "status": status,
        "checks": checks,
        "page_delta_threshold": delta_threshold,
    }


def _apply_fidelity_acceptance(
    quality: dict[str, Any],
    render_result: dict[str, Any],
    *,
    ssim_threshold: float,
) -> None:
    """把高保真验收标准写入质量报告并生成告警。"""
    ssim = render_result.get("ssim") or {}
    text_layout = render_result.get("text_layout") or {}
    matched_lines = int(text_layout.get("matched_line_count") or 0)
    text_layout_available = (
        text_layout.get("status") == "succeeded" and matched_lines > 0
    )
    ssim_available = ssim.get("status") == "succeeded"
    checks = {
        "page_count_match": render_result.get("page_delta") == 0,
        "no_unexpected_blank_pages": not render_result.get(
            "unexpected_blank_pages"
        ),
        "ssim_ok": (
            (not ssim_available)
            or (ssim.get("min_ssim") or 0.0) >= ssim_threshold
        ),
        "bbox_ok": (
            not text_layout_available
            or bool(text_layout.get("passes_bbox", False))
        ),
        "font_size_ok": (
            not text_layout_available
            or bool(text_layout.get("passes_font_size", False))
        ),
    }
    quality["fidelity_acceptance"] = {
        **checks,
        "ssim_available": ssim_available,
        "text_layout_available": text_layout_available,
        "matched_line_count": matched_lines,
        "ssim_threshold": ssim_threshold,
        "min_ssim": ssim.get("min_ssim"),
        "pages_below_threshold": ssim.get("pages_below_threshold", []),
        "max_bbox_error": text_layout.get("max_bbox_error"),
        "max_font_size_error": text_layout.get("max_font_size_error"),
        "unmatched_line_count": text_layout.get("unmatched_line_count"),
        "unexpected_blank_pages": render_result.get("unexpected_blank_pages", []),
        "fallback_region_count": (
            (quality.get("fidelity") or {}).get("fallback_region_count", 0)
        ),
    }
    if not checks["page_count_match"]:
        _append_quality_warning(
            quality,
            code="fidelity_page_count_mismatch",
            message="高保真验收失败：Word 渲染页数与源 PDF 不一致。",
            page=None,
            severity="error",
        )
    if not checks["no_unexpected_blank_pages"]:
        _append_quality_warning(
            quality,
            code="fidelity_unexpected_blank_page",
            message="高保真验收失败：出现源 PDF 中不存在的空白页。",
            page=None,
            severity="error",
        )
    if ssim.get("status") == "succeeded" and not checks["ssim_ok"]:
        _append_quality_warning(
            quality,
            code="fidelity_ssim_below_threshold",
            message=(
                f"高保真验收失败：SSIM 低于 {ssim_threshold}，"
                f"未达标页 {ssim.get('pages_below_threshold')}。"
            ),
            page=None,
        )
    if text_layout.get("status") == "succeeded" and not checks["bbox_ok"]:
        _append_quality_warning(
            quality,
            code="fidelity_bbox_error",
            message=(
                "高保真验收失败：文本 bbox 误差 "
                f"{text_layout.get('max_bbox_error')}pt 超过 3pt。"
            ),
            page=None,
        )
    if text_layout.get("status") == "succeeded" and not checks["font_size_ok"]:
        _append_quality_warning(
            quality,
            code="fidelity_font_size_error",
            message=(
                "高保真验收失败：字号误差 "
                f"{text_layout.get('max_font_size_error')}pt 超过 0.5pt。"
            ),
            page=None,
        )


def _merge_render_validation(
    quality: dict[str, Any],
    render_result: dict[str, Any],
    *,
    expected_source_page_count: int,
    ssim_threshold: float,
) -> None:
    """把一次渲染回读结果合并进质量报告。"""
    quality["rendered_page_count"] = render_result["rendered_page_count"]
    quality["page_delta"] = render_result["page_delta"]
    quality["blank_pages"] = render_result["blank_pages"]
    quality["render_validation"] = render_result
    if render_result["page_delta"] != 0:
        _append_quality_warning(
            quality,
            code="render_page_mismatch",
            message=(
                f"Word 渲染页数 {render_result['rendered_page_count']} "
                f"与源页数 {expected_source_page_count} 不一致。"
            ),
            page=expected_source_page_count,
        )
    for blank_page in render_result.get(
        "unexpected_blank_pages",
        render_result["blank_pages"],
    ):
        _append_quality_warning(
            quality,
            code="render_blank_page",
            message=f"第 {blank_page} 页在 Word 渲染结果中为空白页。",
            page=blank_page,
        )
    _apply_fidelity_acceptance(
        quality,
        render_result,
        ssim_threshold=ssim_threshold,
    )


def _apply_page_image_fallback(
    ir: Any,
    page_numbers: list[int],
    *,
    source_pdf: Path,
    dpi: float,
    max_pixels: int,
) -> list[int]:
    """把 SSIM 不达标的页面替换为整页 PNG 兜底。"""
    applied: list[int] = []
    wanted = {int(number) for number in page_numbers}
    for page in ir.pages:
        if page.page_number not in wanted:
            continue
        try:
            image_bytes = render_page_image_png(
                source_pdf,
                page.page_number - 1,
                dpi=dpi,
                max_pixels=max_pixels,
            )
        except Exception:
            continue
        page.blocks = [
            IRBlock(
                kind="page_image",
                page=page.page_number,
                source="auto_fidelity_fallback",
                image_bytes=image_bytes,
                image_width=page.width,
                image_height=page.height,
                bbox=(0.0, 0.0, page.width, page.height),
                layer="background",
                confidence=1.0,
                fallback_reason=_PAGE_IMAGE_FALLBACK_REASON,
            )
        ]
        page.route = "page_image"
        page.editable = False
        page.reconstruction_confidence = 0.8
        page.fidelity = {
            "rebuild_confidence": 0.8,
            "force_page_image": True,
            "force_page_image_reason": _PAGE_IMAGE_FALLBACK_REASON,
            "page_image_fallback": True,
            "fallback_regions": [
                {
                    "kind": "page_image",
                    "layer": "background",
                    "bbox": [0.0, 0.0, round(page.width, 2), round(page.height, 2)],
                    "reason": _PAGE_IMAGE_FALLBACK_REASON,
                }
            ],
            "native_block_count": 0,
            "image_fallback_count": 1,
        }
        page.add_warning(
            IRWarning(
                code="fidelity_page_image_fallback",
                message=(
                    f"第 {page.page_number} 页重建后 SSIM 低于阈值，"
                    "已自动改为整页图片兜底。"
                ),
                page=page.page_number,
                severity="warning",
            )
        )
        applied.append(page.page_number)
    return applied
