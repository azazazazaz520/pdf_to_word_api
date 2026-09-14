"""质量报告与验收判定：告警、门禁、渲染结果合并与保真验收。

质量门禁与验收结论在此处生成，供 worker 编排层与对外接口读取。
"""

from __future__ import annotations

from typing import Any



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
    fidelity_acceptance = quality.get("fidelity_acceptance")
    if isinstance(fidelity_acceptance, dict):
        for key, name in (
            ("page_count_match", "fidelity_page_count"),
            ("no_unexpected_blank_pages", "fidelity_blank_pages"),
            ("ssim_ok", "fidelity_ssim"),
            ("bbox_ok", "fidelity_bbox"),
            ("font_size_ok", "fidelity_font_size"),
        ):
            value = fidelity_acceptance.get(key)
            checks.append(
                {
                    "name": name,
                    "status": "passed" if value is True else "failed",
                    "detail": f"{key}={value!r}",
                }
            )

    page_results = quality.get("page_results") or []
    visual_only_pages = [
        int(item["page"])
        for item in page_results
        if item.get("route") != "blank" and not bool(item.get("editable", True))
    ]
    fallback_pages = list(
        (quality.get("fidelity") or {}).get("page_image_fallback_pages") or []
    )
    if visual_only_pages or fallback_pages:
        checks.append(
            {
                "name": "editable_pages",
                "status": "failed",
                "detail": (
                    f"visual_only_pages={sorted(set(visual_only_pages))}, "
                    f"page_image_fallback_pages={sorted(set(fallback_pages))}"
                ),
            }
        )
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
