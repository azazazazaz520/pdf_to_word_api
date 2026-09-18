"""质量报告与验收判定：告警、门禁、渲染结果合并与保真验收。

质量门禁与验收结论在此处生成，供 worker 编排层与对外接口读取。
"""

from __future__ import annotations

from typing import Any
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET


def _check_status(
    name: str,
    status: str,
    detail: str,
    *,
    required: bool = True,
) -> dict[str, Any]:
    return {
        "name": name,
        "status": status,
        "detail": detail,
        "required": required,
    }


def evaluate_final_content_quality(
    *,
    source_char_ids: list[str] | tuple[str, ...] | set[str],
    placements: list[dict[str, Any]] | tuple[dict[str, Any], ...],
    source_text: str = "",
    output_text: str = "",
) -> dict[str, Any]:
    """独立核对源字符与最终导出归属，缺少证据时返回未验证。"""
    source_ids = {str(value) for value in source_char_ids if str(value)}
    output_ids = [
        str(char_id)
        for placement in placements
        for char_id in placement.get("source_char_ids", ())
        if str(char_id)
    ]
    checks: list[dict[str, Any]] = []
    if source_ids and output_ids:
        output_set = set(output_ids)
        missing = sorted(source_ids - output_set)
        duplicates = sorted(
            char_id
            for char_id in set(output_ids)
            if output_ids.count(char_id) > 1
        )
        checks.append(
            _check_status(
                "source_character_coverage",
                "failed" if missing else "passed",
                f"missing={missing[:20]}",
            )
        )
        checks.append(
            _check_status(
                "source_character_unique_placement",
                "failed" if duplicates else "passed",
                f"duplicates={duplicates[:20]}",
            )
        )
    else:
        checks.extend(
            (
                _check_status(
                    "source_character_coverage",
                    "unverified",
                    "缺少源字符或最终 DOCX 归属映射",
                ),
                _check_status(
                    "source_character_unique_placement",
                    "unverified",
                    "缺少最终 DOCX 归属映射",
                ),
            )
        )

    normalized_source = "".join(str(source_text).split())
    normalized_output = "".join(str(output_text).split())
    if normalized_source and normalized_output:
        text_status = (
            "passed"
            if normalized_source in normalized_output
            else "failed"
            if len(normalized_output) < len(normalized_source) * 0.99
            else "unverified"
        )
        checks.append(
            _check_status(
                "text_content_presence",
                text_status,
                f"source_length={len(normalized_source)}, output_length={len(normalized_output)}",
            )
        )
    else:
        checks.append(
            _check_status(
                "text_content_presence",
                "unverified",
                "缺少独立的源文本或回读文本",
            )
        )
    required_statuses = [item["status"] for item in checks if item["required"]]
    if "failed" in required_statuses:
        status = "failed"
    elif "unverified" in required_statuses:
        status = "unverified"
    else:
        status = "passed"
    return {
        "status": status,
        "checks": checks,
        "source_character_count": len(source_ids),
        "output_character_mapping_count": len(output_ids),
    }


def read_docx_text(path: Path) -> str:
    """从最终 DOCX XML 回读可见文字，供内容门禁独立取证。"""
    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    texts: list[str] = []
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            root = ET.fromstring(archive.read(name))
            texts.extend(
                node.text or ""
                for node in root.iter(f"{namespace}t")
            )
    return "".join(texts)



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
    structured_mode = quality.get("export_mode") in {"structured", "flow"}
    checks = [
        {
            "name": "render_page_delta",
            "status": (
                "passed"
                if structured_mode or abs(page_delta) <= delta_threshold
                else "failed"
            ),
            "detail": (
                f"page_delta={page_delta}, threshold={delta_threshold}"
                + (", structured_reflow_allowed=True" if structured_mode else "")
            ),
        },
        {
            "name": "blank_pages",
            "status": "passed" if not blank_pages else "failed",
            "detail": f"blank_pages={blank_pages}",
        },
    ]
    fidelity_acceptance = quality.get("fidelity_acceptance")
    if isinstance(fidelity_acceptance, dict) and not structured_mode:
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
    final_content = quality.get("final_content")
    if isinstance(final_content, dict) and final_content.get("status") not in {
        "not_applicable",
        None,
    }:
        checks.append(
            {
                "name": "final_content",
                "status": final_content.get("status", "unverified"),
                "detail": "最终 DOCX 内容归属与回读检查",
            }
        )
    if any(check["status"] == "failed" for check in checks):
        status = "failed"
    elif any(check["status"] == "unverified" for check in checks):
        status = "unverified"
    else:
        status = "passed"
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


def _apply_structured_acceptance(
    quality: dict[str, Any],
    render_result: dict[str, Any],
    *,
    ssim_threshold: float,
) -> None:
    """记录结构化导出的渲染信息；允许正文重排，不按坐标判定失败。"""
    ssim = render_result.get("ssim") or {}
    text_coverage = render_result.get("text_coverage") or {}
    quality["structured_acceptance"] = {
        "page_count_match": render_result.get("page_delta") == 0,
        "reflow_allowed": True,
        "no_unexpected_blank_pages": not render_result.get(
            "unexpected_blank_pages"
        ),
        "ssim_available": ssim.get("status") == "succeeded",
        "ssim_threshold": ssim_threshold,
        "min_ssim": ssim.get("min_ssim"),
        "pages_below_threshold": ssim.get("pages_below_threshold", []),
        "text_coverage": text_coverage,
        "unexpected_blank_pages": render_result.get(
            "unexpected_blank_pages", []
        ),
    }


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
    structured_mode = quality.get("export_mode") in {"structured", "flow"}
    if render_result["page_delta"] != 0:
        _append_quality_warning(
            quality,
            code=("render_page_reflow" if structured_mode else "render_page_mismatch"),
            message=(
                f"Word 渲染页数 {render_result['rendered_page_count']} "
                f"与源页数 {expected_source_page_count} 不一致。"
            ),
            page=expected_source_page_count,
            severity="info" if structured_mode else "warning",
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
    if structured_mode:
        _apply_structured_acceptance(
            quality,
            render_result,
            ssim_threshold=ssim_threshold,
        )
    else:
        _apply_fidelity_acceptance(
            quality,
            render_result,
            ssim_threshold=ssim_threshold,
        )
