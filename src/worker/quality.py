"""质量报告与验收判定：告警、门禁、渲染结果合并与保真验收。

质量门禁与验收结论在此处生成，供 worker 编排层与对外接口读取。
"""

from __future__ import annotations

from collections import Counter
import posixpath
import re
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

    normalized_source = _canonical_content_text(source_text)
    normalized_output = _canonical_content_text(output_text)
    if normalized_source and normalized_output:
        occurrences = _count_overlapping(normalized_output, normalized_source)
        detail = "source_sequence_matched_once"
        if occurrences == 0:
            math_source = _compact_math_spacing(normalized_source)
            math_occurrences = _count_overlapping(normalized_output, math_source)
            if math_occurrences:
                occurrences = math_occurrences
                detail = "source_sequence_matched_once_with_math_spacing"
        if occurrences == 0:
            if _has_compacted_word_pair(normalized_source, normalized_output):
                text_status = "failed"
                detail = "source_word_space_lost"
            elif _visible_character_sequence(normalized_source) == _visible_character_sequence(
                normalized_output
            ):
                text_status = "passed"
                detail = "source_sequence_matched_with_whitespace_variations"
            elif _visible_character_counter(normalized_source) == _visible_character_counter(
                normalized_output
            ):
                text_status = "unverified"
                detail = "visible_characters_matched_but_sequence_unverified"
            else:
                text_status = "failed"
                detail = "source_sequence_not_found"
        elif occurrences > 1:
            text_status = "failed"
            detail = f"source_sequence_repeated={occurrences}"
        else:
            text_status = "passed"
            detail = "source_sequence_matched_once"
        checks.append(
            _check_status(
                "text_content_presence",
                text_status,
                (
                f"source_length={len(normalized_source)}, "
                f"output_length={len(normalized_output)}, {detail}"
                ),
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


def formula_text_exception_is_local(
    *,
    source_blocks: list[Any] | tuple[Any, ...],
    placements: list[dict[str, Any]] | tuple[dict[str, Any], ...],
    output_text: str,
) -> bool:
    """仅当公式有实际导出证据且普通文本顺序仍完整时允许局部未验证。"""
    formula_blocks = [
        block
        for block in source_blocks
        if getattr(block, "kind", "") == "formula"
        and str(getattr(block, "text", "") or "").strip()
    ]
    if not formula_blocks:
        return False
    formula_ids = {
        str(getattr(block, "block_id", ""))
        for block in formula_blocks
        if str(getattr(block, "block_id", ""))
    }
    evidenced_ids = {
        str(placement.get("block_id"))
        for placement in placements
        if str(placement.get("block_id", "")) in formula_ids
        and placement.get("status") in {"native", "image_fallback", "covered_by_image_fallback"}
        and (
            "formula" in str(placement.get("reason", ""))
            or str(placement.get("reason", "")) == "flow_formula_omml"
        )
    }
    if evidenced_ids != formula_ids:
        return False
    ordinary_source = _canonical_content_text(
        " ".join(
            str(getattr(block, "text", "") or "")
            for block in source_blocks
            if getattr(block, "kind", "") != "formula"
        )
    )
    normalized_output = _canonical_content_text(output_text)
    if not ordinary_source or not normalized_output:
        return False
    for block in formula_blocks:
        formula_text = _canonical_content_text(
            str(getattr(block, "text", "") or "")
        )
        compact_text = re.sub(r"\s+", "", formula_text)
        for candidate in sorted({formula_text, compact_text}, key=len, reverse=True):
            if candidate:
                normalized_output = normalized_output.replace(candidate, " ")
    normalized_output = _canonical_content_text(normalized_output)
    return _count_overlapping(normalized_output, ordinary_source) == 1


def _canonical_content_text(value: str) -> str:
    """统一段落换行的表现，但保留可见词间空格。"""
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _count_overlapping(value: str, needle: str) -> int:
    if not value or not needle:
        return 0
    return sum(
        value[index : index + len(needle)] == needle
        for index in range(len(value) - len(needle) + 1)
    )


def _compact_math_spacing(value: str) -> str:
    """兼容 OMML 回读时公式标签丢失的词间空格。"""
    pattern = re.compile(
        r"(?<![A-Za-z])([A-Za-z][A-Za-z ]{1,60})\s*=\s*"
        r"([A-Za-z0-9.+\-/()]+)"
    )

    def replace(match: re.Match[str]) -> str:
        left = re.sub(r"\s+", "", match.group(1))
        return f"{left}={match.group(2)}"

    return pattern.sub(replace, value)


def _visible_character_counter(value: str) -> Counter[str]:
    return Counter(character for character in value if not character.isspace())


def _visible_character_sequence(value: str) -> str:
    """返回去除空白后的可见字符顺序，用于识别排版产生的空白差异。"""
    return "".join(character for character in value if not character.isspace())


def _has_compacted_word_pair(source: str, output: str) -> bool:
    for match in re.finditer(
        r"(?<![A-Za-z])([A-Za-z]{2,})\s+([A-Za-z]{2,})(?![A-Za-z])",
        source,
    ):
        left, right = match.groups()
        if f"{left}{right}" in output and f"{left} {right}" not in output:
            return True
    return False


def read_docx_text(path: Path, *, include_headers_footers: bool = True) -> str:
    """回读 DOCX 可见文字，可按质量门禁需要只读取正文。"""
    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    math_namespace = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
    markup_namespace = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
    rel_namespace = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    paragraph_tag = f"{namespace}p"
    text_tags = {f"{namespace}t", f"{math_namespace}t"}
    alternate_content_tag = f"{markup_namespace}AlternateContent"
    choice_tag = f"{markup_namespace}Choice"
    fallback_tag = f"{markup_namespace}Fallback"

    def visit_active(element: ET.Element, paragraphs: list[ET.Element]) -> None:
        if element.tag == alternate_content_tag:
            choice = next(
                (child for child in element if child.tag == choice_tag),
                None,
            )
            if choice is None:
                choice = next(
                    (child for child in element if child.tag == fallback_tag),
                    None,
                )
            if choice is not None:
                for child in choice:
                    visit_active(child, paragraphs)
            return
        if element.tag == fallback_tag:
            return
        if element.tag == paragraph_tag:
            paragraphs.append(element)
        for child in element:
            visit_active(child, paragraphs)

    def active_leaf_paragraphs(root: ET.Element) -> list[ET.Element]:
        paragraphs: list[ET.Element] = []
        visit_active(root, paragraphs)
        active_paragraph_ids = {id(paragraph) for paragraph in paragraphs}
        return [
            paragraph
            for paragraph in paragraphs
            if not any(
                id(child) in active_paragraph_ids
                for child in paragraph.iter(paragraph_tag)
                if child is not paragraph
            )
        ]

    def story_text(root: ET.Element) -> str:
        paragraphs: list[str] = []
        for paragraph in active_leaf_paragraphs(root):
            value_parts: list[str] = []

            def append_active_text(element: ET.Element) -> None:
                if element.tag == alternate_content_tag:
                    choice = next(
                        (child for child in element if child.tag == choice_tag),
                        None,
                    )
                    if choice is None:
                        choice = next(
                            (child for child in element if child.tag == fallback_tag),
                            None,
                        )
                    if choice is not None:
                        for child in choice:
                            append_active_text(child)
                    return
                if element.tag == fallback_tag:
                    return
                if element.tag in text_tags:
                    value_parts.append(element.text or "")
                    return
                for child in element:
                    append_active_text(child)

            append_active_text(paragraph)
            value = "".join(value_parts)
            if value:
                paragraphs.append(value)
        return "\n".join(paragraphs)

    with zipfile.ZipFile(path) as archive:
        document_name = "word/document.xml"
        document_root = ET.fromstring(archive.read(document_name))
        stories = [story_text(document_root)]
        rels_name = "word/_rels/document.xml.rels"
        if include_headers_footers and rels_name in archive.namelist():
            rels_root = ET.fromstring(archive.read(rels_name))
            for relationship in rels_root.findall(f"{rel_namespace}Relationship"):
                relationship_type = str(relationship.get("Type") or "")
                if not (
                    relationship_type.endswith("/header")
                    or relationship_type.endswith("/footer")
                ):
                    continue
                target = str(relationship.get("Target") or "")
                target = target.lstrip("/")
                target = posixpath.normpath(posixpath.join("word", target))
                if target.startswith("../") or target not in archive.namelist():
                    continue
                stories.append(story_text(ET.fromstring(archive.read(target))))
    return "\n".join(value for value in stories if value)



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
    if structured_mode:
        structured_acceptance = quality.get("structured_acceptance")
        if not isinstance(structured_acceptance, dict):
            checks.append(
                {
                    "name": "structured_visual_evidence",
                    "status": "unverified",
                    "detail": "结构化导出缺少视觉验收记录",
                }
            )
        else:
            ssim = render.get("ssim") or {}
            text_coverage = structured_acceptance.get("text_coverage") or {}
            if text_coverage.get("status") == "failed":
                visual_status = "failed"
            elif (
                ssim.get("status") != "succeeded"
                or text_coverage.get("status") != "succeeded"
            ):
                visual_status = "unverified"
            elif structured_acceptance.get("pages_below_threshold"):
                # 结构化路线允许重排，像素差异用于标记人工复核，不能直接
                # 当作坐标保真失败；存在明确差异时仍不能报告为完全通过。
                visual_status = "unverified"
            else:
                visual_status = "passed"
            checks.append(
                {
                    "name": "structured_visual_evidence",
                    "status": visual_status,
                    "detail": (
                        f"ssim_status={ssim.get('status')}, "
                        f"pages_below_threshold="
                        f"{structured_acceptance.get('pages_below_threshold', [])}, "
                        f"text_coverage_status={text_coverage.get('status')}"
                    ),
                }
            )
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
            status = (
                "passed"
                if value is True
                else "failed"
                if value is False
                else "unverified"
            )
            checks.append(
                {
                    "name": name,
                    "status": status,
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
    page_delta = render_result.get("page_delta")
    unexpected_blank_pages = render_result.get("unexpected_blank_pages")
    page_count_match = None if page_delta is None else page_delta == 0
    no_unexpected_blank_pages = (
        None
        if unexpected_blank_pages is None
        else not unexpected_blank_pages
    )
    if ssim_available and isinstance(ssim.get("min_ssim"), (int, float)):
        ssim_ok = float(ssim["min_ssim"]) >= ssim_threshold
    elif ssim.get("status") == "failed":
        ssim_ok = False
    else:
        ssim_ok = None
    if text_layout_available:
        bbox_ok = bool(text_layout.get("passes_bbox", False))
        font_size_ok = bool(text_layout.get("passes_font_size", False))
    elif text_layout.get("status") == "failed":
        bbox_ok = False
        font_size_ok = False
    else:
        bbox_ok = None
        font_size_ok = None
    checks = {
        "page_count_match": page_count_match,
        "no_unexpected_blank_pages": no_unexpected_blank_pages,
        "ssim_ok": ssim_ok,
        "bbox_ok": bbox_ok,
        "font_size_ok": font_size_ok,
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
