from __future__ import annotations

import json
import os
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from pypdf import PdfReader

from .validate.report import validate_docx_rendering
from .ir.model import IRBlock, IRWarning
from .fonts.embedding import build_font_plan
from .fonts.metrics import FontMetricRequest, calibrate_font_metrics
from .ir.builder import build_document_ir
from .ir.outline import apply_outline_to_headings, extract_pdf_outline
from .layout.images import DEFAULT_EMBEDDED_IMAGE_JPEG_QUALITY, DEFAULT_EMBEDDED_IMAGE_MAX_PIXELS, DEFAULT_EMBEDDED_IMAGE_PNG_OPTIMIZE
from .layout.layout import extract_pdf_layout
from .pdf_routing import analyze_pdf_text
from .service.ocr_quality import summarize_ocr_page
from .export.document_setup import page_render_scale, render_page_image_png
from .export.fidelity import DEFAULT_FIDELITY_FALLBACK_DPI, export_fidelity_docx
from .page_render import render_page_image
from .run_validation import build_pipeline


def _prepare_fonts(
    payload: dict[str, Any],
    *,
    input_path: Path,
    output_path: Path,
    ir: Any,
    writer: Any,
    route: str,
    route_reason: str,
) -> tuple[Any, dict[str, float], dict[str, Any], dict[str, Any] | None]:
    """构建字体嵌入计划并标定 framePr 偏移。

    返回 (font_plan, font_offsets, font_programs, font_metrics_report)；
    未启用内嵌或标定失败时返回空的偏移与报告，不影响导出继续。
    """
    font_plan = None
    font_offsets: dict[str, float] = {}
    font_programs: dict[str, Any] = {}
    font_metrics_report: dict[str, Any] | None = None
    if bool(payload.get("embed_pdf_fonts", True)):
        try:
            font_plan = build_font_plan(input_path, ir.font_requests())
        except Exception as error:
            font_plan = None
            writer.emit(
                "font_embedding_failed",
                progress=88,
                route=route,
                route_reason=route_reason,
                error=f"{type(error).__name__}: {error}",
            )
        if font_plan is not None:
            for program in font_plan.programs:
                font_programs.setdefault(program.word_name, program)
            writer.emit(
                "font_plan_ready",
                progress=89,
                route=route,
                route_reason=route_reason,
                embedded_font_count=len(font_plan.programs),
                embedded_font_bytes=font_plan.report()["embedded_font_bytes"],
                fallback_fonts=font_plan.report()["fallback_fonts"],
            )
            if bool(payload.get("calibrate_font_metrics", True)):
                try:
                    metric_requests: list[FontMetricRequest] = []
                    seen_fonts: set[str] = set()
                    samples: dict[str, str] = {}
                    for page in ir.pages:
                        for block in page.blocks:
                            for line in block.lines:
                                for span in line.spans:
                                    if len(span.text) > len(
                                        samples.get(span.font_name, "")
                                    ):
                                        samples[span.font_name] = span.text
                    for entry in font_plan.entries.values():
                        if entry.word_name in seen_fonts:
                            continue
                        seen_fonts.add(entry.word_name)
                        metric_requests.append(
                            FontMetricRequest(
                                word_name=entry.word_name,
                                family=entry.word_name,
                                sample=samples.get(entry.word_name, "")[:24],
                                program=font_programs.get(entry.word_name),
                            )
                        )
                    metric_table = calibrate_font_metrics(
                        metric_requests,
                        work_dir=output_path.parent / "font_metrics",
                        cache_path=payload.get("font_metrics_cache"),
                        timeout_seconds=float(
                            payload.get("font_metrics_timeout_seconds", 180.0)
                        ),
                    )
                    font_offsets = dict(metric_table.offsets)
                    font_metrics_report = metric_table.to_dict()
                    writer.emit(
                        "font_metrics_ready",
                        progress=89,
                        route=route,
                        route_reason=route_reason,
                        offsets=font_offsets,
                    )
                except Exception as error:
                    font_offsets = {}
                    font_metrics_report = {
                        "status": "failed",
                        "error": f"{type(error).__name__}: {error}",
                    }
                    writer.emit(
                        "font_metrics_failed",
                        progress=89,
                        route=route,
                        route_reason=route_reason,
                        error=font_metrics_report["error"],
                    )
    return font_plan, font_offsets, font_programs, font_metrics_report


def _run_render_validation(
    payload: dict[str, Any],
    quality: dict[str, Any],
    *,
    ir: Any,
    input_path: Path,
    output_path: Path,
    page_count: int,
    writer: Any,
    route: str,
    route_reason: str,
    emit_export_stage: Any,
    export_fidelity_docx: Any,
    font_plan: Any,
    font_offsets: dict[str, float],
    font_programs: dict[str, Any],
    font_usage_report: dict[str, Any],
    font_plan_report: dict[str, Any] | None,
    font_metrics_report: dict[str, Any] | None,
) -> dict[str, Any]:
    """渲染回读、保真验收与整页图片兜底。

    整页兜底会改写 IR 的页面路由与可编辑状态，因此需要按兜底后的 IR
    重建质量报告并返回给调用方；仅在函数内重新绑定局部变量不会影响调用方。

    由调用方按 render_validation 开关决定是否执行；渲染失败只记录告警。
    """
    if bool(payload.get("render_validation", False)):
        render_started = time.perf_counter()
        try:
            intentional_blank_pages = [
                page.page_number
                for page in ir.pages
                if page.route == "blank"
            ]
            toc_extra_pages = (
                1 if bool(payload.get("include_toc", False)) else 0
            )
            expected_source_page_count = page_count + toc_extra_pages
            expected_blank_pages = intentional_blank_pages
            ssim_threshold = float(payload.get("ssim_threshold", 0.98))
            ssim_dpi = float(payload.get("render_ssim_dpi", 110.0))

            def run_validation(
                *,
                page_indices: list[int] | None = None,
            ) -> dict[str, Any]:
                return validate_docx_rendering(
                    output_path,
                    source_page_count=expected_source_page_count,
                    work_dir=output_path.parent,
                    timeout_seconds=float(
                        payload.get("render_timeout_seconds", 180.0)
                    ),
                    source_pdf=input_path,
                    compare_ssim=bool(payload.get("render_ssim", True)),
                    compare_text=bool(
                        payload.get("render_text_compare", True)
                    ),
                    compare_coverage=bool(
                        payload.get("render_coverage", True)
                    ),
                    ssim_dpi=ssim_dpi,
                    ssim_threshold=ssim_threshold,
                    max_compare_pages=payload.get("render_compare_max_pages"),
                    expected_blank_pages=expected_blank_pages,
                    page_indices=page_indices,
                )

            render_result = run_validation()
            _merge_render_validation(
                quality,
                render_result,
                expected_source_page_count=expected_source_page_count,
                ssim_threshold=ssim_threshold,
            )
            acceptance = quality.get("fidelity_acceptance") or {}
            if toc_extra_pages:
                acceptance["toc_extra_pages"] = toc_extra_pages
                quality["fidelity_acceptance"] = acceptance
            auto_fallback_pages: list[int] = []
            layout_pages = {
                int(item.get("page") or 0): item
                for item in (
                    (render_result.get("text_layout") or {}).get("pages") or []
                )
            }
            coverage_pages = {
                int(item.get("page") or 0): float(item.get("coverage") or 1.0)
                for item in (
                    (render_result.get("text_coverage") or {}).get("pages") or []
                )
            }
            coverage_threshold = float(
                payload.get("coverage_fallback_threshold", 0.90)
            )
            quality["fidelity_fallback_threshold"] = coverage_threshold
            candidates = []
            for page_number, coverage in coverage_pages.items():
                if not 1 <= page_number <= len(ir.pages):
                    continue
                page = ir.pages[page_number - 1]
                if page.route in {"page_image", "blank"}:
                    # 本来就是整页图片/空白页，无需再兜底
                    continue
                # 只依据文字丢失程度贴图；图像相似度不参与判定
                if coverage < coverage_threshold:
                    candidates.append(page_number)
            candidates.sort()
            # 默认不生成兜底整页图片：文字缺失应由导出修正，而不是用截图掩盖
            if not bool(payload.get("fidelity_auto_fallback", False)):
                candidates = []
            auto_fallback_pages: list[int] = []
            if candidates:
                auto_fallback_pages = _apply_page_image_fallback(
                    ir,
                    candidates,
                    source_pdf=input_path,
                    dpi=float(payload.get("fidelity_fallback_dpi", ssim_dpi)),
                    max_pixels=int(
                        payload.get("fidelity_page_image_max_pixels", 8_000_000)
                    ),
                )
            if auto_fallback_pages:
                writer.emit(
                    "fidelity_auto_fallback_started",
                    progress=98,
                    route=route,
                    route_reason="text_coverage_below_threshold",
                    pages=auto_fallback_pages,
                )
                export_fidelity_docx(
                    ir,
                    output_path,
                    title=Path(str(payload["filename"])).stem,
                    source_pdf=input_path,
                    include_toc=bool(payload.get("include_toc", False)),
                    include_bookmarks=bool(
                        payload.get("include_bookmarks", True)
                    ),
                    mode="fidelity",
                    min_text_confidence=payload.get(
                        "fidelity_min_text_confidence"
                    ),
                    fallback_dpi=float(
                        payload.get(
                            "fidelity_fallback_dpi",
                            DEFAULT_FIDELITY_FALLBACK_DPI,
                        )
                    ),
                    fallback_max_pixels=int(
                        payload.get("fidelity_fallback_max_pixels", 2_000_000)
                    ),
                    font_plan=font_plan,
                    font_offsets=font_offsets,
                    font_programs=font_programs,
                    stage_callback=emit_export_stage,
                )
                quality = ir.quality_report(compact=page_count > 200)
                # 页面被替换成整页图片后 IR 里不再有文字，字体报告沿用兜底前的识别结果
                quality["fonts"] = {
                    **font_usage_report,
                    "embedding": font_plan_report,
                    "metrics": font_metrics_report,
                    "measured_before_fallback": True,
                }
                quality["fidelity_auto_fallback_pages"] = auto_fallback_pages
                _append_quality_warning(
                    quality,
                    code="fidelity_page_image_fallback_applied",
                    message=(
                        "以下页面重建后视觉差异过大，已自动改为整页"
                        f"图片兜底：{auto_fallback_pages}。"
                    ),
                    page=None,
                )
                if bool(
                    payload.get("fidelity_revalidate_after_fallback", True)
                ):
                    render_result = run_validation(
                        page_indices=[
                            number - 1 for number in auto_fallback_pages
                        ]
                    )
                    _merge_render_validation(
                        quality,
                        render_result,
                        expected_source_page_count=expected_source_page_count,
                        ssim_threshold=ssim_threshold,
                    )
                else:
                    quality["fidelity_revalidation_skipped"] = True
                quality["fidelity_auto_fallback_pages"] = auto_fallback_pages
            writer.emit(
                "render_validation_completed",
                progress=98,
                route=route,
                route_reason=route_reason,
                elapsed_sec=round(time.perf_counter() - render_started, 3),
                rendered_page_count=render_result["rendered_page_count"],
                page_delta=render_result["page_delta"],
                blank_page_count=len(render_result["blank_pages"]),
                unexpected_blank_page_count=len(
                    render_result.get("unexpected_blank_pages", [])
                ),
                min_ssim=(render_result.get("ssim") or {}).get("min_ssim"),
                max_bbox_error=(
                    render_result.get("text_layout") or {}
                ).get("max_bbox_error"),
                auto_fallback_pages=auto_fallback_pages,
            )
        except Exception as error:
            quality["render_validation"] = {
                "status": "failed",
                "error": f"{type(error).__name__}: {error}",
            }
            _append_quality_warning(
                quality,
                code="render_validation_failed",
                message=f"DOCX 渲染回读失败：{error}",
                page=None,
            )
            writer.emit(
                "render_validation_failed",
                progress=98,
                route=route,
                route_reason=route_reason,
                elapsed_sec=round(time.perf_counter() - render_started, 3),
                error=f"{type(error).__name__}: {error}",
            )
    return quality


def process_job(payload: dict[str, Any]) -> dict[str, Any]:
    """在独立 worker 进程中执行逐页路由的 PDF 转 Word 任务。"""
    job_id = str(payload["job_id"])
    input_path = Path(payload["input_path"])
    output_path = Path(payload["output_path"])
    progress_path = Path(payload["progress_path"])
    stage_log_path = Path(payload["stage_log_path"])
    cancel_path = Path(payload["cancel_path"])
    page_count = int(payload["page_count"])
    route_mode = str(payload.get("route_mode", "auto"))
    export_mode = "fidelity"
    engine = str(payload.get("engine", "structure-lite"))
    task_started = time.perf_counter()
    task_timeout_seconds = float(payload.get("task_timeout_seconds", 300.0))
    ocr_time_budget_seconds = float(payload.get("ocr_time_budget_seconds", 60.0))
    writer = _ProgressWriter(progress_path, stage_log_path)

    route = "ocr" if route_mode == "ocr" else "pending"
    route_reason = "forced_ocr" if route_mode == "ocr" else "pending"
    ocr_budget_active = route_mode == "ocr"

    def raise_if_timed_out(stage: str) -> None:
        budget_seconds = (
            ocr_time_budget_seconds if ocr_budget_active else task_timeout_seconds
        )
        elapsed_seconds = time.perf_counter() - task_started
        if elapsed_seconds > budget_seconds:
            raise _TimedOut(stage, elapsed_seconds, budget_seconds)

    total_pages = max(page_count, 1)

    def emit_export_stage(stage: str, details: dict[str, Any]) -> None:
        _raise_if_cancelled(cancel_path)
        raise_if_timed_out(stage)
        progress = None
        if stage == "ir_page_completed":
            page_number = int(details.get("page") or 0)
            progress = min(90, 60 + int(30 * page_number / total_pages))
        elif stage == "ir_export_completed":
            progress = 95
        clean_details = dict(details)
        clean_details.pop("route", None)
        clean_details.pop("route_reason", None)
        writer.emit(
            stage,
            progress=progress,
            route=route,
            route_reason=route_reason,
            **clean_details,
        )

    try:
        writer.emit(
            "started",
            progress=5,
            job_id=job_id,
            page_count=page_count,
            filename=str(payload["filename"]),
            route_mode=route_mode,
            export_mode=export_mode,
            task_timeout_seconds=task_timeout_seconds,
            ocr_time_budget_seconds=ocr_time_budget_seconds,
        )

        analysis = None
        page_routes: list[str] = []
        if route_mode != "ocr":
            analysis_started = time.perf_counter()
            writer.emit("text_analysis_started", progress=8, job_id=job_id)
            analysis = analyze_pdf_text(
                input_path,
                min_page_chars=int(payload["text_min_page_chars"]),
                full_page_image_min_pixels=int(
                    payload["text_full_page_image_min_pixels"]
                ),
                garbled_char_ratio_threshold=float(
                    payload["text_garbled_char_ratio"]
                ),
            )
            writer.emit(
                "text_analysis_completed",
                progress=12,
                job_id=job_id,
                elapsed_sec=round(time.perf_counter() - analysis_started, 3),
                usable_page_count=analysis.usable_page_count,
                page_count=analysis.page_count,
                text_char_count=analysis.text_char_count,
                usable_page_ratio=round(analysis.usable_page_ratio, 3),
                high_quality_page_count=analysis.high_quality_page_count,
                high_quality_page_ratio=round(analysis.high_quality_page_ratio, 3),
                full_page_image_page_count=analysis.full_page_image_page_count,
                garbled_char_count=analysis.garbled_char_count,
            )
            min_page_chars = int(payload["text_min_page_chars"])
            if route_mode == "text":
                if not analysis.has_usable_text_layer(
                    min_page_chars=min_page_chars,
                    min_page_ratio=float(payload["text_min_page_ratio"]),
                ):
                    raise RuntimeError("PDF 不满足文本层快速路线的检测阈值")
                page_routes = ["text"] * analysis.page_count
            elif route_mode == "auto":
                document_has_text = analysis.usable_page_count > 0
                for page_analysis in analysis.pages:
                    if page_analysis.text_char_count < min_page_chars:
                        if document_has_text and (
                            page_analysis.image_count > 0
                            or page_analysis.full_page_image_count > 0
                        ):
                            page_routes.append("page_image")
                        else:
                            page_routes.append("ocr")
                    elif (
                        page_analysis.image_count >= 8
                        and page_analysis.text_char_count < 1500
                    ):
                        page_routes.append("page_image")
                    elif page_analysis.is_high_quality(min_page_chars=min_page_chars):
                        page_routes.append("text")
                    else:
                        page_routes.append("page_image")
            else:
                raise RuntimeError(f"不支持的路由模式：{route_mode}")

        if not page_routes:
            page_routes = ["ocr"] * page_count
        page_count = len(page_routes)
        total_pages = max(page_count, 1)
        route_summary: dict[str, int] = {}
        for page_route in page_routes:
            route_summary[page_route] = route_summary.get(page_route, 0) + 1

        if route_mode == "ocr":
            route = "ocr"
            route_reason = "forced_ocr"
        elif route_mode == "text":
            route = "text"
            route_reason = "forced_text"
        elif all(item == "text" for item in page_routes):
            route = "text"
            route_reason = "text_layer_complete"
        elif all(item == "ocr" for item in page_routes):
            route = "ocr"
            route_reason = "text_layer_not_usable"
        elif all(item == "page_image" for item in page_routes):
            route = "page_image"
            route_reason = "text_layer_incomplete"
        else:
            route = "mixed"
            route_reason = "per_page_auto"

        ocr_indices = [index for index, value in enumerate(page_routes) if value == "ocr"]
        ocr_budget_active = bool(ocr_indices)
        writer.emit(
            "route_selected",
            progress=15,
            route=route,
            route_reason=route_reason,
            page_routes=page_routes,
            route_summary=route_summary,
        )

        pdf_parsable = True
        try:
            reader = PdfReader(str(input_path))
            page_sizes = [
                (
                    max(float(page.mediabox.width), 0.01),
                    max(float(page.mediabox.height), 0.01),
                )
                for page in reader.pages
            ]
        except Exception:
            pdf_parsable = False
            page_sizes = []
        while len(page_sizes) < page_count:
            page_sizes.append((612.0, 792.0))

        ocr_results: list[Any] = [
            {"parsing_res_list": []} for _ in range(page_count)
        ]
        ocr_quality: list[dict[str, Any] | None] = [None] * page_count
        ocr_page_images: dict[int, bytes] = {}
        if ocr_indices:
            model_started = time.perf_counter()
            writer.emit(
                "model_loading_started",
                progress=15,
                route=route,
                route_reason=route_reason,
                engine=engine,
            )
            pipeline = _get_pipeline(engine)
            writer.emit(
                "model_loading_completed",
                progress=20,
                route=route,
                route_reason=route_reason,
                elapsed_sec=round(time.perf_counter() - model_started, 3),
            )
            inference_started = time.perf_counter()
            writer.emit(
                "inference_started",
                progress=20,
                route=route,
                route_reason=route_reason,
                input=input_path.name,
                ocr_page_count=len(ocr_indices),
                per_page=pdf_parsable,
            )
            completed_ocr_pages = 0
            if pdf_parsable:
                ocr_image_dir = output_path.parent / "ocr_pages"
                ocr_image_dir.mkdir(parents=True, exist_ok=True)
                for index in ocr_indices:
                    _raise_if_cancelled(cancel_path)
                    raise_if_timed_out("ocr_page_render_completed")
                    image_bytes = render_page_image(
                        input_path,
                        index,
                        max_pixels=int(payload["page_image_max_pixels"]),
                        jpeg_quality=int(payload["page_image_jpeg_quality"]),
                    )
                    ocr_page_images[index] = image_bytes
                    temporary_image = ocr_image_dir / f"page_{index + 1}.jpg"
                    temporary_image.write_bytes(image_bytes)
                    page_result: Any = {"parsing_res_list": []}
                    try:
                        for result in pipeline.predict_iter(
                            str(temporary_image)
                        ):
                            page_result = result
                            break
                    finally:
                        temporary_image.unlink(missing_ok=True)
                    _raise_if_cancelled(cancel_path)
                    raise_if_timed_out("inference_page_completed")
                    ocr_results[index] = page_result
                    quality = summarize_ocr_page(
                        page_result,
                        page_number=index + 1,
                    )
                    ocr_quality[index] = quality
                    completed_ocr_pages += 1
                    progress = min(
                        80,
                        20
                        + int(
                            60
                            * completed_ocr_pages
                            / max(len(ocr_indices), 1)
                        ),
                    )
                    writer.emit(
                        "inference_page_completed",
                        progress=progress,
                        route=route,
                        route_reason=route_reason,
                        page=index + 1,
                        **quality,
                    )
            else:
                for result in pipeline.predict_iter(str(input_path)):
                    _raise_if_cancelled(cancel_path)
                    raise_if_timed_out("inference_page_completed")
                    index = ocr_indices[
                        min(completed_ocr_pages, len(ocr_indices) - 1)
                    ]
                    ocr_results[index] = result
                    quality = summarize_ocr_page(
                        result,
                        page_number=index + 1,
                    )
                    ocr_quality[index] = quality
                    completed_ocr_pages += 1
                    progress = min(
                        80,
                        20
                        + int(
                            60
                            * completed_ocr_pages
                            / max(len(ocr_indices), 1)
                        ),
                    )
                    writer.emit(
                        "inference_page_completed",
                        progress=progress,
                        route=route,
                        route_reason=route_reason,
                        page=index + 1,
                        **quality,
                    )
            _raise_if_cancelled(cancel_path)
            writer.emit(
                "inference_completed",
                progress=85,
                route=route,
                route_reason=route_reason,
                result_count=completed_ocr_pages,
                elapsed_sec=round(time.perf_counter() - inference_started, 3),
            )
            ocr_budget_active = False

        layout = None
        if any(value == "text" for value in page_routes):
            layout_started = time.perf_counter()
            text_page_indices = {
                index
                for index, value in enumerate(page_routes)
                if value == "text"
            }
            writer.emit(
                "text_layout_started",
                progress=60,
                route=route,
                route_reason=route_reason,
                page_count=page_count,
            )
            layout = extract_pdf_layout(
                input_path,
                include_page_images=text_page_indices,
                include_fidelity=True,
                image_max_pixels=int(
                    payload.get(
                        "embedded_image_max_pixels",
                        DEFAULT_EMBEDDED_IMAGE_MAX_PIXELS,
                    )
                ),
                image_png_optimize=bool(
                    payload.get(
                        "embedded_image_png_optimize",
                        DEFAULT_EMBEDDED_IMAGE_PNG_OPTIMIZE,
                    )
                ),
                image_jpeg_quality=int(
                    payload.get(
                        "embedded_image_jpeg_quality",
                        DEFAULT_EMBEDDED_IMAGE_JPEG_QUALITY,
                    )
                ),
            )
            adjusted_pages = 0
            for index, current_route in enumerate(page_routes):
                if current_route != "text" or index >= len(layout.pages):
                    continue
                page = layout.pages[index]
                body_lines = [
                    line for line in page.lines if not line.is_header_footer
                ]
                if (
                    len(body_lines) >= 58
                    and not page.tables
                    and not page.images
                ):
                    page_routes[index] = "page_image"
                    adjusted_pages += 1
            if adjusted_pages:
                route_summary = {}
                for page_route in page_routes:
                    route_summary[page_route] = route_summary.get(page_route, 0) + 1
                if all(item == "text" for item in page_routes):
                    route = "text"
                    route_reason = "text_layer_complete"
                elif all(item == "page_image" for item in page_routes):
                    route = "page_image"
                    route_reason = "dense_or_image_page"
                elif all(item == "ocr" for item in page_routes):
                    route = "ocr"
                    route_reason = "text_layer_not_usable"
                else:
                    route = "mixed"
                    route_reason = "per_page_auto"
                writer.emit(
                    "per_page_route_adjusted",
                    progress=66,
                    route=route,
                    route_reason="dense_text_page",
                    adjusted_pages=adjusted_pages,
                    route_summary=route_summary,
                )
            writer.emit(
                "text_layout_completed",
                progress=65,
                route=route,
                route_reason=route_reason,
                page_count=len(layout.pages),
                table_count=layout.table_count,
                elapsed_sec=round(time.perf_counter() - layout_started, 3),
            )

        page_images: dict[int, bytes] = dict(ocr_page_images)
        ocr_scales: list[float | None] = [None] * page_count
        for index in ocr_indices:
            width, height = (
                page_sizes[index] if index < len(page_sizes) else (612.0, 792.0)
            )
            ocr_scales[index] = page_render_scale(
                width,
                height,
                int(payload["page_image_max_pixels"]),
            )
        image_indices = [
            index for index, value in enumerate(page_routes) if value == "page_image"
        ]
        image_indices.extend(index for index in ocr_indices)
        image_indices = sorted(set(image_indices))
        if image_indices:
            writer.emit(
                "page_images_started",
                progress=65,
                route=route,
                route_reason=route_reason,
                page_count=len(image_indices),
            )
            for image_index in image_indices:
                if image_index in page_images:
                    continue
                _raise_if_cancelled(cancel_path)
                raise_if_timed_out("page_image_render_completed")
                page_images[image_index] = render_page_image(
                    input_path,
                    image_index,
                    max_pixels=int(payload["page_image_max_pixels"]),
                    jpeg_quality=int(payload["page_image_jpeg_quality"]),
                )
                writer.emit(
                    "page_image_completed",
                    progress=min(85, 65 + int(20 * (len(page_images) / max(len(image_indices), 1)))),
                    route=route,
                    route_reason=route_reason,
                    page=image_index + 1,
                    page_count=len(page_images),
                )

        text_confidence = (
            [page.quality_score for page in analysis.pages] if analysis else None
        )
        ir = build_document_ir(
            source_pdf=input_path,
            page_routes=page_routes,
            page_sizes=page_sizes,
            layout=layout,
            ocr_results=ocr_results,
            ocr_quality=ocr_quality,
            ocr_scales=ocr_scales,
            page_images=page_images,
            text_confidence=text_confidence,
            route_mode=route_mode,
            export_mode=export_mode,
            engine=engine,
            fidelity=True,
            keep_header_footer=True,
        )
        if bool(payload.get("include_bookmarks", True)):
            try:
                outline_entries = extract_pdf_outline(input_path)
            except Exception:
                outline_entries = []
            ir.metadata["outline"] = outline_entries
            apply_outline_to_headings(ir, outline_entries)
        else:
            ir.metadata["outline"] = []
        font_plan, font_offsets, font_programs, font_metrics_report = _prepare_fonts(
            payload,
            input_path=input_path,
            output_path=output_path,
            ir=ir,
            writer=writer,
            route=route,
            route_reason=route_reason,
        )
        quality = ir.quality_report(compact=page_count > 200)

        writer.emit(
            "docx_export_started",
            progress=90,
            route=route,
            route_reason=route_reason,
            page_count=page_count,
            ir_table_count=ir.table_count,
            ir_media_count=ir.media_count,
        )
        export_started = time.perf_counter()
        fidelity_report = export_fidelity_docx(
            ir,
            output_path,
            title=Path(str(payload["filename"])).stem,
            source_pdf=input_path,
            include_toc=bool(payload.get("include_toc", False)),
            include_bookmarks=bool(payload.get("include_bookmarks", True)),
            mode="fidelity",
            min_text_confidence=payload.get("fidelity_min_text_confidence"),
            fallback_dpi=float(payload.get("fidelity_fallback_dpi", 200.0)),
            fallback_max_pixels=int(
                payload.get("fidelity_fallback_max_pixels", 2_000_000)
            ),
            font_plan=font_plan,
            font_offsets=font_offsets,
            font_programs=font_programs,
            stage_callback=emit_export_stage,
        )
        quality["fidelity"] = fidelity_report
        font_usage_report = ir.font_usage()
        font_plan_report = font_plan.report() if font_plan is not None else None
        quality["fonts"] = {
            **font_usage_report,
            "embedding": font_plan_report,
            "metrics": font_metrics_report,
        }
        export_elapsed = time.perf_counter() - export_started
        output_bytes = output_path.stat().st_size
        writer.emit(
            "docx_export_completed",
            progress=95,
            route=route,
            route_reason=route_reason,
            elapsed_sec=round(export_elapsed, 3),
            output_bytes=output_bytes,
            table_count=ir.table_count,
            media_count=ir.media_count,
            needs_review_pages=quality.get("needs_review_pages", []),
        )
        if bool(payload.get("render_validation", False)):
            quality = _run_render_validation(
                payload,
                quality,
                ir=ir,
                input_path=input_path,
                output_path=output_path,
                page_count=page_count,
                writer=writer,
                route=route,
                route_reason=route_reason,
                emit_export_stage=emit_export_stage,
                export_fidelity_docx=export_fidelity_docx,
                font_plan=font_plan,
                font_offsets=font_offsets,
                font_programs=font_programs,
                font_usage_report=font_usage_report,
                font_plan_report=font_plan_report,
                font_metrics_report=font_metrics_report,
            )
        quality["quality_gate"] = _evaluate_quality_gate(
            quality,
            enabled=bool(payload.get("quality_gate_enabled", True)),
            page_delta_warn_ratio=float(
                payload.get("page_delta_warn_ratio", 0.05)
            ),
            page_delta_warn_absolute=int(
                payload.get("page_delta_warn_absolute", 3)
            ),
        )
        if quality["quality_gate"].get("status") == "failed":
            _append_quality_warning(
                quality,
                code="quality_gate_failed",
                message="质量门禁未通过：请检查渲染页数差或空白页。",
                page=None,
            )
        writer.emit(
            "quality_gate_evaluated",
            progress=99,
            route=route,
            route_reason=route_reason,
            quality_gate=quality["quality_gate"],
        )
        writer.emit(
            "finished",
            status="succeeded",
            progress=100,
            route=route,
            route_reason=route_reason,
            elapsed_sec=round(time.perf_counter() - task_started, 3),
            output_bytes=output_bytes,
            quality=quality,
            route_summary=quality.get("route_summary", route_summary),
        )
        return {
            "status": "succeeded",
            "progress": 100,
            "route": route,
            "route_reason": route_reason,
            "worker_pid": os.getpid(),
            "quality": quality,
        }
    except _TimedOut as error:
        writer.emit(
            "timed_out",
            status="timed_out",
            progress=0,
            route=route,
            route_reason=route_reason,
            error=str(error),
            timeout_stage=error.stage,
            elapsed_sec=round(error.elapsed_seconds, 3),
            timeout_budget_sec=error.budget_seconds,
        )
        return {
            "status": "timed_out",
            "progress": 0,
            "route": route,
            "route_reason": route_reason,
            "worker_pid": os.getpid(),
            "error": str(error),
        }
    except _Cancelled:
        writer.emit(
            "cancelled",
            status="cancelled",
            progress=0,
            route=route,
            route_reason=route_reason,
            elapsed_sec=round(time.perf_counter() - task_started, 3),
        )
        return {
            "status": "cancelled",
            "progress": 0,
            "route": route,
            "route_reason": route_reason,
            "worker_pid": os.getpid(),
        }
    except Exception as error:
        writer.emit(
            "failed",
            status="failed",
            progress=0,
            route=route,
            route_reason=route_reason,
            error=f"{type(error).__name__}: {error}",
            elapsed_sec=round(time.perf_counter() - task_started, 3),
        )
        return {
            "status": "failed",
            "progress": 0,
            "route": route,
            "route_reason": route_reason,
            "worker_pid": os.getpid(),
            "error": f"{type(error).__name__}: {error}",
        }

from .worker.progress import (  # noqa: F401  # 保持原导入路径可用
    _PIPELINES,
    _Cancelled,
    _ProgressWriter,
    _TimedOut,
    _get_pipeline,
    _is_cancelled,
    _raise_if_cancelled,
)
from .worker.quality import (  # noqa: F401  # 保持原导入路径可用
    _append_quality_warning,
    _apply_fidelity_acceptance,
    _apply_page_image_fallback,
    _evaluate_quality_gate,
    _merge_render_validation,
)
