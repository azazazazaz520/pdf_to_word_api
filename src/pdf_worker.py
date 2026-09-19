from __future__ import annotations

import json
import os
import re
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
from .worker.source_audit import audit_pdfium_source_characters
from .export.document_setup import page_render_scale
from .export.fidelity import (
    DEFAULT_FIDELITY_FALLBACK_DPI,
    FIDELITY_MODES,
    export_fidelity_docx,
)
from .export.structured import STRUCTURED_MODES, export_structured_docx
from .page_render import render_page_image
from .run_validation import build_pipeline


def _quality_source_block_text(block: IRBlock) -> str:
    """按导出后的可见顺序生成质量门使用的块文本。"""
    if block.kind in {"table", "html_table"}:
        table = getattr(block, "table", None)
        rows = tuple(getattr(table, "rows", ()) or ())
        if rows:
            return "\n".join(
                "".join(str(value or "") for value in row)
                for row in rows
                if any(str(value or "").strip() for value in row)
            )
        return "\n".join(
            str(cell.text or "")
            for cell in getattr(table, "cells", ())
            if str(cell.text or "").strip()
        )
    if block.kind == "bullet" and block.lines:
        return "\n".join(
            str(line.text or "")
            for line in block.lines
            if str(line.text or "").strip()
        )
    return block.text


def _layout_detection_boxes(result: Any) -> tuple[list[dict[str, Any]], int | None, int | None]:
    """从版面模型结果提取可序列化的区域框和输入像素尺寸。"""
    if isinstance(result, dict):
        detector = result.get("layout_det_res")
        if detector is None and result.get("boxes") is not None:
            detector = result
        width = result.get("width")
        height = result.get("height")
    else:
        detector = getattr(result, "layout_det_res", None)
        if detector is None and getattr(result, "boxes", None) is not None:
            detector = result
        width = getattr(result, "width", None)
        height = getattr(result, "height", None)
    input_image = (
        result.get("input_img")
        if isinstance(result, dict)
        else getattr(result, "input_img", None)
    )
    shape = getattr(input_image, "shape", None)
    if shape is not None and len(shape) >= 2:
        height = height or int(shape[0])
        width = width or int(shape[1])
    if isinstance(detector, dict):
        boxes = detector.get("boxes") or ()
        width = width or detector.get("width")
        height = height or detector.get("height")
    else:
        boxes = getattr(detector, "boxes", ()) or ()
    normalized: list[dict[str, Any]] = []
    for box in boxes:
        if isinstance(box, dict):
            coordinate = box.get("coordinate") or box.get("bbox") or box.get("box")
            label = box.get("label") or box.get("kind") or "body"
            score = box.get("score", box.get("confidence", 0.0))
        else:
            coordinate = getattr(box, "coordinate", None) or getattr(box, "bbox", None)
            label = getattr(box, "label", "body")
            score = getattr(box, "score", getattr(box, "confidence", 0.0))
        if coordinate is None or len(coordinate) < 4:
            continue
        normalized.append(
            {
                "coordinate": [float(value) for value in tuple(coordinate)[:4]],
                "label": str(label),
                "score": float(score or 0.0),
            }
        )
    return normalized, int(width) if width else None, int(height) if height else None


def _collect_layout_model_regions(
    *,
    pipeline: Any,
    input_path: Path,
    page_indices: list[int],
    page_image_max_pixels: int,
    page_image_jpeg_quality: int,
    output_dir: Path,
    raise_if_timed_out: Any,
    cancel_path: Path,
) -> dict[int, dict[str, Any]]:
    """对复杂文字页运行版面检测，结果只回填区域，不替换原生文字。"""
    result: dict[int, dict[str, Any]] = {}
    layout_dir = output_dir / "layout_pages"
    layout_dir.mkdir(parents=True, exist_ok=True)
    temporary_images: list[tuple[int, Path]] = []
    try:
        for index in page_indices:
            _raise_if_cancelled(cancel_path)
            raise_if_timed_out("layout_model_page_started")
            image_bytes = render_page_image(
                input_path,
                index,
                max_pixels=page_image_max_pixels,
                jpeg_quality=page_image_jpeg_quality,
            )
            temporary_image = layout_dir / f"page_{index + 1}.jpg"
            temporary_image.write_bytes(image_bytes)
            temporary_images.append((index, temporary_image))

        _raise_if_cancelled(cancel_path)
        raise_if_timed_out("layout_model_batch_started")
        candidates = pipeline.predict_iter(
            [str(path) for _, path in temporary_images]
        )
        for (index, _), page_result in zip(temporary_images, candidates):
            boxes, pixel_width, pixel_height = _layout_detection_boxes(page_result)
            if boxes:
                result[index + 1] = {
                    "detections": boxes,
                    "pixel_width": pixel_width,
                    "pixel_height": pixel_height,
                    "trigger": "complex_text_page",
                }
    finally:
        for _, temporary_image in temporary_images:
            temporary_image.unlink(missing_ok=True)
    return result


def _prepare_fonts(
    payload: dict[str, Any],
    *,
    input_path: Path,
    output_path: Path,
    ir: Any,
    writer: Any,
    route: str,
    route_reason: str,
    calibrate_offsets: bool = True,
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
            if (
                calibrate_offsets
                and bool(payload.get("calibrate_font_metrics", True))
            ):
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
) -> dict[str, Any]:
    """渲染回读并按导出模式执行验收；不修改 IR，也不替换页面内容。"""
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
            acceptance_key = (
                "structured_acceptance"
                if quality.get("export_mode") in {"structured", "flow"}
                else "fidelity_acceptance"
            )
            acceptance = quality.get(acceptance_key) or {}
            if toc_extra_pages:
                acceptance["toc_extra_pages"] = toc_extra_pages
                quality[acceptance_key] = acceptance
            # 保留字段以兼容旧版质量报告；当前策略永远不替换页面内容。
            quality["fidelity_auto_fallback_pages"] = []
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
                auto_fallback_pages=[],
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
    export_mode = str(payload.get("export_mode") or "structured").strip().lower()
    if export_mode == "flow":
        export_mode = "structured"
    if export_mode not in STRUCTURED_MODES | FIDELITY_MODES:
        raise ValueError(f"不支持的 DOCX 导出模式：{export_mode}")
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
        elif stage == "structured_page_completed":
            page_number = int(details.get("page") or 0)
            progress = min(90, 60 + int(30 * page_number / total_pages))
        elif stage == "ir_export_completed":
            progress = 95
        elif stage == "structured_export_completed":
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
                # 只要 PDF 有文字就进入文字布局；低质量文本也必须保留下来，
                # 这样字符顺序、坐标和字体问题会在结果中暴露出来。
                for page_analysis in analysis.pages:
                    if page_analysis.text_char_count > 0:
                        page_routes.append("text")
                    else:
                        page_routes.append("ocr")
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
        pipeline: Any = None
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

        layout_model_regions: dict[int, dict[str, Any]] = {}
        text_page_indices = [
            index for index, value in enumerate(page_routes) if value == "text"
        ]
        complex_text_indices = [
            index
            for index in text_page_indices
            if analysis is not None
            and index < len(analysis.pages)
            and int(analysis.pages[index].text_char_count) >= int(
                payload.get("layout_model_min_chars", 500)
            )
        ]
        if (
            complex_text_indices
            and bool(payload.get("layout_model_enabled", True))
        ):
            model_started = time.perf_counter()
            writer.emit(
                "layout_model_started",
                progress=45,
                route=route,
                route_reason=route_reason,
                page_count=len(complex_text_indices),
                trigger="complex_text_page",
            )
            try:
                layout_pipeline = _get_layout_pipeline(engine)
                layout_model_regions = _collect_layout_model_regions(
                    pipeline=layout_pipeline,
                    input_path=input_path,
                    page_indices=complex_text_indices,
                    page_image_max_pixels=int(payload["page_image_max_pixels"]),
                    page_image_jpeg_quality=int(payload["page_image_jpeg_quality"]),
                    output_dir=output_path.parent,
                    raise_if_timed_out=raise_if_timed_out,
                    cancel_path=cancel_path,
                )
                writer.emit(
                    "layout_model_completed",
                    progress=52,
                    route=route,
                    route_reason=route_reason,
                    page_count=len(complex_text_indices),
                    region_page_count=len(layout_model_regions),
                    region_count=sum(
                        len(item.get("detections", ()))
                        for item in layout_model_regions.values()
                    ),
                    elapsed_sec=round(time.perf_counter() - model_started, 3),
                )
            except Exception as error:
                writer.emit(
                    "layout_model_failed",
                    progress=52,
                    route=route,
                    route_reason=route_reason,
                    error=f"{type(error).__name__}: {error}",
                    elapsed_sec=round(time.perf_counter() - model_started, 3),
                )

        layout = None
        if any(value == "text" for value in page_routes):
            layout_started = time.perf_counter()
            text_page_indices = set(text_page_indices)
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
                block_reading_order_enabled=bool(
                    payload.get("block_reading_order_enabled", True)
                ),
                block_reading_order_fallback_enabled=bool(
                    payload.get("block_reading_order_fallback_enabled", True)
                ),
                model_regions_by_page=layout_model_regions,
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

        # OCR 图片只用于推理输入，不能进入 DOCX 成为页面背景或兜底内容。
        page_images: dict[int, bytes] = {}
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
            calibrate_offsets=export_mode in FIDELITY_MODES,
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
        if export_mode in STRUCTURED_MODES:
            structured_report = export_structured_docx(
                ir,
                output_path,
                title=Path(str(payload["filename"])).stem,
                source_pdf=input_path,
                include_toc=bool(payload.get("include_toc", False)),
                include_bookmarks=bool(payload.get("include_bookmarks", True)),
                max_fallback_pixels=int(
                    payload.get("fidelity_fallback_max_pixels", 2_000_000)
                ),
                font_plan=font_plan,
                stage_callback=emit_export_stage,
            )
            quality["structured"] = structured_report
        else:
            fidelity_report = export_fidelity_docx(
                ir,
                output_path,
                title=Path(str(payload["filename"])).stem,
                source_pdf=input_path,
                include_toc=bool(payload.get("include_toc", False)),
                include_bookmarks=bool(payload.get("include_bookmarks", True)),
                mode=export_mode,
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
        source_blocks = [
            block
            for page in ir.pages
            for block in page.blocks
            if not str(block.source).startswith("header_footer")
        ]
        ir_source_char_ids = tuple(
            dict.fromkeys(
                char_id
                for block in source_blocks
                for char_id in block.meta.get("source_char_ids", ())
            )
        )
        ir_source_char_id_set = set(ir_source_char_ids)
        layout_source_char_ids = tuple(
            dict.fromkeys(
                char_id
                for page in (layout.pages if layout is not None else ())
                for line in page.lines
                for char_id in line.source_char_ids
                if char_id in ir_source_char_id_set
            )
        )
        source_char_ids = layout_source_char_ids or ir_source_char_ids
        export_report = quality.get("structured") or quality.get("fidelity") or {}
        placements = list(export_report.get("placements") or [])
        if not placements:
            placements = [
                placement
                for page_report in export_report.get("pages", [])
                for placement in page_report.get("placements", [])
            ]
        positioned_block_ids = {
            str(placement.get("block_id"))
            for placement in placements
            if placement.get("reason") == "positioned_text"
        }
        source_text_parts: list[str] = []
        previous_kind = ""
        for block in source_blocks:
            block_text = _quality_source_block_text(block)
            if not block_text:
                continue
            if (
                block.kind == "bullet"
                and str(block.block_id) in positioned_block_ids
                and not str(block_text).lstrip().startswith("•")
            ):
                block_text = f"• {block_text}"
            if source_text_parts and not (
                previous_kind == "formula" and block.kind == "formula"
            ):
                source_text_parts.append("\n")
            source_text_parts.append(
                re.sub(r"\s+", "", block_text)
                if block.kind == "formula"
                else block_text
            )
            previous_kind = block.kind
        source_text = "".join(source_text_parts)
        try:
            output_text = read_docx_text(
                output_path,
                include_headers_footers=False,
            )
        except Exception as error:
            output_text = ""
            quality["content_readback_error"] = (
                f"{type(error).__name__}: {error}"
            )
        if source_char_ids:
            quality["final_content"] = evaluate_final_content_quality(
                source_char_ids=source_char_ids,
                placements=placements,
                source_text=source_text,
                output_text=output_text,
            )
            if (
                quality["final_content"]["status"] == "failed"
                and formula_text_exception_is_local(
                    source_blocks=source_blocks,
                    placements=placements,
                    output_text=output_text,
                )
            ):
                text_check = next(
                    (
                        check
                        for check in quality["final_content"]["checks"]
                        if check["name"] == "text_content_presence"
                    ),
                    None,
                )
                if text_check is not None and text_check["status"] == "failed":
                    text_check["status"] = "unverified"
                    text_check["detail"] = (
                        "公式块已进入 OMML，连续文本回读无法独立确认；"
                        "字符归属检查仍保留"
                    )
                    if not any(
                        check["status"] == "failed"
                        for check in quality["final_content"]["checks"]
                    ):
                        quality["final_content"]["status"] = "unverified"
        else:
            quality["final_content"] = {
                "status": "unverified",
                "checks": [
                    {
                        "name": "source_character_coverage",
                        "status": "unverified",
                        "detail": "OCR 或其他路径未提供独立的源字符归属清单",
                        "required": True,
                    },
                    {
                        "name": "text_content_presence",
                        "status": "unverified",
                        "detail": "缺少独立的源文本与字符级归属证据",
                        "required": True,
                    },
                ],
                "source_character_count": 0,
                "output_character_mapping_count": 0,
            }
        source_audit = audit_pdfium_source_characters(
            input_path,
            source_char_ids=source_char_ids,
        )
        quality["source_audit"] = {
            **source_audit,
            "source": (
                "pdf_text_layout"
                if layout_source_char_ids
                else "ir_blocks"
                if ir_source_char_ids
                else "unavailable"
            ),
            "status": source_audit.get("status", "unverified"),
            "provenance": (
                f"{source_audit.get('provenance')}; filtered_layout_source_ids"
                if layout_source_char_ids
                else source_audit.get("provenance", "unavailable")
            ),
            "ir_character_count": len(ir_source_char_ids),
        }
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
                message="质量门禁未通过：请检查页数、空白页、高保真指标与编辑性。",
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
    _get_layout_pipeline,
    _get_pipeline,
    _is_cancelled,
    _raise_if_cancelled,
)
from .worker.quality import (  # noqa: F401  # 保持原导入路径可用
    _append_quality_warning,
    _apply_fidelity_acceptance,
    _evaluate_quality_gate,
    _merge_render_validation,
    evaluate_final_content_quality,
    formula_text_exception_is_local,
    read_docx_text,
)
