"""运行固定公开样本的串行 PDF 转 Word 基准和证据收集。"""

from __future__ import annotations

import argparse
import hashlib
import json
import platform
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from pypdf import PdfReader

# 直接以 scripts/benchmark_short_public_pdfs.py 启动时，把项目根加入导入路径。
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.pdf_worker import process_job


SAMPLE_FILES = {
    "bleu": "bleu.pdf",
    "cn_newspaper": "cn_newspaper.pdf",
    "hkex_return": "hkex_return.pdf",
    "cn_announcement": "cn_announcement.pdf",
    "irs_1040": "irs_1040.pdf",
}


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _source_fingerprint(root: Path) -> str:
    digest = hashlib.sha256()
    paths = [
        *root.glob("src/**/*.py"),
        *root.glob("scripts/*.py"),
        *root.glob("tests/*.py"),
    ]
    for path in sorted(path for path in paths if path.is_file()):
        digest.update(str(path.relative_to(root)).encode("utf-8"))
        digest.update(path.read_bytes())
    return digest.hexdigest()


def _git_head(root: Path) -> str | None:
    try:
        return subprocess.check_output(
            ["git", "rev-parse", "HEAD"],
            cwd=root,
            text=True,
            stderr=subprocess.DEVNULL,
        ).strip()
    except (OSError, subprocess.CalledProcessError):
        return None


def _metadata(source_root: Path, sample_id: str) -> dict[str, Any]:
    path = source_root / SAMPLE_FILES[sample_id]
    metadata_path = path.with_suffix(".json")
    metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
    expected_hash = str(metadata.get("sha256") or "")
    actual_hash = _sha256(path)
    if expected_hash and expected_hash != actual_hash:
        raise ValueError(
            f"样本 {sample_id} 的 SHA-256 不匹配：{actual_hash} != {expected_hash}"
        )
    return {
        "sample": sample_id,
        "input": str(path),
        "sha256": actual_hash,
        "metadata": metadata,
        "source_page_count": len(PdfReader(str(path)).pages),
    }


def _payload(
    *,
    sample: dict[str, Any],
    mode: str,
    output_dir: Path,
    render_word: bool,
) -> dict[str, Any]:
    input_path = Path(sample["input"])
    return {
        "job_id": f"benchmark-{sample['sample']}-{mode}",
        "filename": input_path.name,
        "input_path": str(input_path.resolve()),
        "output_path": str((output_dir / "result.docx").resolve()),
        "progress_path": str((output_dir / "progress.json").resolve()),
        "stage_log_path": str((output_dir / "stages.jsonl").resolve()),
        "cancel_path": str((output_dir / "cancel.requested").resolve()),
        "page_count": int(sample["source_page_count"]),
        "engine": "structure-lite",
        "route_mode": "auto",
        "export_mode": mode,
        "text_min_page_chars": 20,
        "text_min_page_ratio": 0.6,
        "text_high_quality_ratio": 0.8,
        "text_full_page_image_min_pixels": 300000,
        "text_garbled_char_ratio": 0.05,
        "page_image_max_pixels": 4 * 1024 * 1024,
        "page_image_jpeg_quality": 88,
        "embedded_image_max_pixels": 6_000_000,
        "embedded_image_jpeg_quality": 85,
        "embedded_image_png_optimize": False,
        "task_timeout_seconds": 1800.0,
        "ocr_time_budget_seconds": 300.0,
        "render_validation": render_word,
        "render_ssim": True,
        "render_text_compare": True,
        "render_coverage": True,
        "render_timeout_seconds": 300.0,
        "quality_gate_enabled": True,
        "embed_pdf_fonts": True,
        "calibrate_font_metrics": True,
        "include_bookmarks": True,
    }


def _run_one(
    *,
    root: Path,
    source_root: Path,
    output_root: Path,
    sample_id: str,
    mode: str,
    render_word: bool,
) -> dict[str, Any]:
    sample = _metadata(source_root, sample_id)
    output_dir = output_root / sample_id / mode
    output_dir.mkdir(parents=True, exist_ok=False)
    started = datetime.now(timezone.utc).isoformat()
    source_hash_start = _source_fingerprint(root)
    start_time = time.perf_counter()
    result = process_job(
        _payload(
            sample=sample,
            mode=mode,
            output_dir=output_dir,
            render_word=render_word,
        )
    )
    elapsed = round(time.perf_counter() - start_time, 3)
    source_hash_end = _source_fingerprint(root)
    result_record = {
        **sample,
        "mode": mode,
        "started_at": started,
        "finished_at": datetime.now(timezone.utc).isoformat(),
        "elapsed_sec": elapsed,
        "git_head": _git_head(root),
        "source_hash_start": source_hash_start,
        "source_hash_end": source_hash_end,
        "source_consistent": source_hash_start == source_hash_end,
        "render_word_requested": render_word,
        "result": result,
    }
    (output_dir / "result.json").write_text(
        json.dumps(result_record, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return result_record


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--mode", choices=("structured", "fidelity", "both"), default="both")
    parser.add_argument("--samples", nargs="+", choices=tuple(SAMPLE_FILES), required=True)
    parser.add_argument("--render-word", action="store_true")
    return parser


def main() -> int:
    args = _parser().parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    root = Path(__file__).resolve().parents[1]
    source_root = args.source_root.resolve()
    output_root = args.output_root.resolve()
    if not source_root.is_dir():
        raise SystemExit(f"样本目录不存在：{source_root}")
    if output_root.exists() and any(output_root.iterdir()):
        raise SystemExit(f"输出目录已有内容，拒绝覆盖：{output_root}")
    output_root.mkdir(parents=True, exist_ok=True)
    modes = ("structured", "fidelity") if args.mode == "both" else (args.mode,)
    results: list[dict[str, Any]] = []
    for sample_id in args.samples:
        for mode in modes:
            results.append(
                _run_one(
                    root=root,
                    source_root=source_root,
                    output_root=output_root,
                    sample_id=sample_id,
                    mode=mode,
                    render_word=args.render_word,
                )
            )
    run_record = {
        "created_at": datetime.now(timezone.utc).isoformat(),
        "python": sys.version,
        "platform": platform.platform(),
        "git_head": _git_head(root),
        "source_root": str(source_root),
        "output_root": str(output_root),
        "mode": args.mode,
        "samples": args.samples,
        "render_word": args.render_word,
        "results": results,
    }
    (output_root / "run.json").write_text(
        json.dumps(run_record, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(run_record, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
