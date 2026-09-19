"""任务执行的基础设施：取消与超时、阶段日志写入、OCR 流水线缓存。

被 worker 编排层与质量判定层共用，不依赖任务编排本身。
"""

from __future__ import annotations

import json
import os
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from ..run_validation import build_layout_detection_pipeline, build_pipeline


_PIPELINES: dict[str, Any] = {}
_LAYOUT_PIPELINES: dict[str, Any] = {}


class _Cancelled(Exception):
    pass


class _TimedOut(Exception):
    def __init__(self, stage: str, elapsed_seconds: float, budget_seconds: float):
        self.stage = stage
        self.elapsed_seconds = elapsed_seconds
        self.budget_seconds = budget_seconds
        super().__init__(
            f"任务超过时间预算：stage={stage}, "
            f"elapsed={elapsed_seconds:.3f}s, budget={budget_seconds:.3f}s"
        )


class _ProgressWriter:
    """以原子方式写入任务最新状态，并追加阶段日志。"""

    def __init__(self, progress_path: Path, stage_log_path: Path) -> None:
        self.progress_path = progress_path
        self.stage_log_path = stage_log_path
        self.progress_path.parent.mkdir(parents=True, exist_ok=True)
        self.stage_log_path.parent.mkdir(parents=True, exist_ok=True)

    def emit(
        self,
        stage: str,
        *,
        status: str = "processing",
        progress: int | None = None,
        route: str | None = None,
        route_reason: str | None = None,
        error: str | None = None,
        **details: Any,
    ) -> None:
        event = {
            "stage": stage,
            "status": status,
            "progress": progress,
            "route": route,
            "route_reason": route_reason,
            "error": error,
            "worker_pid": os.getpid(),
            "updated_at": datetime.now(timezone.utc).isoformat(),
            **details,
        }
        temporary_path = self.progress_path.with_name(
            f".{self.progress_path.name}.tmp"
        )
        temporary_path.write_text(
            json.dumps(event, ensure_ascii=False), encoding="utf-8"
        )
        for attempt in range(10):
            try:
                os.replace(temporary_path, self.progress_path)
                break
            except PermissionError:
                if attempt == 9:
                    raise
                time.sleep(0.02 * (attempt + 1))
        with self.stage_log_path.open("a", encoding="utf-8") as stream:
            stream.write(json.dumps(event, ensure_ascii=False) + "\n")


def _get_pipeline(engine: str) -> Any:
    pipeline = _PIPELINES.get(engine)
    if pipeline is None:
        pipeline = build_pipeline(engine)
        _PIPELINES[engine] = pipeline
    return pipeline


def _get_layout_pipeline(engine: str = "structure-lite") -> Any:
    pipeline = _LAYOUT_PIPELINES.get(engine)
    if pipeline is None:
        pipeline = build_layout_detection_pipeline()
        _LAYOUT_PIPELINES[engine] = pipeline
    return pipeline


def _is_cancelled(cancel_path: Path) -> bool:
    return cancel_path.exists()


def _raise_if_cancelled(cancel_path: Path) -> None:
    if _is_cancelled(cancel_path):
        raise _Cancelled()
