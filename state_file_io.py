from __future__ import annotations

import errno
import json
import os
import time
from pathlib import Path
from typing import Any


RETRYABLE_WINERRORS = {5, 32}
RETRYABLE_ERRNOS = {errno.EACCES, errno.EPERM}


def _is_retryable_file_error(exc: OSError) -> bool:
    return (
        isinstance(exc, PermissionError)
        or getattr(exc, "winerror", None) in RETRYABLE_WINERRORS
        or getattr(exc, "errno", None) in RETRYABLE_ERRNOS
    )


def _unique_temp_path(path: Path, attempt: int) -> Path:
    return path.with_name(f".{path.name}.{os.getpid()}.{time.time_ns()}.{attempt}.tmp")


def save_json_atomic(path: Path, state: Any, attempts: int = 10) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(state, ensure_ascii=False, indent=2) + "\n"
    last_error: OSError | None = None

    for attempt in range(1, attempts + 1):
        temp_path = _unique_temp_path(path, attempt)
        try:
            temp_path.write_text(payload, encoding="utf-8")
            temp_path.replace(path)
            return
        except OSError as exc:
            last_error = exc
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass
            if not _is_retryable_file_error(exc) or attempt >= attempts:
                raise
            time.sleep(min(1.0, 0.15 * attempt))

    if last_error is not None:
        raise last_error
