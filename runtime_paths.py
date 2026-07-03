from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Optional


APP_NAME = "ScrmDailyExporter"
TASK_NAME = "每日企微社群任务导出"
DEFAULT_OUTPUT_FOLDER_NAME = "每日企微社群任务导出"


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def app_dir() -> Path:
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def user_base_dir() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / APP_NAME
    return Path.home() / f".{APP_NAME}"


def documents_dir() -> Path:
    return Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Documents"


def default_config_dir() -> Path:
    configured = os.environ.get("SCRM_CONFIG_DIR")
    if configured:
        return Path(os.path.expandvars(configured)).expanduser()
    if is_frozen():
        return user_base_dir()
    return app_dir() / "dev-runtime"


def default_data_dir() -> Path:
    configured = os.environ.get("SCRM_DATA_DIR")
    if configured:
        return Path(os.path.expandvars(configured)).expanduser()
    if is_frozen():
        return documents_dir() / DEFAULT_OUTPUT_FOLDER_NAME
    return app_dir() / "dev-data"


def resolve_dir(value: Optional[str], default: Path) -> Path:
    if value:
        return Path(os.path.expandvars(value)).expanduser().resolve()
    return default.resolve()


def ensure_runtime_dirs(config_dir: Path, data_dir: Path) -> None:
    data_dir.mkdir(parents=True, exist_ok=True)
    for child in ("logs", "state", "chrome-profile"):
        (config_dir / child).mkdir(parents=True, exist_ok=True)


def env_path(config_dir: Path) -> Path:
    return config_dir / ".env"


def state_dir(config_dir: Path) -> Path:
    return config_dir / "state"


def logs_dir(config_dir: Path) -> Path:
    return config_dir / "logs"


def chrome_profile_dir(config_dir: Path) -> Path:
    return config_dir / "chrome-profile"
