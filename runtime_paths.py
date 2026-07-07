from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Optional


APP_NAME = "ScrmDailyExporter"
TASK_NAME = "每日企微私域任务导出"
TEST_TASK_NAME = "每日企微私域任务导出-App测试"
DEFAULT_OUTPUT_FOLDER_NAME = "每日企微私域任务导出"
TEST_OUTPUT_FOLDER_NAME = "ScrmDailyExporterTestData"


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


def test_user_base_dir() -> Path:
    return user_base_dir().with_name(f"{APP_NAME}Test")


def documents_dir() -> Path:
    return Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Documents"


def default_config_dir() -> Path:
    configured = os.environ.get("SCRM_CONFIG_DIR")
    if configured:
        return Path(os.path.expandvars(configured)).expanduser()
    if is_frozen():
        return user_base_dir()
    return app_dir() / "dev-runtime"


def default_test_config_dir() -> Path:
    return test_user_base_dir() if is_frozen() else app_dir() / "dev-runtime-test"


def default_data_dir() -> Path:
    configured = os.environ.get("SCRM_DATA_DIR")
    if configured:
        return Path(os.path.expandvars(configured)).expanduser()
    if is_frozen():
        return documents_dir() / DEFAULT_OUTPUT_FOLDER_NAME
    return app_dir() / "dev-data"


def default_test_data_dir() -> Path:
    return documents_dir() / TEST_OUTPUT_FOLDER_NAME if is_frozen() else app_dir() / "dev-data-test"


def scheduled_task_name(test_mode: bool = False) -> str:
    return TEST_TASK_NAME if test_mode else TASK_NAME


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
