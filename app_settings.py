from __future__ import annotations

import json
import os
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Dict, Optional

import runtime_paths


SETTINGS_FILE_NAME = "app_settings.json"
STATE_FILE_NAME = "export_state.json"


def settings_path(config_dir: Path) -> Path:
    return config_dir / SETTINGS_FILE_NAME


def load_settings(config_dir: Path) -> Dict[str, Any]:
    path = settings_path(config_dir)
    if not path.exists():
        return {}
    try:
        value = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError:
        backup = path.with_suffix(f".bad-{date.today():%Y%m%d}.json")
        path.replace(backup)
        return {}
    return value if isinstance(value, dict) else {}


def save_settings(config_dir: Path, settings: Dict[str, Any]) -> None:
    config_dir.mkdir(parents=True, exist_ok=True)
    path = settings_path(config_dir)
    temp_path = path.with_suffix(".tmp")
    temp_path.write_text(json.dumps(settings, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temp_path.replace(path)


def normalize_data_dir(value: Any) -> Optional[Path]:
    if not isinstance(value, str) or not value.strip():
        return None
    return Path(os.path.expandvars(value.strip())).expanduser()


def normalize_start_date(value: Any) -> Optional[str]:
    if not isinstance(value, str) or not value.strip():
        return None
    parsed = date.fromisoformat(value.strip())
    return parsed.isoformat()


def has_export_history(config_dir: Path) -> bool:
    state_path = runtime_paths.state_dir(config_dir) / STATE_FILE_NAME
    if not state_path.exists():
        return False
    try:
        state = json.loads(state_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError:
        return True
    if not isinstance(state, dict):
        return True
    return bool(state.get("dates") or state.get("last_successful_dates") or state.get("last_successful_date"))


def default_global_start_date(config_dir: Path, today: date) -> Optional[str]:
    if has_export_history(config_dir):
        return None
    return (today - timedelta(days=1)).isoformat()


def ensure_settings(config_dir: Path, default_data_dir: Path, today: date) -> Dict[str, Any]:
    settings = load_settings(config_dir)
    changed = False
    if not settings.get("data_dir"):
        settings["data_dir"] = str(default_data_dir)
        changed = True
    if "global_start_date" not in settings:
        start_date = default_global_start_date(config_dir, today)
        if start_date:
            settings["global_start_date"] = start_date
            changed = True
    if changed:
        save_settings(config_dir, settings)
    return settings
