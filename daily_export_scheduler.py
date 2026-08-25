#!/usr/bin/env python3
"""
Catch up daily SCRM exports with the least possible manual work.

This scheduler keeps export state locally, refreshes the logged-in SCRM
browser token through Chrome DevTools Protocol after a manual QR login, and
then runs each registered export task once per pending date.
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import re
import socket
import struct
import subprocess
import sys
import threading
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.error import HTTPError, URLError
from urllib.parse import quote, unquote, urlparse
from urllib.request import Request, urlopen

import export_chat_group_analysis_by_chat as chat_analysis_exporter
import export_group_send_customer_group as customer_group_exporter
import export_reach_daily_excel as reach_excel_exporter
import export_reach_customer_summary as reach_summary_exporter
import export_super_group_undelivered as exporter
import sync_feishu_reach_workbook as feishu_sync
import app_settings
import runtime_paths
import state_file_io


LOGIN_URL = "https://scrm.cotticoffee.cc/login"
SCRM_HOST = "scrm.cotticoffee.cc"
DEFAULT_LOGIN_WAIT_MINUTES = 20
DEFAULT_CATCHUP_LOOKBACK_DAYS = 7
DEFAULT_CHROME_DEBUG_PORT = 9333
STATE_DIR_NAME = "state"
STATE_FILE_NAME = "export_state.json"
STATUS_FILE_NAME = "latest_status.txt"
SUPER_GROUP_TASK_ID = "super_group_undelivered"
CHAT_ANALYSIS_TASK_ID = "chat_group_analysis_by_chat"
REACH_SUMMARY_TASK_ID = "reach_customer_summary"
CUSTOMER_GROUP_TASK_ID = "group_send_customer_group_export"
REACH_EXCEL_SUMMARY_TASK_ID = "reach_excel_summary"
STORE_GROUP_REACH_TASK_ID = "store_group_reach_summary"
FEISHU_SYNC_TASK_ID = "feishu_reach_sync"


class LoginRefreshError(RuntimeError):
    pass


@dataclass
class SchedulerConfig:
    root: Path
    config_dir: Path
    env_path: Path
    state_dir: Path
    state_path: Path
    status_path: Path
    chrome_profile_dir: Path
    login_wait_minutes: int
    catchup_lookback_days: int
    global_start_date: Optional[date]
    chrome_debug_port: int


@dataclass(frozen=True)
class ExportTask:
    task_id: str
    label: str
    initial_mode: str


EXPORT_TASKS = [
    ExportTask(SUPER_GROUP_TASK_ID, "超级群发未送达", "lookback"),
    ExportTask(CHAT_ANALYSIS_TASK_ID, "客户群分析-按群聊", "lookback"),
    ExportTask(REACH_SUMMARY_TASK_ID, "群发客户及朋友圈触达人数", "lookback"),
    ExportTask(CUSTOMER_GROUP_TASK_ID, "群发客户群导出", "lookback"),
    ExportTask(REACH_EXCEL_SUMMARY_TASK_ID, "触达人数汇总", "lookback"),
    ExportTask(STORE_GROUP_REACH_TASK_ID, "门店分组触达人数", "lookback"),
    ExportTask(FEISHU_SYNC_TASK_ID, "同步在线表", "lookback"),
]


class WebSocketConnection:
    def __init__(self, url: str, timeout: float = 10.0) -> None:
        parsed = urlparse(url)
        if parsed.scheme != "ws":
            raise LoginRefreshError(f"Only ws:// CDP URLs are supported: {url}")
        self.host = parsed.hostname or "127.0.0.1"
        self.port = parsed.port or 80
        self.path = parsed.path
        if parsed.query:
            self.path = f"{self.path}?{parsed.query}"
        self.timeout = timeout
        self.sock: Optional[socket.socket] = None

    def __enter__(self) -> "WebSocketConnection":
        self.sock = socket.create_connection((self.host, self.port), timeout=self.timeout)
        self.sock.settimeout(self.timeout)
        key = base64.b64encode(os.urandom(16)).decode("ascii")
        request = (
            f"GET {self.path} HTTP/1.1\r\n"
            f"Host: {self.host}:{self.port}\r\n"
            "Upgrade: websocket\r\n"
            "Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\n"
            "Sec-WebSocket-Version: 13\r\n"
            "\r\n"
        )
        self.sock.sendall(request.encode("ascii"))
        response = self._recv_until(b"\r\n\r\n")
        if b" 101 " not in response.split(b"\r\n", 1)[0]:
            raise LoginRefreshError("Chrome CDP websocket handshake failed")
        accept = base64.b64encode(
            hashlib.sha1((key + "258EAFA5-E914-47DA-95CA-C5AB0DC85B11").encode("ascii")).digest()
        ).decode("ascii")
        if accept.encode("ascii") not in response:
            raise LoginRefreshError("Chrome CDP websocket accept header mismatch")
        return self

    def __exit__(self, *_: Any) -> None:
        if self.sock:
            try:
                self.sock.close()
            finally:
                self.sock = None

    def _recv_exact(self, size: int) -> bytes:
        if not self.sock:
            raise LoginRefreshError("Websocket is not connected")
        chunks: List[bytes] = []
        remaining = size
        while remaining:
            chunk = self.sock.recv(remaining)
            if not chunk:
                raise LoginRefreshError("Websocket closed unexpectedly")
            chunks.append(chunk)
            remaining -= len(chunk)
        return b"".join(chunks)

    def _recv_until(self, marker: bytes) -> bytes:
        if not self.sock:
            raise LoginRefreshError("Websocket is not connected")
        data = bytearray()
        while marker not in data:
            chunk = self.sock.recv(4096)
            if not chunk:
                raise LoginRefreshError("Connection closed before websocket handshake finished")
            data.extend(chunk)
        return bytes(data)

    def send_text(self, text: str) -> None:
        if not self.sock:
            raise LoginRefreshError("Websocket is not connected")
        payload = text.encode("utf-8")
        if len(payload) < 126:
            header = bytes([0x81, 0x80 | len(payload)])
        elif len(payload) < 65536:
            header = bytes([0x81, 0x80 | 126]) + struct.pack("!H", len(payload))
        else:
            header = bytes([0x81, 0x80 | 127]) + struct.pack("!Q", len(payload))
        mask = os.urandom(4)
        masked = bytes(byte ^ mask[index % 4] for index, byte in enumerate(payload))
        self.sock.sendall(header + mask + masked)

    def send_pong(self, payload: bytes) -> None:
        if not self.sock:
            return
        if len(payload) > 125:
            payload = payload[:125]
        mask = os.urandom(4)
        masked = bytes(byte ^ mask[index % 4] for index, byte in enumerate(payload))
        self.sock.sendall(bytes([0x8A, 0x80 | len(payload)]) + mask + masked)

    def recv_text(self) -> str:
        fragments: List[bytes] = []
        while True:
            first, second = self._recv_exact(2)
            fin = bool(first & 0x80)
            opcode = first & 0x0F
            masked = bool(second & 0x80)
            length = second & 0x7F
            if length == 126:
                length = struct.unpack("!H", self._recv_exact(2))[0]
            elif length == 127:
                length = struct.unpack("!Q", self._recv_exact(8))[0]
            mask_key = self._recv_exact(4) if masked else b""
            payload = self._recv_exact(length) if length else b""
            if masked:
                payload = bytes(byte ^ mask_key[index % 4] for index, byte in enumerate(payload))

            if opcode == 0x8:
                raise LoginRefreshError("Chrome CDP websocket closed")
            if opcode == 0x9:
                self.send_pong(payload)
                continue
            if opcode == 0xA:
                continue
            if opcode in (0x1, 0x0):
                fragments.append(payload)
                if fin:
                    return b"".join(fragments).decode("utf-8")


class CdpClient:
    def __init__(self, websocket_url: str) -> None:
        self.websocket_url = websocket_url
        self.next_id = 0

    def __enter__(self) -> "CdpClient":
        self.ws = WebSocketConnection(self.websocket_url).__enter__()
        return self

    def __exit__(self, *args: Any) -> None:
        self.ws.__exit__(*args)

    def call(self, method: str, params: Optional[Dict[str, Any]] = None, timeout: float = 10.0) -> Dict[str, Any]:
        self.next_id += 1
        message_id = self.next_id
        self.ws.send_text(json.dumps({"id": message_id, "method": method, "params": params or {}}))
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            data = json.loads(self.ws.recv_text())
            if data.get("id") != message_id:
                continue
            if "error" in data:
                raise LoginRefreshError(f"CDP {method} failed: {data['error']}")
            return data.get("result") or {}
        raise LoginRefreshError(f"CDP {method} timed out")


def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Catch up SCRM super-group exports.")
    parser.add_argument("--today", help="Override today's date for tests, YYYY-MM-DD.")
    parser.add_argument("--config-dir", help="Runtime config directory. Defaults to user app data.")
    parser.add_argument("--data-dir", help="Export output directory. Defaults to the user's Documents folder.")
    parser.add_argument("--start-date", help="Start catch-up from this date, YYYY-MM-DD.")
    parser.add_argument("--plan-only", action="store_true", help="Print pending dates without exporting or writing state.")
    parser.add_argument("--login-only", action="store_true", help="Open Chrome, wait for QR login, update .env, then exit.")
    return parser.parse_args(argv)


def load_config(data_dir: Path, config_dir: Path) -> SchedulerConfig:
    env_path = runtime_paths.env_path(config_dir)
    env_file = exporter.load_dotenv(env_path)
    settings = app_settings.load_settings(config_dir)
    global_start_date: Optional[date] = None
    raw_global_start_date = settings.get("global_start_date")
    if isinstance(raw_global_start_date, str) and raw_global_start_date.strip():
        try:
            global_start_date = date.fromisoformat(raw_global_start_date.strip())
        except ValueError:
            print(
                f"Ignoring invalid global_start_date in app settings: {raw_global_start_date}",
                file=sys.stderr,
            )
    login_wait_minutes = int(
        exporter.env_value(env_file, "LOGIN_WAIT_MINUTES", str(DEFAULT_LOGIN_WAIT_MINUTES))
    )
    catchup_lookback_days = int(
        exporter.env_value(env_file, "CATCHUP_LOOKBACK_DAYS", str(DEFAULT_CATCHUP_LOOKBACK_DAYS))
    )
    chrome_debug_port = int(
        exporter.env_value(env_file, "CHROME_DEBUG_PORT", str(DEFAULT_CHROME_DEBUG_PORT))
    )
    state_dir = runtime_paths.state_dir(config_dir)
    return SchedulerConfig(
        root=data_dir,
        config_dir=config_dir,
        env_path=env_path,
        state_dir=state_dir,
        state_path=state_dir / STATE_FILE_NAME,
        status_path=state_dir / STATUS_FILE_NAME,
        chrome_profile_dir=runtime_paths.chrome_profile_dir(config_dir),
        login_wait_minutes=max(1, login_wait_minutes),
        catchup_lookback_days=max(1, catchup_lookback_days),
        global_start_date=global_start_date,
        chrome_debug_port=chrome_debug_port,
    )


def parse_today(value: Optional[str]) -> date:
    if value:
        return datetime.strptime(value, "%Y-%m-%d").date()
    return date.today()


def parse_start_date(value: Optional[str], today: date) -> Optional[date]:
    if not value:
        return None
    parsed = datetime.strptime(value, "%Y-%m-%d").date()
    yesterday = today - timedelta(days=1)
    if parsed > yesterday:
        raise ValueError(f"--start-date cannot be later than yesterday ({yesterday:%Y-%m-%d}).")
    return parsed


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def emit_event(event: str, **fields: Any) -> None:
    payload = {"event": event, "time": now_text()}
    payload.update(fields)
    print(
        "@@SCRM_STATUS " + json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
        flush=True,
    )


def output_dir_for(target_date: date) -> str:
    return f"社群任务{target_date:%m%d}"


def output_path_for(config: SchedulerConfig, target_date: date) -> Path:
    return config.root / output_dir_for(target_date)


def date_range(start: date, end: date) -> Iterable[date]:
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def load_state(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {"schema_version": 2, "dates": {}, "task_start_dates": {}}
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        backup = path.with_suffix(f".bad-{datetime.now():%Y%m%d%H%M%S}.json")
        path.replace(backup)
        return {"schema_version": 2, "dates": {}, "task_start_dates": {}}
    if not isinstance(state, dict):
        return {"schema_version": 2, "dates": {}, "task_start_dates": {}}
    state.setdefault("schema_version", 2)
    state.setdefault("dates", {})
    state.setdefault("task_start_dates", {})
    return state


def save_state(path: Path, state: Dict[str, Any]) -> None:
    state_file_io.save_json_atomic(path, state)


def write_status(config: SchedulerConfig, status: str, message: str) -> None:
    config.state_dir.mkdir(parents=True, exist_ok=True)
    safe_message = re.sub(r"(Bearer\s+)[A-Za-z0-9._~+/=-]+", r"\1<redacted>", message)
    text = f"time={now_text()}\nstatus={status}\nmessage={safe_message}\n"
    config.status_path.write_text(text, encoding="utf-8")


def parse_mmdd_folder(name: str, today: date) -> Optional[date]:
    match = re.fullmatch(r"社群任务(\d{2})(\d{2})", name)
    if not match:
        return None
    month = int(match.group(1))
    day = int(match.group(2))
    try:
        candidate = date(today.year, month, day)
    except ValueError:
        return None
    if candidate > today - timedelta(days=1):
        try:
            candidate = date(today.year - 1, month, day)
        except ValueError:
            return None
    return candidate


def has_xlsx_files(path: Path) -> bool:
    return any(child.is_file() and child.suffix.lower() == ".xlsx" for child in path.iterdir())


def has_customer_group_export_files(path: Path) -> bool:
    return any(
        child.is_file()
        and child.suffix.lower() == ".xlsx"
        and child.name.startswith("群发客户群_客户群统计")
        for child in path.iterdir()
    )


def get_date_record(state: Dict[str, Any], target_date: date) -> Dict[str, Any]:
    dates = state.setdefault("dates", {})
    iso = target_date.isoformat()
    record = dates.get(iso)
    if not isinstance(record, dict):
        record = {}
        dates[iso] = record
    record.setdefault("output_dir", output_dir_for(target_date))
    record.setdefault("tasks", {})
    return record


def task_record(state: Dict[str, Any], target_date: date, task_id: str) -> Dict[str, Any]:
    record = get_date_record(state, target_date)
    tasks = record.setdefault("tasks", {})
    task = tasks.get(task_id)
    if not isinstance(task, dict):
        task = {}
        tasks[task_id] = task
    return task


def task_successful_dates(state: Dict[str, Any], task_id: str) -> List[date]:
    result: List[date] = []
    for iso, record in (state.get("dates") or {}).items():
        if not isinstance(record, dict):
            continue
        tasks = record.get("tasks") or {}
        task = tasks.get(task_id)
        if not isinstance(task, dict) or task.get("status") != "success":
            continue
        try:
            result.append(datetime.strptime(iso, "%Y-%m-%d").date())
        except ValueError:
            continue
    return sorted(result)


def update_last_success(state: Dict[str, Any]) -> bool:
    changed = False
    last_by_task: Dict[str, str] = {}
    for task in EXPORT_TASKS:
        dates = task_successful_dates(state, task.task_id)
        if dates:
            last_by_task[task.task_id] = dates[-1].isoformat()
    if state.get("last_successful_dates") != last_by_task:
        state["last_successful_dates"] = last_by_task
        changed = True
    super_last = last_by_task.get(SUPER_GROUP_TASK_ID)
    if super_last and state.get("last_successful_date") != super_last:
        state["last_successful_date"] = super_last
        changed = True
    return changed


def migrate_state(state: Dict[str, Any]) -> bool:
    changed = False
    dates = state.setdefault("dates", {})
    for iso, record in list(dates.items()):
        if not isinstance(record, dict):
            dates[iso] = {"output_dir": f"社群任务{iso[5:7]}{iso[8:10]}", "tasks": {}}
            changed = True
            continue
        record.setdefault("output_dir", record.get("output_dir") or f"社群任务{iso[5:7]}{iso[8:10]}")
        tasks = record.setdefault("tasks", {})
        if record.get("status") == "success" and SUPER_GROUP_TASK_ID not in tasks:
            tasks[SUPER_GROUP_TASK_ID] = {
                "status": "success",
                "completed_at": record.get("completed_at") or now_text(),
                "output_dir": record.get("output_dir"),
                "source": record.get("source") or "legacy_state",
            }
            changed = True
        reach_task = tasks.get(REACH_SUMMARY_TASK_ID)
        if isinstance(reach_task, dict) and reach_task.get("label") == "统计触达客户":
            reach_task["label"] = "群发客户及朋友圈触达人数"
            changed = True

    if state.get("schema_version") != 2:
        state["schema_version"] = 2
        changed = True
    state.setdefault("task_start_dates", {})
    if update_last_success(state):
        changed = True
    return changed


def initialize_state_from_existing_dirs(state: Dict[str, Any], root: Path, today: date) -> bool:
    if task_successful_dates(state, SUPER_GROUP_TASK_ID):
        return False

    changed = False
    for child in root.iterdir():
        if not child.is_dir():
            continue
        target_date = parse_mmdd_folder(child.name, today)
        if not target_date or not has_xlsx_files(child):
            continue
        get_date_record(state, target_date)["output_dir"] = child.name
        task_record(state, target_date, SUPER_GROUP_TASK_ID).update(
            {
                "status": "success",
                "completed_at": now_text(),
                "output_dir": child.name,
                "source": "existing_folder",
            }
        )
        changed = True

    if changed:
        update_last_success(state)
    return changed


def initialize_reach_state_from_existing_docx(state: Dict[str, Any], root: Path) -> bool:
    docx_path = root / reach_summary_exporter.DEFAULT_OUTPUT_DOCX
    if not docx_path.exists():
        return False

    changed = False
    for target_date in reach_summary_exporter.extract_docx_dates(docx_path):
        if is_task_success(state, target_date, REACH_SUMMARY_TASK_ID):
            continue
        task_record(state, target_date, REACH_SUMMARY_TASK_ID).update(
            {
                "status": "success",
                "completed_at": now_text(),
                "output_docx": docx_path.name,
                "label": "群发客户及朋友圈触达人数",
                "source": "existing_docx",
            }
        )
        changed = True

    if changed:
        update_last_success(state)
    return changed


def initialize_customer_group_state_from_existing_files(
    state: Dict[str, Any], root: Path, today: date
) -> bool:
    if task_successful_dates(state, CUSTOMER_GROUP_TASK_ID):
        return False

    changed = False
    for child in root.iterdir():
        if not child.is_dir():
            continue
        target_date = parse_mmdd_folder(child.name, today)
        if not target_date or not has_customer_group_export_files(child):
            continue
        get_date_record(state, target_date)["output_dir"] = child.name
        task_record(state, target_date, CUSTOMER_GROUP_TASK_ID).update(
            {
                "status": "success",
                "completed_at": now_text(),
                "output_dir": child.name,
                "label": "群发客户群导出",
                "source": "existing_folder",
            }
        )
        changed = True

    if changed:
        update_last_success(state)
    return changed


def initialize_reach_excel_state_from_existing_workbook(state: Dict[str, Any], root: Path) -> bool:
    xlsx_path = root / reach_excel_exporter.DEFAULT_OUTPUT_NAME
    if not xlsx_path.exists():
        return False

    changed = False
    for target_date in reach_excel_exporter.reach_summary_dates(xlsx_path):
        if is_task_success(state, target_date, REACH_EXCEL_SUMMARY_TASK_ID):
            continue
        task_record(state, target_date, REACH_EXCEL_SUMMARY_TASK_ID).update(
            {
                "status": "success",
                "completed_at": now_text(),
                "output_xlsx": xlsx_path.name,
                "label": "触达人数汇总",
                "source": "existing_xlsx",
            }
        )
        changed = True

    for target_date in reach_excel_exporter.store_group_dates(xlsx_path):
        if is_task_success(state, target_date, STORE_GROUP_REACH_TASK_ID):
            continue
        task_record(state, target_date, STORE_GROUP_REACH_TASK_ID).update(
            {
                "status": "success",
                "completed_at": now_text(),
                "output_xlsx": xlsx_path.name,
                "label": "门店分组触达人数",
                "source": "existing_xlsx",
            }
        )
        changed = True

    if changed:
        update_last_success(state)
    return changed


def initialize_feishu_sync_state_from_online(state: Dict[str, Any], config: SchedulerConfig) -> bool:
    if task_successful_dates(state, FEISHU_SYNC_TASK_ID):
        return False

    xlsx_path = config.root / reach_excel_exporter.DEFAULT_OUTPUT_NAME
    if not xlsx_path.exists():
        return False

    try:
        complete_through = feishu_sync.complete_through_date(config.env_path)
    except feishu_sync.FeishuSyncError as exc:
        print(f"Feishu sync state initialization skipped: {exc}", file=sys.stderr)
        return False
    if complete_through is None:
        return False

    changed = False
    for target_date in feishu_sync.workbook_dates_through(xlsx_path, complete_through):
        if is_task_success(state, target_date, FEISHU_SYNC_TASK_ID):
            continue
        task_record(state, target_date, FEISHU_SYNC_TASK_ID).update(
            {
                "status": "success",
                "completed_at": now_text(),
                "output_xlsx": xlsx_path.name,
                "label": "同步在线表",
                "source": "existing_feishu",
            }
        )
        changed = True

    if changed:
        update_last_success(state)
    return changed


def ensure_task_start_dates(state: Dict[str, Any], today: date) -> bool:
    state.setdefault("task_start_dates", {})
    return False


def is_task_success(state: Dict[str, Any], target_date: date, task_id: str) -> bool:
    record = (state.get("dates") or {}).get(target_date.isoformat())
    if not isinstance(record, dict):
        return False
    tasks = record.get("tasks") or {}
    task = tasks.get(task_id)
    return isinstance(task, dict) and task.get("status") == "success"


def first_pending_date_for_task(
    state: Dict[str, Any],
    task: ExportTask,
    today: date,
    lookback_days: int,
) -> date:
    yesterday = today - timedelta(days=1)
    successes = task_successful_dates(state, task.task_id)
    if successes:
        return successes[-1] + timedelta(days=1)
    return yesterday - timedelta(days=lookback_days - 1)


def pending_task_runs(
    state: Dict[str, Any],
    today: date,
    lookback_days: int,
    start_date: Optional[date] = None,
) -> List[Tuple[date, ExportTask]]:
    yesterday = today - timedelta(days=1)
    if yesterday < date(2000, 1, 1):
        return []

    pending: List[Tuple[date, ExportTask]] = []
    for task in EXPORT_TASKS:
        start = start_date or first_pending_date_for_task(state, task, today, lookback_days)
        if start > yesterday:
            continue
        for target_date in date_range(start, yesterday):
            if not is_task_success(state, target_date, task.task_id):
                pending.append((target_date, task))

    order = {task.task_id: index for index, task in enumerate(EXPORT_TASKS)}
    return sorted(pending, key=lambda item: (item[0], order[item[1].task_id]))


def mark_task_success(state: Dict[str, Any], target_date: date, task: ExportTask) -> None:
    record = {
        "status": "success",
        "completed_at": now_text(),
        "output_dir": output_dir_for(target_date),
        "label": task.label,
    }
    if task.task_id == REACH_SUMMARY_TASK_ID:
        record["output_docx"] = reach_summary_exporter.DEFAULT_OUTPUT_DOCX
    if task.task_id in {REACH_EXCEL_SUMMARY_TASK_ID, STORE_GROUP_REACH_TASK_ID, FEISHU_SYNC_TASK_ID}:
        record["output_xlsx"] = reach_excel_exporter.DEFAULT_OUTPUT_NAME
    task_record(state, target_date, task.task_id).update(record)
    update_last_success(state)


def mark_task_failure(state: Dict[str, Any], target_date: date, task: ExportTask, message: str) -> None:
    safe_message = re.sub(r"(Bearer\s+)[A-Za-z0-9._~+/=-]+", r"\1<redacted>", message)
    record = {
        "status": "failed",
        "failed_at": now_text(),
        "output_dir": output_dir_for(target_date),
        "label": task.label,
        "message": safe_message,
    }
    if task.task_id == REACH_SUMMARY_TASK_ID:
        record["output_docx"] = reach_summary_exporter.DEFAULT_OUTPUT_DOCX
    if task.task_id in {REACH_EXCEL_SUMMARY_TASK_ID, STORE_GROUP_REACH_TASK_ID, FEISHU_SYNC_TASK_ID}:
        record["output_xlsx"] = reach_excel_exporter.DEFAULT_OUTPUT_NAME
    task_record(state, target_date, task.task_id).update(record)


def is_auth_error(exc: BaseException) -> bool:
    text = str(exc)
    return bool(
        re.search(r"\b401\b", text)
        or re.search(r"unauthorized|not\s*login|login\s*expired", text, flags=re.I)
        or "未登录" in text
        or "登录" in text and "过期" in text
    )


def is_chrome_debug_port_error(exc: BaseException) -> bool:
    text = f"{type(exc).__name__}: {exc}"
    lowered = text.lower()
    return bool(
        "winerror 10061" in lowered
        or "127.0.0.1:9222" in lowered
        or "127.0.0.1:9333" in lowered
        or "localhost:9222" in lowered
        or "localhost:9333" in lowered
        or "chrome debug port" in lowered
        or ("debug port" in lowered and "9222" in lowered)
        or ("debug port" in lowered and "9333" in lowered)
        or ("connection refused" in lowered and ("127.0.0.1" in lowered or "localhost" in lowered))
        or ("actively refused" in lowered and ("127.0.0.1" in lowered or "localhost" in lowered))
        or "由于目标计算机积极拒绝" in text
    )


def should_refresh_login(exc: BaseException) -> bool:
    return is_auth_error(exc) or is_chrome_debug_port_error(exc)


def run_super_group_export(target_date: date, config: SchedulerConfig) -> None:
    output_dir = output_path_for(config, target_date)
    print(
        f"\n=== {target_date:%Y-%m-%d} | 超级群发未送达 -> {output_dir} ===",
        flush=True,
    )
    exporter.main(
        [
            "--date",
            target_date.isoformat(),
            "--output-dir",
            str(output_dir),
        ]
    )


def run_chat_analysis_export(target_date: date, config: SchedulerConfig) -> None:
    output_dir = output_path_for(config, target_date)
    print(
        f"\n=== {target_date:%Y-%m-%d} | 客户群分析-按群聊 -> {output_dir} ===",
        flush=True,
    )
    chat_analysis_exporter.main(
        [
            "--date",
            target_date.isoformat(),
            "--output-dir",
            str(output_dir),
        ]
    )


def run_reach_summary_export(target_date: date, config: SchedulerConfig) -> None:
    print(
        f"\n=== {target_date:%Y-%m-%d} | 群发客户及朋友圈触达人数 -> {reach_summary_exporter.DEFAULT_OUTPUT_DOCX} ===",
        flush=True,
    )
    reach_summary_exporter.main(
        [
            "--date",
            target_date.isoformat(),
            "--output-docx",
            str(config.root / reach_summary_exporter.DEFAULT_OUTPUT_DOCX),
        ]
    )


def run_reach_excel_summary_export(target_date: date, config: SchedulerConfig) -> None:
    output_xlsx = config.root / reach_excel_exporter.DEFAULT_OUTPUT_NAME
    reach_docx = config.root / reach_summary_exporter.DEFAULT_OUTPUT_DOCX
    print(
        f"\n=== {target_date:%Y-%m-%d} | 触达人数汇总 -> {output_xlsx} ===",
        flush=True,
    )
    reach_excel_exporter.update_reach_summary_sheet(
        source_root=config.root,
        output_xlsx=output_xlsx,
        reach_docx=reach_docx,
        target_dates=[target_date],
        recent_days=config.catchup_lookback_days,
        today=target_date + timedelta(days=1),
    )


def run_store_group_reach_export(target_date: date, config: SchedulerConfig) -> None:
    output_xlsx = config.root / reach_excel_exporter.DEFAULT_OUTPUT_NAME
    print(
        f"\n=== {target_date:%Y-%m-%d} | 门店分组触达人数 -> {output_xlsx} ===",
        flush=True,
    )
    reach_excel_exporter.update_store_group_sheet(
        source_root=config.root,
        output_xlsx=output_xlsx,
        target_dates=[target_date],
        recent_days=config.catchup_lookback_days,
        today=target_date + timedelta(days=1),
    )


def run_feishu_sync_export(target_date: date, config: SchedulerConfig) -> None:
    output_xlsx = config.root / reach_excel_exporter.DEFAULT_OUTPUT_NAME
    print(
        f"\n=== {target_date:%Y-%m-%d} | 同步在线表 -> {output_xlsx} ===",
        flush=True,
    )
    feishu_sync.sync_date(
        workbook_path=output_xlsx,
        env_path=config.env_path,
        target_date=target_date,
    )


def run_customer_group_export(target_date: date, config: SchedulerConfig) -> None:
    output_dir = output_path_for(config, target_date)
    print(
        f"\n=== {target_date:%Y-%m-%d} | 群发客户群导出 -> {output_dir} ===",
        flush=True,
    )
    customer_group_exporter.main(
        [
            "--date",
            target_date.isoformat(),
            "--output-dir",
            str(output_dir),
        ]
    )


def run_task_export(task: ExportTask, target_date: date, config: SchedulerConfig) -> None:
    if task.task_id == SUPER_GROUP_TASK_ID:
        run_super_group_export(target_date, config)
        return
    if task.task_id == CHAT_ANALYSIS_TASK_ID:
        run_chat_analysis_export(target_date, config)
        return
    if task.task_id == REACH_SUMMARY_TASK_ID:
        run_reach_summary_export(target_date, config)
        return
    if task.task_id == CUSTOMER_GROUP_TASK_ID:
        run_customer_group_export(target_date, config)
        return
    if task.task_id == REACH_EXCEL_SUMMARY_TASK_ID:
        run_reach_excel_summary_export(target_date, config)
        return
    if task.task_id == STORE_GROUP_REACH_TASK_ID:
        run_store_group_reach_export(target_date, config)
        return
    if task.task_id == FEISHU_SYNC_TASK_ID:
        run_feishu_sync_export(target_date, config)
        return
    raise RuntimeError(f"Unknown export task: {task.task_id}")


def heartbeat_loop(stop_event: threading.Event, target_date: date, task: ExportTask) -> None:
    while not stop_event.wait(30):
        emit_event(
            "RUN_HEARTBEAT",
            date=target_date.isoformat(),
            task_id=task.task_id,
            label=task.label,
            message="Task is still running.",
        )


def run_task_export_with_heartbeat(task: ExportTask, target_date: date, config: SchedulerConfig) -> None:
    stop_event = threading.Event()
    thread = threading.Thread(
        target=heartbeat_loop,
        args=(stop_event, target_date, task),
        daemon=True,
    )
    thread.start()
    try:
        run_task_export(task, target_date, config)
    finally:
        stop_event.set()
        thread.join(timeout=1)


def chrome_candidates() -> List[str]:
    candidates: List[str] = []
    for env_key in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
        base = os.environ.get(env_key)
        if not base:
            continue
        candidates.append(str(Path(base) / "Google" / "Chrome" / "Application" / "chrome.exe"))
    candidates.extend(["chrome.exe", "msedge.exe"])
    return candidates


def http_json(url: str, timeout: float = 3.0, method: str = "GET") -> Any:
    request = Request(url, method=method)
    with urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def is_debug_port_ready(port: int) -> bool:
    try:
        http_json(f"http://127.0.0.1:{port}/json/version", timeout=1.5)
        return True
    except (OSError, URLError, HTTPError, json.JSONDecodeError):
        return False


def launch_chrome(config: SchedulerConfig) -> None:
    if is_debug_port_ready(config.chrome_debug_port):
        return

    config.chrome_profile_dir.mkdir(parents=True, exist_ok=True)
    args_tail = [
        f"--remote-debugging-port={config.chrome_debug_port}",
        "--remote-debugging-address=127.0.0.1",
        f"--user-data-dir={config.chrome_profile_dir}",
        "--new-window",
        LOGIN_URL,
    ]
    last_error: Optional[BaseException] = None
    for candidate in chrome_candidates():
        try:
            subprocess.Popen(
                [candidate, *args_tail],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                close_fds=True,
            )
            break
        except OSError as exc:
            last_error = exc
    else:
        raise LoginRefreshError(f"Cannot start Chrome: {last_error}")

    deadline = time.monotonic() + 20
    while time.monotonic() < deadline:
        if is_debug_port_ready(config.chrome_debug_port):
            return
        time.sleep(0.5)
    raise LoginRefreshError("Chrome did not expose the local debug port in time")


def list_cdp_targets(port: int) -> List[Dict[str, Any]]:
    targets = http_json(f"http://127.0.0.1:{port}/json", timeout=3)
    if not isinstance(targets, list):
        return []
    return [target for target in targets if isinstance(target, dict)]


def ensure_login_target(port: int, prefer_new: bool = False) -> str:
    encoded_url = quote(LOGIN_URL, safe="")
    if prefer_new:
        try:
            target = http_json(f"http://127.0.0.1:{port}/json/new?{encoded_url}", timeout=3, method="PUT")
        except HTTPError:
            target = http_json(f"http://127.0.0.1:{port}/json/new?{encoded_url}", timeout=3)
        websocket_url = target.get("webSocketDebuggerUrl") if isinstance(target, dict) else None
        if websocket_url:
            return str(websocket_url)

    targets = list_cdp_targets(port)
    for target in targets:
        if target.get("type") == "page" and SCRM_HOST in str(target.get("url", "")):
            websocket_url = target.get("webSocketDebuggerUrl")
            if websocket_url:
                return str(websocket_url)

    try:
        target = http_json(f"http://127.0.0.1:{port}/json/new?{encoded_url}", timeout=3, method="PUT")
    except HTTPError:
        target = http_json(f"http://127.0.0.1:{port}/json/new?{encoded_url}", timeout=3)
    websocket_url = target.get("webSocketDebuggerUrl") if isinstance(target, dict) else None
    if not websocket_url:
        raise LoginRefreshError("Could not create a Chrome login tab")
    return str(websocket_url)


def evaluate_page_auth(websocket_url: str) -> Dict[str, Any]:
    expression = r"""
(() => {
  const readStorage = (storage) => {
    const values = {};
    for (let index = 0; index < storage.length; index += 1) {
      const key = storage.key(index);
      values[key] = storage.getItem(key);
    }
    return values;
  };
  return {
    href: location.href,
    cookie: document.cookie || "",
    localStorage: readStorage(window.localStorage),
    sessionStorage: readStorage(window.sessionStorage),
    userAgent: navigator.userAgent
  };
})()
"""
    with CdpClient(websocket_url) as client:
        result = client.call(
            "Runtime.evaluate",
            {
                "expression": expression,
                "returnByValue": True,
                "awaitPromise": False,
            },
            timeout=10,
        )
        cookie_result = client.call(
            "Network.getCookies",
            {"urls": [f"https://{SCRM_HOST}/"]},
            timeout=10,
        )
    value = ((result.get("result") or {}).get("value")) or {}
    if not isinstance(value, dict):
        value = {}

    cookies = cookie_result.get("cookies") or []
    if isinstance(cookies, list) and cookies:
        parts = []
        for cookie in cookies:
            if not isinstance(cookie, dict):
                continue
            name = str(cookie.get("name") or "").strip()
            cookie_value = str(cookie.get("value") or "")
            if name:
                parts.append(f"{name}={cookie_value}")
        if parts:
            value["cookie"] = "; ".join(parts)

    return value


def maybe_json(value: str) -> Any:
    value = value.strip()
    if not value or value[0] not in "[{\"":
        return value
    try:
        return json.loads(value)
    except json.JSONDecodeError:
        return value


def clean_token(value: Any) -> Optional[str]:
    if value is None:
        return None
    if not isinstance(value, str):
        return None
    token = value.strip().strip('"').strip("'")
    if not token or token.lower() in {"null", "undefined", "none"}:
        return None
    if token.startswith("{") or token.startswith("["):
        nested = extract_token(maybe_json(token))
        if nested:
            return nested
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    if len(token) < 20:
        return None
    if any(char in token for char in "\r\n;"):
        return None
    return token


def extract_token(value: Any, key_hint: str = "") -> Optional[str]:
    preferred_keys = {
        "current-token",
        "currenttoken",
        "token",
        "accesstoken",
        "access_token",
        "authorization",
    }
    if isinstance(value, str):
        parsed = maybe_json(value)
        if parsed is not value:
            return extract_token(parsed, key_hint)
        if key_hint.lower().replace("-", "").replace("_", "") in preferred_keys or "token" in key_hint.lower():
            return clean_token(value)
        return None
    if isinstance(value, dict):
        for key, nested in value.items():
            normalized = str(key).lower().replace("-", "").replace("_", "")
            if normalized in preferred_keys:
                token = extract_token(nested, str(key)) or clean_token(nested)
                if token:
                    return token
        for key, nested in value.items():
            if "token" in str(key).lower() or "authorization" in str(key).lower():
                token = extract_token(nested, str(key)) or clean_token(nested)
                if token:
                    return token
        for key, nested in value.items():
            token = extract_token(nested, str(key))
            if token:
                return token
    if isinstance(value, list):
        for nested in value:
            token = extract_token(nested, key_hint)
            if token:
                return token
    return None


def parse_cookie_value(cookie: str, key: str) -> str:
    for part in cookie.split(";"):
        if "=" not in part:
            continue
        item_key, value = part.strip().split("=", 1)
        if item_key == key:
            return unquote(value)
    return ""


def auth_updates_from_page(page: Dict[str, Any]) -> Dict[str, str]:
    cookie = str(page.get("cookie") or "").strip()
    local_storage = page.get("localStorage") if isinstance(page.get("localStorage"), dict) else {}
    session_storage = page.get("sessionStorage") if isinstance(page.get("sessionStorage"), dict) else {}

    token = extract_token(session_storage) or extract_token(local_storage) or clean_token(parse_cookie_value(cookie, "token"))
    updates: Dict[str, str] = {}
    if token:
        updates["SCRM_AUTHORIZATION"] = token if token.startswith("Bearer ") else f"Bearer {token}"
        updates["SCRM_TOKEN"] = token
    if cookie:
        updates["SCRM_COOKIE"] = cookie

    tenant_id = parse_cookie_value(cookie, "corpId")
    if tenant_id:
        updates["SCRM_TENANT_ID"] = tenant_id

    device_id = str(local_storage.get("__ws_device_id") or "").strip()
    if device_id:
        updates["SCRM_DEVICE_ID"] = device_id

    user_agent = str(page.get("userAgent") or "").strip()
    if user_agent:
        updates["SCRM_USER_AGENT"] = user_agent

    if "SCRM_AUTHORIZATION" not in updates and "SCRM_COOKIE" not in updates:
        return {}
    return updates


def validate_page_login(websocket_url: str) -> Tuple[bool, str]:
    target_date = (date.today() - timedelta(days=1)).isoformat()
    expression = f"""
(async () => {{
  const getCookie = (name) => {{
    const prefix = name + '=';
    for (const part of document.cookie.split(';')) {{
      const item = part.trim();
      if (item.startsWith(prefix)) return decodeURIComponent(item.slice(prefix.length));
    }}
    return '';
  }};
  const token = sessionStorage.getItem('current-token') || localStorage.getItem('current-token') || '';
  const corpId = getCookie('corpId');
  const context = JSON.stringify({{tenantId: corpId}});
  const headers = {{
    Accept: 'application/json, text/plain, */*',
    'Content-Type': 'application/json',
    'x-admin-header': '1',
    'x-header-host': 'scrm.cotticoffee.cc',
    'x-clientType-header': 'pc',
    'x-requestMsgId-header': 'login-check-' + Date.now(),
    'CONTEXT-JSON': context,
    'CONTEXT_JSON': context
  }};
  if (token) headers.Authorization = token.startsWith('Bearer ') ? token : 'Bearer ' + token;
  const payload = {{
    currentIndex: 1,
    pageSize: 1,
    chatIdList: [],
    dateType: 4,
    startDate: {json.dumps(target_date)},
    endDate: {json.dumps(target_date)},
    sortField: '',
    sortType: ''
  }};
  try {{
    const res = await fetch('/bff/customer/private/pc/chatSummary/detail/chat', {{
      method: 'POST',
      credentials: 'include',
      headers,
      body: JSON.stringify(payload)
    }});
    const text = await res.text();
    let data = null;
    try {{ data = JSON.parse(text); }} catch (_error) {{}}
    const code = data && data.code !== undefined && data.code !== null ? String(data.code) : '';
    const message = data ? String(data.msg || data.message || '') : text.slice(0, 120);
    const success = data && (data.success === true || code === '00000' || code === '200');
    return {{
      ok: res.status < 400 && success,
      status: res.status,
      code,
      message,
      href: location.href
    }};
  }} catch (error) {{
    return {{ok: false, status: 0, code: '', message: String(error), href: location.href}};
  }}
}})()
"""
    with CdpClient(websocket_url) as client:
        result = client.call(
            "Runtime.evaluate",
            {"expression": expression, "returnByValue": True, "awaitPromise": True},
            timeout=15,
        )
    value = ((result.get("result") or {}).get("value")) or {}
    if not isinstance(value, dict):
        return False, "login check did not return a result"
    if value.get("ok") is True:
        return True, "ok"
    status = value.get("status") or 0
    code = value.get("code") or ""
    message = str(value.get("message") or "").strip()
    if len(message) > 80:
        message = message[:80] + "..."
    return False, f"status={status}, code={code}, message={message or 'not ready'}"


def normalized_auth_value(key: str, value: str) -> str:
    value = str(value or "").strip()
    if key in {"SCRM_AUTHORIZATION", "SCRM_TOKEN"} and value.lower().startswith("bearer "):
        value = value[7:].strip()
    return value


def has_new_login_state(old_env: Dict[str, str], updates: Dict[str, str]) -> bool:
    for key in ("SCRM_AUTHORIZATION", "SCRM_TOKEN", "SCRM_COOKIE"):
        new_value = normalized_auth_value(key, updates.get(key, ""))
        if not new_value:
            continue
        old_value = normalized_auth_value(key, old_env.get(key, ""))
        if not old_value or new_value != old_value:
            return True
    return False


def update_dotenv(path: Path, updates: Dict[str, str]) -> None:
    existing = path.read_text(encoding="utf-8-sig").splitlines() if path.exists() else []
    remaining = dict(updates)
    output: List[str] = []
    key_re = re.compile(r"^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=")
    for line in existing:
        match = key_re.match(line)
        if match and match.group(1) in remaining:
            key = match.group(1)
            output.append(f"{key}={remaining.pop(key)}")
        else:
            output.append(line)
    if remaining and output and output[-1].strip():
        output.append("")
    for key, value in remaining.items():
        output.append(f"{key}={value}")
    path.write_text("\n".join(output) + "\n", encoding="utf-8")


def refresh_login(config: SchedulerConfig) -> bool:
    print("SCRM login expired. Opening Chrome for QR login...", flush=True)
    old_env = exporter.load_dotenv(config.env_path)
    launch_chrome(config)
    websocket_url = ensure_login_target(config.chrome_debug_port, prefer_new=True)
    deadline = time.monotonic() + config.login_wait_minutes * 60
    next_notice = 0.0

    while time.monotonic() < deadline:
        try:
            page = evaluate_page_auth(websocket_url)
            updates = auth_updates_from_page(page)
            if updates:
                verified, reason = validate_page_login(websocket_url)
                if not verified:
                    if time.monotonic() >= next_notice:
                        if has_new_login_state(old_env, updates):
                            print(
                                f"Login state detected but not verified yet ({reason}); waiting for QR login...",
                                flush=True,
                            )
                        else:
                            print(
                                f"Existing login state is not valid yet ({reason}); waiting for QR login refresh...",
                                flush=True,
                            )
                        next_notice = time.monotonic() + 30
                    time.sleep(5)
                    continue
                update_dotenv(config.env_path, updates)
                print("Login state verified and .env updated.", flush=True)
                return True
        except (LoginRefreshError, OSError, URLError, HTTPError, json.JSONDecodeError) as exc:
            try:
                websocket_url = ensure_login_target(config.chrome_debug_port)
            except Exception:
                pass
            if time.monotonic() >= next_notice:
                print(f"Waiting for Chrome login page... ({exc})", flush=True)
                next_notice = time.monotonic() + 30

        if time.monotonic() >= next_notice:
            remaining = int(deadline - time.monotonic())
            print(f"Waiting for QR login, {max(0, remaining // 60)} minute(s) left...", flush=True)
            next_notice = time.monotonic() + 30
        time.sleep(5)

    return False


def run_scheduler(
    config: SchedulerConfig,
    today: date,
    plan_only: bool,
    start_date: Optional[date] = None,
) -> int:
    state = load_state(config.state_path)
    migrated = migrate_state(state)
    initialized = initialize_state_from_existing_dirs(state, config.root, today)
    reach_initialized = initialize_reach_state_from_existing_docx(state, config.root)
    customer_group_initialized = initialize_customer_group_state_from_existing_files(
        state, config.root, today
    )
    reach_excel_initialized = initialize_reach_excel_state_from_existing_workbook(state, config.root)
    feishu_sync_initialized = initialize_feishu_sync_state_from_online(state, config)
    starts_initialized = ensure_task_start_dates(state, today)
    effective_start_date = start_date or config.global_start_date
    pending = pending_task_runs(state, today, config.catchup_lookback_days, effective_start_date)
    pending_text = ", ".join(f"{item_date.isoformat()}:{task.task_id}" for item_date, task in pending)

    print(f"Today: {today:%Y-%m-%d}")
    print(f"Start date override: {start_date:%Y-%m-%d}" if start_date else "Start date override: (none)")
    print(
        f"Global start date: {config.global_start_date:%Y-%m-%d}"
        if config.global_start_date
        else "Global start date: (none)"
    )
    print(f"State migrated: {migrated}")
    print(f"State initialized from existing folders: {initialized}")
    print(f"Reach summary initialized from existing docx: {reach_initialized}")
    print(f"Customer group initialized from existing files: {customer_group_initialized}")
    print(f"Reach Excel initialized from existing workbook: {reach_excel_initialized}")
    print(f"Feishu sync initialized from online sheets: {feishu_sync_initialized}")
    print(f"Task start dates initialized: {starts_initialized}")
    print(f"Pending task runs: {pending_text or '(none)'}")
    emit_event("RUN_PENDING", pending=pending_text, today=today.isoformat())

    if plan_only:
        return 0

    if (
        migrated
        or initialized
        or reach_initialized
        or customer_group_initialized
        or reach_excel_initialized
        or feishu_sync_initialized
        or starts_initialized
    ):
        save_state(config.state_path, state)

    if not pending:
        write_status(config, "OK", "No pending export task runs.")
        emit_event("RUN_OK", message="No pending export task runs.")
        return 0

    write_status(config, "RUNNING", f"Pending task runs: {pending_text}")
    login_refreshed = False
    for target_date, task in pending:
        attempts = 0
        while True:
            attempts += 1
            try:
                write_status(
                    config,
                    "RUNNING",
                    f"Exporting {target_date:%Y-%m-%d} [{task.task_id}] {task.label}.",
                )
                emit_event(
                    "TASK_START",
                    date=target_date.isoformat(),
                    task_id=task.task_id,
                    label=task.label,
                    output_dir=str(output_path_for(config, target_date)),
                )
                run_task_export_with_heartbeat(task, target_date, config)
                state = load_state(config.state_path)
                mark_task_success(state, target_date, task)
                save_state(config.state_path, state)
                emit_event(
                    "TASK_OK",
                    date=target_date.isoformat(),
                    task_id=task.task_id,
                    label=task.label,
                    output_dir=str(output_path_for(config, target_date)),
                )
                break
            except (
                exporter.ScrmError,
                chat_analysis_exporter.ChatAnalysisError,
                customer_group_exporter.GroupSendCustomerGroupError,
                reach_summary_exporter.ReachSummaryError,
                reach_excel_exporter.ReachSummaryError,
                feishu_sync.FeishuSyncError,
            ) as exc:
                if should_refresh_login(exc) and not login_refreshed:
                    write_status(
                        config,
                        "WAITING_LOGIN",
                        f"Waiting for QR login before exporting {target_date:%Y-%m-%d} [{task.task_id}].",
                    )
                    emit_event(
                        "WAITING_LOGIN",
                        date=target_date.isoformat(),
                        task_id=task.task_id,
                        label=task.label,
                    )
                    if not refresh_login(config):
                        state = load_state(config.state_path)
                        mark_task_failure(state, target_date, task, "Login timed out.")
                        save_state(config.state_path, state)
                        write_status(
                            config,
                            "FAILED",
                            f"Login timed out before exporting {target_date:%Y-%m-%d} [{task.task_id}].",
                        )
                        emit_event(
                            "TASK_FAILED",
                            date=target_date.isoformat(),
                            task_id=task.task_id,
                            label=task.label,
                            message="Login timed out.",
                        )
                        return 1
                    login_refreshed = True
                    continue
                state = load_state(config.state_path)
                mark_task_failure(state, target_date, task, str(exc))
                save_state(config.state_path, state)
                write_status(config, "FAILED", f"{target_date:%Y-%m-%d} [{task.task_id}]: {exc}")
                emit_event(
                    "TASK_FAILED",
                    date=target_date.isoformat(),
                    task_id=task.task_id,
                    label=task.label,
                    message=str(exc),
                )
                print(f"Export failed for {target_date:%Y-%m-%d} [{task.task_id}]: {exc}", file=sys.stderr)
                return 1
            except Exception as exc:
                if should_refresh_login(exc) and not login_refreshed:
                    write_status(
                        config,
                        "WAITING_LOGIN",
                        f"Waiting for QR login before exporting {target_date:%Y-%m-%d} [{task.task_id}].",
                    )
                    emit_event(
                        "WAITING_LOGIN",
                        date=target_date.isoformat(),
                        task_id=task.task_id,
                        label=task.label,
                    )
                    if not refresh_login(config):
                        state = load_state(config.state_path)
                        mark_task_failure(state, target_date, task, "Login timed out.")
                        save_state(config.state_path, state)
                        write_status(
                            config,
                            "FAILED",
                            f"Login timed out before exporting {target_date:%Y-%m-%d} [{task.task_id}].",
                        )
                        emit_event(
                            "TASK_FAILED",
                            date=target_date.isoformat(),
                            task_id=task.task_id,
                            label=task.label,
                            message="Login timed out.",
                        )
                        return 1
                    login_refreshed = True
                    continue
                state = load_state(config.state_path)
                mark_task_failure(state, target_date, task, f"{type(exc).__name__}: {exc}")
                save_state(config.state_path, state)
                write_status(config, "FAILED", f"{target_date:%Y-%m-%d} [{task.task_id}]: {type(exc).__name__}: {exc}")
                emit_event(
                    "TASK_FAILED",
                    date=target_date.isoformat(),
                    task_id=task.task_id,
                    label=task.label,
                    message=f"{type(exc).__name__}: {exc}",
                )
                print(f"Export failed for {target_date:%Y-%m-%d} [{task.task_id}]: {exc}", file=sys.stderr)
                return 1

            if attempts > 2:
                state = load_state(config.state_path)
                mark_task_failure(state, target_date, task, "Too many attempts.")
                save_state(config.state_path, state)
                write_status(config, "FAILED", f"Too many attempts for {target_date:%Y-%m-%d} [{task.task_id}].")
                emit_event(
                    "TASK_FAILED",
                    date=target_date.isoformat(),
                    task_id=task.task_id,
                    label=task.label,
                    message="Too many attempts.",
                )
                return 1

    last_date, _last_task = pending[-1]
    write_status(config, "OK", f"Exported task runs through {last_date:%Y-%m-%d}.")
    emit_event("RUN_OK", date=last_date.isoformat(), message=f"Exported task runs through {last_date:%Y-%m-%d}.")
    return 0


def main(argv: List[str]) -> int:
    args = parse_args(argv)
    today = parse_today(args.today)
    config_dir = runtime_paths.resolve_dir(args.config_dir, runtime_paths.default_config_dir())
    settings = app_settings.load_settings(config_dir)
    saved_data_dir = app_settings.normalize_data_dir(settings.get("data_dir"))
    data_dir = runtime_paths.resolve_dir(args.data_dir, saved_data_dir or runtime_paths.default_data_dir())
    runtime_paths.ensure_runtime_dirs(config_dir, data_dir)
    os.environ["SCRM_CONFIG_DIR"] = str(config_dir)
    os.environ["SCRM_DATA_DIR"] = str(data_dir)
    os.environ["SCRM_ENV_PATH"] = str(runtime_paths.env_path(config_dir))
    os.chdir(data_dir)
    app_settings.ensure_settings(config_dir, data_dir, today)
    config = load_config(data_dir, config_dir)

    if args.login_only:
        write_status(config, "WAITING_LOGIN", "Manual login refresh requested.")
        if refresh_login(config):
            write_status(config, "OK", "Login state refreshed.")
            return 0
        write_status(config, "FAILED", "Login refresh timed out.")
        return 1

    try:
        start_date = parse_start_date(args.start_date, today)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    return run_scheduler(config, today, args.plan_only, start_date)


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        raise SystemExit(130)
