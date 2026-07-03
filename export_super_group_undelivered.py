#!/usr/bin/env python3
"""
Export yesterday's SCRM super-group-send undelivered customer-group stats.

Secrets are read from .env or environment variables. Do not hard-code tokens.
"""

from __future__ import annotations

import argparse
import json
import os
import random
import re
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional
from urllib.error import HTTPError, URLError
from urllib.parse import quote, unquote, urlparse, urlunparse
from urllib.request import Request, urlopen

import scrm_browser_fetch
import runtime_paths


BASE_URL = "https://scrm.cotticoffee.cc"
DEFAULT_EXCLUDE_KEYWORDS = ["测试", "海外", "境外"]
ILLEGAL_FILENAME_CHARS = r'<>:"/\|?*'
DEFAULT_REQUEST_TIMEOUT_SECONDS = 60.0
DEFAULT_REQUEST_MAX_ATTEMPTS = 3
DEFAULT_REQUEST_RETRY_BASE_DELAY_SECONDS = 5.0
DEFAULT_CHROME_DEBUG_PORT = 9222
RETRY_HTTP_STATUS_CODES = {408, 429, 500, 502, 503, 504}
RETRY_SCRM_MESSAGE_KEYWORDS = (
    "服务器正忙",
    "系统繁忙",
    "请稍后",
    "稍后再试",
    "服务异常",
    "网络异常",
    "timeout",
    "timed out",
)


class ScrmError(RuntimeError):
    pass


@dataclass
class Task:
    template_id: str
    name: str
    send_time: str
    status: Any
    complete_rate: Any
    raw: Dict[str, Any]


def load_dotenv(path: Path) -> Dict[str, str]:
    values: Dict[str, str] = {}
    if not path.exists():
        return values

    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[len("export ") :].strip()
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key:
            values[key] = value
    return values


def default_env_path() -> Path:
    configured = os.environ.get("SCRM_ENV_PATH")
    if configured:
        return Path(configured)
    return runtime_paths.env_path(runtime_paths.default_config_dir())


def env_value(env_file: Dict[str, str], key: str, default: str = "") -> str:
    return os.environ.get(key) or env_file.get(key) or default


def env_int(env_file: Dict[str, str], key: str, default: int) -> int:
    value = env_value(env_file, key, str(default)).strip()
    try:
        return max(1, int(value))
    except ValueError:
        return default


def env_float(env_file: Dict[str, str], key: str, default: float) -> float:
    value = env_value(env_file, key, str(default)).strip()
    try:
        return max(0.1, float(value))
    except ValueError:
        return default


def env_bool(env_file: Dict[str, str], key: str, default: bool) -> bool:
    value = env_value(env_file, key, "1" if default else "0").strip().lower()
    return value not in {"0", "false", "no", "off", "disabled"}


def split_keywords(value: str) -> List[str]:
    if not value:
        return []
    return [item.strip() for item in re.split(r"[,，\n;；]+", value) if item.strip()]


def parse_cookie(cookie: str) -> Dict[str, str]:
    result: Dict[str, str] = {}
    for part in cookie.split(";"):
        if "=" not in part:
            continue
        key, value = part.strip().split("=", 1)
        result[key] = value
    return result


def make_request_msg_id(device_id: str) -> str:
    suffix = f"{random.randrange(10**11):x}"[:10].upper()
    return f"{device_id}{suffix}"


def sanitize_filename(name: str, max_len: int = 160) -> str:
    cleaned = "".join("_" if c in ILLEGAL_FILENAME_CHARS else c for c in name)
    cleaned = re.sub(r"\s+", " ", cleaned).strip().strip(".")
    return (cleaned or "download")[:max_len]


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    for index in range(2, 1000):
        candidate = path.with_name(f"{stem} ({index}){suffix}")
        if not candidate.exists():
            return candidate
    raise ScrmError(f"Too many duplicate filenames for {path.name}")


def infer_filename(download_url: str, headers: Any, fallback: str) -> str:
    disposition = headers.get("Content-Disposition") or headers.get("content-disposition")
    if disposition:
        match = re.search(r"filename\*=UTF-8''([^;]+)", disposition, flags=re.I)
        if match:
            return sanitize_filename(unquote(match.group(1)))
        match = re.search(r'filename="?([^";]+)"?', disposition, flags=re.I)
        if match:
            return sanitize_filename(unquote(match.group(1)))

    path_name = Path(unquote(urlparse(download_url).path)).name
    return sanitize_filename(path_name or fallback)


def encode_download_url(download_url: str) -> str:
    parsed = urlparse(download_url)
    encoded_path = quote(unquote(parsed.path), safe="/%")
    return urlunparse(
        (
            parsed.scheme,
            parsed.netloc,
            encoded_path,
            parsed.params,
            parsed.query,
            parsed.fragment,
        )
    )


class ScrmClient:
    def __init__(self, env_file: Dict[str, str]) -> None:
        self.base_url = env_value(env_file, "SCRM_BASE_URL", BASE_URL).rstrip("/")
        self.cookie = env_value(env_file, "SCRM_COOKIE")
        self.tenant_id = env_value(env_file, "SCRM_TENANT_ID")
        self.admin_header = env_value(env_file, "SCRM_ADMIN_HEADER", "1")
        self.client_type = env_value(env_file, "SCRM_CLIENT_TYPE", "pc")
        self.device_id = env_value(env_file, "SCRM_DEVICE_ID", "CODX")
        self.user_agent = env_value(
            env_file,
            "SCRM_USER_AGENT",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36",
        )
        self.request_timeout = env_float(
            env_file,
            "SCRM_REQUEST_TIMEOUT_SECONDS",
            DEFAULT_REQUEST_TIMEOUT_SECONDS,
        )
        self.request_max_attempts = env_int(
            env_file,
            "SCRM_REQUEST_MAX_ATTEMPTS",
            DEFAULT_REQUEST_MAX_ATTEMPTS,
        )
        self.request_retry_base_delay = env_float(
            env_file,
            "SCRM_REQUEST_RETRY_BASE_DELAY_SECONDS",
            DEFAULT_REQUEST_RETRY_BASE_DELAY_SECONDS,
        )
        self.browser_fetch_fallback = env_bool(env_file, "SCRM_BROWSER_FETCH_FALLBACK", True)
        self.chrome_debug_port = env_int(env_file, "CHROME_DEBUG_PORT", DEFAULT_CHROME_DEBUG_PORT)

        cookie_map = parse_cookie(self.cookie)
        if not self.tenant_id:
            self.tenant_id = unquote(cookie_map.get("corpId", ""))

        authorization = env_value(env_file, "SCRM_AUTHORIZATION")
        token = env_value(env_file, "SCRM_TOKEN") or unquote(cookie_map.get("token", ""))
        if authorization:
            self.authorization = authorization
        elif token:
            self.authorization = token if token.startswith("Bearer ") else f"Bearer {token}"
        else:
            self.authorization = ""

        if not self.authorization and not self.cookie:
            raise ScrmError(
                "Missing auth. Put SCRM_AUTHORIZATION or SCRM_TOKEN/SCRM_COOKIE in .env."
            )

    def common_headers(self, referer: str) -> Dict[str, str]:
        context = json.dumps({"tenantId": self.tenant_id}, ensure_ascii=False)
        headers = {
            "Accept": "application/json, text/plain, */*",
            "Content-Type": "application/json",
            "User-Agent": self.user_agent,
            "Referer": referer,
            "x-admin-header": self.admin_header,
            "x-header-host": "scrm.cotticoffee.cc",
            "x-clientType-header": self.client_type,
            "x-requestMsgId-header": make_request_msg_id(self.device_id),
            "CONTEXT-JSON": context,
            "CONTEXT_JSON": context,
        }
        if self.authorization:
            headers["Authorization"] = self.authorization
        if self.cookie:
            headers["Cookie"] = self.cookie
        return headers

    def should_retry_scrm_error(self, code: str, message: str) -> bool:
        if code.isdigit() and int(code) in RETRY_HTTP_STATUS_CODES:
            return True
        message_lower = message.lower()
        return any(keyword.lower() in message_lower for keyword in RETRY_SCRM_MESSAGE_KEYWORDS)

    def should_browser_fetch_scrm_error(self, code: str, message: str) -> bool:
        return code == "401" or "服务器正忙" in message or "请稍后" in message

    def sleep_before_retry(self, path: str, attempt: int, error: ScrmError) -> None:
        delay = self.request_retry_base_delay * (2 ** (attempt - 1))
        delay += random.uniform(0, min(1.0, self.request_retry_base_delay * 0.2))
        print(
            f"  Transient SCRM request failed for {path}; "
            f"retrying in {delay:.1f}s ({attempt + 1}/{self.request_max_attempts}): {error}",
            file=sys.stderr,
            flush=True,
        )
        time.sleep(delay)

    def post_json_via_browser(
        self, path: str, payload: Dict[str, Any], referer: str
    ) -> Dict[str, Any]:
        try:
            text = scrm_browser_fetch.post_json(
                port=self.chrome_debug_port,
                path=path,
                payload=payload,
                headers=self.common_headers(referer),
                referer=referer,
                timeout=self.request_timeout,
            )
        except scrm_browser_fetch.BrowserFetchError as exc:
            raise ScrmError(f"Browser fetch fallback failed for {path}: {exc}") from exc

        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ScrmError(f"Non-JSON browser fetch response for {path}: {text[:300]}") from exc

        code = str(data.get("code", ""))
        success = data.get("success")
        if code not in ("00000", "200") and success is not True:
            message = str(data.get("msg") or data.get("message") or "unknown error")
            raise ScrmError(f"SCRM browser fetch error for {path}: code={code}, message={message}")
        return data

    def post_json(self, path: str, payload: Dict[str, Any], referer: str) -> Dict[str, Any]:
        url = f"{self.base_url}{path}"
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        last_error: Optional[ScrmError] = None

        for attempt in range(1, self.request_max_attempts + 1):
            request = Request(url, data=body, headers=self.common_headers(referer), method="POST")
            try:
                with urlopen(request, timeout=self.request_timeout) as response:
                    text = response.read().decode("utf-8")
            except HTTPError as exc:
                error = ScrmError(f"HTTP {exc.code} for {path}: {exc.read().decode('utf-8', 'ignore')}")
                if exc.code in RETRY_HTTP_STATUS_CODES and attempt < self.request_max_attempts:
                    self.sleep_before_retry(path, attempt, error)
                    last_error = error
                    continue
                raise error from exc
            except TimeoutError as exc:
                error = ScrmError(f"Network timeout for {path}: {exc}")
                if attempt < self.request_max_attempts:
                    self.sleep_before_retry(path, attempt, error)
                    last_error = error
                    continue
                raise error from exc
            except URLError as exc:
                error = ScrmError(f"Network error for {path}: {exc.reason}")
                if attempt < self.request_max_attempts:
                    self.sleep_before_retry(path, attempt, error)
                    last_error = error
                    continue
                raise error from exc

            try:
                data = json.loads(text)
            except json.JSONDecodeError as exc:
                error = ScrmError(f"Non-JSON response for {path}: {text[:300]}")
                if attempt < self.request_max_attempts:
                    self.sleep_before_retry(path, attempt, error)
                    last_error = error
                    continue
                raise error from exc

            code = str(data.get("code", ""))
            success = data.get("success")
            if code not in ("00000", "200") and success is not True:
                message = str(data.get("msg") or data.get("message") or "unknown error")
                error = ScrmError(f"SCRM error for {path}: code={code}, message={message}")
                if self.browser_fetch_fallback and self.should_browser_fetch_scrm_error(code, message):
                    print(f"  Direct SCRM request failed for {path}; trying browser fetch fallback.", flush=True)
                    return self.post_json_via_browser(path, payload, referer)
                if self.should_retry_scrm_error(code, message) and attempt < self.request_max_attempts:
                    self.sleep_before_retry(path, attempt, error)
                    last_error = error
                    continue
                if self.browser_fetch_fallback:
                    print(f"  Direct SCRM request failed for {path}; trying browser fetch fallback.", flush=True)
                    return self.post_json_via_browser(path, payload, referer)
                raise error
            return data

        if last_error:
            if self.browser_fetch_fallback:
                print(f"  Direct SCRM request failed for {path}; trying browser fetch fallback.", flush=True)
                return self.post_json_via_browser(path, payload, referer)
            raise last_error
        raise ScrmError(f"SCRM request failed for {path}")

    def list_tasks(
        self,
        target_date: date,
        query_name: str,
        page_size: int,
    ) -> List[Task]:
        referer = f"{self.base_url}/customer-marketing/message/group-send-task?type=super-group-send"
        start = f"{target_date:%Y-%m-%d} 00:00:00"
        end = f"{target_date:%Y-%m-%d} 23:59:59"
        tasks: List[Task] = []
        page_index = 1

        while True:
            payload = {
                "pageIndex": page_index,
                "pageSize": page_size,
                "templateName": query_name,
                "statusList": [],
                "taskType": 3,
                "sendTimeRange": {"startTime": start, "endTime": end},
                "createUserRange": {"userIdList": [], "deptIdList": []},
            }
            data = self.post_json("/bff/marketing/private/pc/supperMass/mass/list", payload, referer)
            page = data.get("data") or {}
            for record in page.get("records") or []:
                template_id = str(record.get("templateId") or "")
                name = str(record.get("templateName") or "")
                if not template_id:
                    continue
                tasks.append(
                    Task(
                        template_id=template_id,
                        name=name,
                        send_time=str(record.get("sendTime") or ""),
                        status=record.get("status"),
                        complete_rate=record.get("completeRate"),
                        raw=record,
                    )
                )

            total_pages = int(page.get("totalPages") or 1)
            if page_index >= total_pages:
                break
            page_index += 1

        return tasks

    def fetch_detail(self, task: Task) -> None:
        referer = f"{self.base_url}/customer-marketing/super-group-send/detailed?id={task.template_id}"
        payload = {
            "templateId": task.template_id,
            "isGetSendRange": True,
            "isGetOverview": True,
            "isGetSentTip": True,
        }
        self.post_json("/bff/marketing/private/pc/supperMass/mass/detail", payload, referer)

    def fetch_undelivered_total(self, task: Task) -> int:
        referer = f"{self.base_url}/customer-marketing/super-group-send/detailed?id={task.template_id}"
        payload = {
            "pageSize": 10,
            "pageIndex": 1,
            "templateId": task.template_id,
            "searchType": 1,
        }
        data = self.post_json(
            "/bff/marketing/private/pc/supperMass/statistics/follow/list",
            payload,
            referer,
        )
        page = data.get("data") or {}
        return int(page.get("total") or 0)

    def create_export(self, task: Task) -> str:
        referer = f"{self.base_url}/customer-marketing/super-group-send/detailed?id={task.template_id}"
        payload = {
            "templateId": task.template_id,
            "exportType": 2,
            "searchRange": {"userIdList": [], "deptIdList": []},
        }
        data = self.post_json("/bff/marketing/private/pc/supperMass/export/create", payload, referer)
        down_id = str((data.get("data") or {}).get("downId") or "")
        if not down_id:
            raise ScrmError(f"No downId returned for task {task.name}")
        return down_id

    def wait_export_url(self, task: Task, down_id: str, interval: float, timeout: float) -> str:
        referer = f"{self.base_url}/customer-marketing/super-group-send/detailed?id={task.template_id}"
        deadline = time.monotonic() + timeout
        last_percent = None
        last_url = ""

        while time.monotonic() < deadline:
            data = self.post_json(
                "/bff/marketing/private/pc/supperMass/export/percent",
                {"down_id": down_id},
                referer,
            )
            progress = data.get("data") or {}
            last_percent = progress.get("downPercent")
            last_url = progress.get("downUrl") or last_url
            if progress.get("downResult") and last_url:
                break
            time.sleep(interval)
        else:
            raise ScrmError(f"Export timed out for {task.name}; last percent={last_percent}")

        result = self.post_json(
            "/bff/marketing/private/pc/supperMass/export/result",
            {"down_id": down_id},
            referer,
        )
        return str(result.get("data") or last_url)

    def download_file(self, download_url: str, output_dir: Path, task: Task) -> Path:
        request = Request(
            encode_download_url(download_url),
            headers={"User-Agent": self.user_agent},
            method="GET",
        )
        try:
            with urlopen(request, timeout=180) as response:
                content = response.read()
                filename = infer_filename(
                    download_url,
                    response.headers,
                    f"超级群发_客户群统计_{task.name}.xlsx",
                )
        except HTTPError as exc:
            raise ScrmError(f"Download HTTP {exc.code} for {task.name}")
        except URLError as exc:
            raise ScrmError(f"Download network error for {task.name}: {exc.reason}") from exc

        output_dir.mkdir(parents=True, exist_ok=True)
        path = unique_path(output_dir / filename)
        path.write_bytes(content)
        return path


def filter_tasks(tasks: Iterable[Task], include: List[str], exclude: List[str]) -> List[Task]:
    selected: List[Task] = []
    for task in tasks:
        name = task.name
        if include and not any(keyword in name for keyword in include):
            continue
        if exclude and any(keyword in name for keyword in exclude):
            continue
        selected.append(task)
    return selected


def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export SCRM super-group-send undelivered customer-group statistics."
    )
    parser.add_argument("--date", help="Task send date, YYYY-MM-DD. Default: yesterday.")
    parser.add_argument("--query", default=None, help="Server-side task name query keyword.")
    parser.add_argument("--include", default=None, help="Local include keywords, comma separated.")
    parser.add_argument("--exclude", default=None, help="Local exclude keywords, comma separated.")
    parser.add_argument("--output-dir", default=None, help="Output directory.")
    parser.add_argument("--page-size", type=int, default=100, help="Task list page size.")
    parser.add_argument("--poll-interval", type=float, default=None, help="Export poll interval seconds.")
    parser.add_argument("--poll-timeout", type=float, default=None, help="Export poll timeout seconds.")
    parser.add_argument("--skip-empty", action="store_true", help="Skip tasks with zero undelivered records.")
    parser.add_argument("--dry-run", action="store_true", help="List matched tasks without exporting.")
    return parser.parse_args(argv)


def resolve_target_date(value: Optional[str]) -> date:
    if value:
        return datetime.strptime(value, "%Y-%m-%d").date()
    return date.today() - timedelta(days=1)


def main(argv: List[str]) -> int:
    args = parse_args(argv)
    env_file = load_dotenv(default_env_path())
    target_date = resolve_target_date(args.date)

    query_name = args.query
    if query_name is None:
        query_name = env_value(env_file, "TASK_NAME_QUERY")

    include_value = args.include
    if include_value is None:
        include_value = env_value(env_file, "INCLUDE_TASK_KEYWORDS")

    exclude_value = args.exclude
    if exclude_value is None:
        exclude_value = env_value(env_file, "EXCLUDE_TASK_KEYWORDS", ",".join(DEFAULT_EXCLUDE_KEYWORDS))

    include_keywords = split_keywords(include_value)
    exclude_keywords = split_keywords(exclude_value)

    output_dir_value = args.output_dir or env_value(
        env_file,
        "OUTPUT_DIR",
        f"社群任务{target_date:%m%d}",
    )
    output_dir = Path(output_dir_value)

    poll_interval = args.poll_interval
    if poll_interval is None:
        poll_interval = float(env_value(env_file, "POLL_INTERVAL_SECONDS", "3"))

    poll_timeout = args.poll_timeout
    if poll_timeout is None:
        poll_timeout = float(env_value(env_file, "POLL_TIMEOUT_SECONDS", "600"))

    client = ScrmClient(env_file)
    print(f"Date: {target_date:%Y-%m-%d}")
    print(f"Server query: {query_name or '(empty)'}")
    print(f"Include keywords: {include_keywords or '(none)'}")
    print(f"Exclude keywords: {exclude_keywords or '(none)'}")

    all_tasks = client.list_tasks(target_date, query_name or "", args.page_size)
    matched_tasks = filter_tasks(all_tasks, include_keywords, exclude_keywords)
    print(f"Tasks found: {len(all_tasks)}, matched: {len(matched_tasks)}")

    if not matched_tasks:
        return 0

    for index, task in enumerate(matched_tasks, start=1):
        print(f"[{index}/{len(matched_tasks)}] {task.name} | {task.send_time} | {task.complete_rate}%")
        if args.dry_run:
            continue

        client.fetch_detail(task)
        undelivered_total = client.fetch_undelivered_total(task)
        print(f"  Undelivered customer-group statistic rows: {undelivered_total}")
        if args.skip_empty and undelivered_total == 0:
            print("  Skipped because --skip-empty is set.")
            continue

        down_id = client.create_export(task)
        download_url = client.wait_export_url(task, down_id, poll_interval, poll_timeout)
        saved_path = client.download_file(download_url, output_dir, task)
        print(f"  Saved: {saved_path}")

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        raise SystemExit(130)
    except ScrmError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
