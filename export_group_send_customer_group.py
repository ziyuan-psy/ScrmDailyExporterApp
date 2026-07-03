#!/usr/bin/env python3
"""
Export SCRM group-send customer-group undelivered statistics.

Secrets are read from .env or environment variables. Do not hard-code tokens.
"""

from __future__ import annotations

import argparse
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import export_super_group_undelivered as exporter


TASK_TYPE_CUSTOMER_GROUP = 2
EXPORT_TYPE_CUSTOMER_GROUP_STATS = 3
SEARCH_TYPE_UNDELIVERED = 0
DEFAULT_OUTPUT_DIR_PREFIX = "社群任务"
DEFAULT_EXCLUDE_KEYWORDS = ["测试", "海外", "境外"]


class GroupSendCustomerGroupError(exporter.ScrmError):
    pass


@dataclass
class Task:
    template_id: str
    name: str
    send_time: str
    status: Any
    complete_rate: Any
    raw: Dict[str, Any]


class GroupSendCustomerGroupClient:
    def __init__(self, env_file: Dict[str, str]) -> None:
        self.client = exporter.ScrmClient(env_file)

    @property
    def base_url(self) -> str:
        return self.client.base_url

    @property
    def user_agent(self) -> str:
        return self.client.user_agent

    def list_tasks(self, target_date: date, query_name: str, page_size: int) -> List[Task]:
        referer = f"{self.base_url}/customer-marketing/message/group-send-task?type=group-send-customer-group"
        start = f"{target_date:%Y-%m-%d} 00:00:00"
        end = f"{target_date:%Y-%m-%d} 23:59:59"
        page_index = 1
        tasks: List[Task] = []

        while True:
            payload = {
                "taskType": TASK_TYPE_CUSTOMER_GROUP,
                "pageSize": page_size,
                "pageIndex": page_index,
                "sendType": None,
                "status": None,
                "templateName": query_name,
                "sendTimeRange": {"startTime": start, "endTime": end},
                "createUserRange": {"userIdList": [], "deptIdList": []},
            }
            data = self.client.post_json(
                "/bff/marketing/private/pc/groupmsg/task/list",
                payload,
                referer,
            )
            page = data.get("data") or {}
            records = page.get("records") or []
            total = int(page.get("total") or len(records) or 0)
            total_pages = page.get("totalPages")

            for record in records:
                template_id = str(record.get("templateId") or "")
                if not template_id:
                    continue
                tasks.append(
                    Task(
                        template_id=template_id,
                        name=str(record.get("templateName") or ""),
                        send_time=str(record.get("sendTime") or ""),
                        status=record.get("status"),
                        complete_rate=record.get("completeRate"),
                        raw=record,
                    )
                )

            if total_pages:
                if page_index >= int(total_pages):
                    break
            elif len(tasks) >= total or not records:
                break
            page_index += 1

        return tasks

    def task_referer(self, task: Task) -> str:
        return f"{self.base_url}/customer-marketing/groupSendTask/customers-count?taskId={task.template_id}"

    def fetch_detail(self, task: Task) -> None:
        payload = {
            "templateId": task.template_id,
            "isGetSendRange": True,
            "isGetOverview": True,
            "isGetSentTip": True,
            "isGetMessagePreview": True,
        }
        self.client.post_json(
            "/bff/marketing/private/pc/groupmsg/task/detail",
            payload,
            self.task_referer(task),
        )

    def fetch_undelivered_total(self, task: Task) -> int:
        payload = {
            "pageSize": 10,
            "pageIndex": 1,
            "templateId": task.template_id,
            "searchType": SEARCH_TYPE_UNDELIVERED,
            "searchRange": {"userIdList": [], "deptIdList": []},
        }
        data = self.client.post_json(
            "/bff/marketing/private/pc/groupmsg/statistics/targetOfStaff/list",
            payload,
            self.task_referer(task),
        )
        page = data.get("data") or {}
        return int(page.get("total") or 0)

    def create_export(self, task: Task) -> str:
        payload = {
            "templateId": task.template_id,
            "exportType": EXPORT_TYPE_CUSTOMER_GROUP_STATS,
            "searchRange": {"userIdList": [], "deptIdList": []},
        }
        data = self.client.post_json(
            "/bff/marketing/private/pc/groupmsg/export/create",
            payload,
            self.task_referer(task),
        )
        down_id = str((data.get("data") or {}).get("downId") or "")
        if not down_id:
            raise GroupSendCustomerGroupError(f"No downId returned for task {task.name}")
        return down_id

    def wait_export_url(self, task: Task, down_id: str, interval: float, timeout: float) -> str:
        deadline = time.monotonic() + timeout
        last_percent = None
        last_url = ""

        while time.monotonic() < deadline:
            data = self.client.post_json(
                "/bff/marketing/private/pc/groupmsg/export/percent",
                {"down_id": down_id},
                self.task_referer(task),
            )
            progress = data.get("data") or {}
            last_percent = progress.get("downPercent")
            last_url = progress.get("downUrl") or last_url
            if progress.get("downResult") and last_url:
                break
            time.sleep(interval)
        else:
            raise GroupSendCustomerGroupError(
                f"Export timed out for {task.name}; last percent={last_percent}"
            )

        result = self.client.post_json(
            "/bff/marketing/private/pc/groupmsg/export/result",
            {"down_id": down_id},
            self.task_referer(task),
        )
        return str(result.get("data") or last_url)

    def download_file(self, download_url: str, output_dir: Path, task: Task) -> Path:
        request = Request(
            exporter.encode_download_url(download_url),
            headers={"User-Agent": self.user_agent},
            method="GET",
        )
        try:
            with urlopen(request, timeout=180) as response:
                content = response.read()
                filename = exporter.infer_filename(
                    download_url,
                    response.headers,
                    f"群发客户群_客户群统计_{task.name}.xlsx",
                )
        except HTTPError as exc:
            raise GroupSendCustomerGroupError(f"Download HTTP {exc.code} for {task.name}")
        except URLError as exc:
            raise GroupSendCustomerGroupError(
                f"Download network error for {task.name}: {exc.reason}"
            ) from exc

        output_dir.mkdir(parents=True, exist_ok=True)
        path = exporter.unique_path(output_dir / filename)
        path.write_bytes(content)
        return path


def filter_tasks(tasks: Iterable[Task], include: List[str], exclude: List[str]) -> List[Task]:
    selected: List[Task] = []
    for task in tasks:
        if include and not any(keyword in task.name for keyword in include):
            continue
        if exclude and any(keyword in task.name for keyword in exclude):
            continue
        selected.append(task)
    return selected


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export SCRM group-send customer-group undelivered statistics."
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
    return parser.parse_args(list(argv))


def resolve_target_date(value: Optional[str]) -> date:
    if value:
        return datetime.strptime(value, "%Y-%m-%d").date()
    return date.today() - timedelta(days=1)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    env_file = exporter.load_dotenv(exporter.default_env_path())
    target_date = resolve_target_date(args.date)

    query_name = args.query
    if query_name is None:
        query_name = exporter.env_value(env_file, "CUSTOMER_GROUP_TASK_NAME_QUERY")

    include_value = args.include
    if include_value is None:
        include_value = exporter.env_value(env_file, "CUSTOMER_GROUP_INCLUDE_KEYWORDS")

    exclude_value = args.exclude
    if exclude_value is None:
        exclude_value = exporter.env_value(
            env_file,
            "CUSTOMER_GROUP_EXCLUDE_KEYWORDS",
            ",".join(DEFAULT_EXCLUDE_KEYWORDS),
        )

    output_dir_value = args.output_dir or exporter.env_value(
        env_file,
        "CUSTOMER_GROUP_OUTPUT_DIR",
        f"{DEFAULT_OUTPUT_DIR_PREFIX}{target_date:%m%d}",
    )
    output_dir = Path(output_dir_value)

    poll_interval = args.poll_interval
    if poll_interval is None:
        poll_interval = float(exporter.env_value(env_file, "POLL_INTERVAL_SECONDS", "3"))

    poll_timeout = args.poll_timeout
    if poll_timeout is None:
        poll_timeout = float(exporter.env_value(env_file, "POLL_TIMEOUT_SECONDS", "600"))

    include_keywords = exporter.split_keywords(include_value)
    exclude_keywords = exporter.split_keywords(exclude_value)
    client = GroupSendCustomerGroupClient(env_file)

    print(f"Date: {target_date:%Y-%m-%d}")
    print(f"Server query: {query_name or '(empty)'}")
    print(f"Include keywords: {include_keywords or '(none)'}")
    print(f"Exclude keywords: {exclude_keywords or '(none)'}")

    all_tasks = client.list_tasks(target_date, query_name or "", args.page_size)
    matched_tasks = filter_tasks(all_tasks, include_keywords, exclude_keywords)
    print(f"Tasks found: {len(all_tasks)}, matched: {len(matched_tasks)}")

    if not matched_tasks:
        print("No matched group-send customer-group tasks. Nothing to export.")
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
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        raise SystemExit(130)
    except exporter.ScrmError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
