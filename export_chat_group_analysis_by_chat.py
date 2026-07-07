#!/usr/bin/env python3
"""
Export SCRM customer-group analysis detail by chat for one date.

Secrets are read from .env or environment variables. Do not hard-code tokens.
"""

from __future__ import annotations

import argparse
import copy
import re
import shutil
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from xml.etree import ElementTree as ET

import export_super_group_undelivered as exporter


BUSINESS_SCENARIO = "CUSTOMER_CHATGROUP_STATBYCHAT_EXPORT"
DEFAULT_OUTPUT_PREFIX = "客户群统计-按群聊导出明细"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)


class ChatAnalysisError(RuntimeError):
    pass


@dataclass
class ExportFile:
    name: str
    url: str


def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export SCRM customer-group analysis detail by chat."
    )
    parser.add_argument("--date", help="Target date, YYYY-MM-DD. Default: yesterday.")
    parser.add_argument("--output-dir", default=None, help="Output directory.")
    parser.add_argument("--dry-run", action="store_true", help="Fetch total only; do not export.")
    parser.add_argument("--poll-interval", type=float, default=None, help="Export poll interval seconds.")
    parser.add_argument("--poll-timeout", type=float, default=None, help="Export poll timeout seconds.")
    return parser.parse_args(argv)


def resolve_target_date(value: Optional[str]) -> date:
    if value:
        return datetime.strptime(value, "%Y-%m-%d").date()
    return date.today() - timedelta(days=1)


def output_dir_for(target_date: date) -> str:
    return f"社群任务{target_date:%m%d}"


class ChatGroupAnalysisClient:
    def __init__(self, env_file: Dict[str, str]) -> None:
        self.client = exporter.ScrmClient(env_file)
        self.base_url = self.client.base_url
        self.user_agent = self.client.user_agent

    def referer(self) -> str:
        return f"{self.base_url}/index/dashboard"

    def detail_payload(self, target_date: date, page_size: int = 10) -> Dict[str, Any]:
        return {
            "currentIndex": 1,
            "pageSize": page_size,
            "chatIdList": [],
            "dateType": 4,
            "startDate": target_date.isoformat(),
            "endDate": target_date.isoformat(),
            "sortField": "",
            "sortType": "",
        }

    def fetch_detail_total(self, target_date: date) -> int:
        data = self.client.post_json(
            "/bff/customer/private/pc/chatSummary/detail/chat",
            self.detail_payload(target_date),
            self.referer(),
        )
        page = data.get("data") or {}
        return int(page.get("total") or 0)

    def create_export(self, target_date: date) -> str:
        params = self.detail_payload(target_date)
        params["type"] = 8
        filter_param_list = [
            {"prop": key, "paramValue": value}
            for key, value in params.items()
        ]
        payload = {
            "businessScenario": BUSINESS_SCENARIO,
            "filterParamList": filter_param_list,
        }
        data = self.client.post_json(
            "/bff/export/task/create",
            payload,
            self.referer(),
        )
        task_id = str((data.get("data") or {}).get("taskId") or "")
        if not task_id:
            raise ChatAnalysisError("No taskId returned when creating chat analysis export.")
        return task_id

    def wait_export_files(self, task_id: str, interval: float, timeout: float) -> List[ExportFile]:
        payload = {"taskId": task_id, "businessScenario": BUSINESS_SCENARIO}
        deadline = time.monotonic() + timeout
        last_percent = None
        last_reported_percent = None
        has_reported = False
        next_heartbeat = 0.0
        while time.monotonic() < deadline:
            data = self.client.post_json(
                "/bff/export/task/progress",
                payload,
                self.referer(),
            )
            progress = data.get("data") or {}
            last_percent = progress.get("downPercent")
            if progress.get("downResult"):
                percent_text = str(last_percent or "").strip()
                suffix = f" ({percent_text})" if percent_text else ""
                print(f"  Export progress: completed{suffix}.", flush=True)
                break
            now = time.monotonic()
            if not has_reported or last_percent != last_reported_percent or now >= next_heartbeat:
                percent_text = str(last_percent or "").strip()
                if percent_text and not percent_text.endswith("%"):
                    percent_text = f"{percent_text}%"
                print(f"  Export progress: {percent_text or 'waiting'}", flush=True)
                last_reported_percent = last_percent
                has_reported = True
                next_heartbeat = now + 60
            time.sleep(interval)
        else:
            raise ChatAnalysisError(f"Export timed out; last percent={last_percent}")

        result = self.client.post_json(
            "/bff/export/task/result",
            payload,
            self.referer(),
        )
        files: List[ExportFile] = []
        for item in result.get("data") or []:
            name = str(item.get("name") or "")
            url = str(item.get("url") or "")
            if url:
                files.append(ExportFile(name=name, url=url))
        if not files:
            raise ChatAnalysisError("No download URL returned for chat analysis export.")
        return files

    def download_file(self, item: ExportFile, output_dir: Path) -> Path:
        request = Request(
            exporter.encode_download_url(item.url),
            headers={"User-Agent": self.user_agent},
            method="GET",
        )
        try:
            with urlopen(request, timeout=300) as response:
                content = response.read()
                filename = exporter.infer_filename(item.url, response.headers, item.name or "download.zip")
        except HTTPError as exc:
            raise ChatAnalysisError(f"Download HTTP {exc.code} for {item.name or item.url}") from exc
        except URLError as exc:
            raise ChatAnalysisError(f"Download network error for {item.name or item.url}: {exc.reason}") from exc

        output_dir.mkdir(parents=True, exist_ok=True)
        path = output_dir / exporter.sanitize_filename(filename)
        path.write_bytes(content)
        return path


def safe_extract_xlsx(zip_path: Path, destination: Path) -> List[Path]:
    extracted: List[Path] = []
    with zipfile.ZipFile(zip_path) as archive:
        for info in archive.infolist():
            name = Path(info.filename).name
            if not name.lower().endswith(".xlsx"):
                continue
            target = exporter.unique_path(destination / exporter.sanitize_filename(name))
            with archive.open(info) as source, target.open("wb") as output:
                shutil.copyfileobj(source, output)
            extracted.append(target)
    return extracted


def ns(tag: str) -> str:
    return f"{{{NS_MAIN}}}{tag}"


def package_ns(tag: str) -> str:
    return f"{{{NS_PACKAGE_REL}}}{tag}"


def column_index(column: str) -> int:
    result = 0
    for char in column:
        result = result * 26 + ord(char.upper()) - ord("A") + 1
    return result


def column_name(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result or "A"


def split_cell_ref(ref: str) -> Tuple[str, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    if not match:
        return "A", 1
    return match.group(1), int(match.group(2))


def xlsx_first_sheet_path(zip_file: zipfile.ZipFile) -> str:
    try:
        workbook = ET.fromstring(zip_file.read("xl/workbook.xml"))
        sheets = workbook.find(ns("sheets"))
        first_sheet = sheets.find(ns("sheet")) if sheets is not None else None
        rel_id = first_sheet.attrib.get(f"{{{NS_REL}}}id") if first_sheet is not None else None
        rels = ET.fromstring(zip_file.read("xl/_rels/workbook.xml.rels"))
        for rel in rels.findall(package_ns("Relationship")):
            if rel.attrib.get("Id") == rel_id:
                target = rel.attrib.get("Target", "worksheets/sheet1.xml")
                return "xl/" + target.lstrip("/")
    except Exception:
        pass
    return "xl/worksheets/sheet1.xml"


def load_shared_strings(zip_file: zipfile.ZipFile) -> Optional[ET.Element]:
    try:
        return ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    except KeyError:
        return None


def row_number(row: ET.Element) -> int:
    value = row.attrib.get("r")
    if value and value.isdigit():
        return int(value)
    cells = row.findall(ns("c"))
    if not cells:
        return 0
    return split_cell_ref(cells[0].attrib.get("r", "A0"))[1]


def update_row_refs(row: ET.Element, new_row_number: int) -> int:
    row.attrib["r"] = str(new_row_number)
    max_col = 1
    for cell in row.findall(ns("c")):
        col, _old_row = split_cell_ref(cell.attrib.get("r", "A1"))
        cell.attrib["r"] = f"{col}{new_row_number}"
        max_col = max(max_col, column_index(col))
    return max_col


def remap_shared_strings(
    row: ET.Element,
    source_strings: Optional[ET.Element],
    target_strings: Optional[ET.Element],
) -> Tuple[Optional[ET.Element], int]:
    appended = 0
    for cell in row.findall(ns("c")):
        if cell.attrib.get("t") != "s":
            continue
        value = cell.find(ns("v"))
        if value is None or value.text is None or source_strings is None:
            continue
        try:
            source_index = int(value.text)
        except ValueError:
            continue
        source_items = source_strings.findall(ns("si"))
        if source_index < 0 or source_index >= len(source_items):
            continue
        if target_strings is None:
            target_strings = ET.Element(ns("sst"), {"xmlns": NS_MAIN})
        new_item = copy.deepcopy(source_items[source_index])
        target_strings.append(new_item)
        value.text = str(len(target_strings.findall(ns("si"))) - 1)
        appended += 1
    return target_strings, appended


def merge_xlsx_files(xlsx_paths: List[Path], output_path: Path) -> Path:
    if not xlsx_paths:
        raise ChatAnalysisError("No xlsx files found inside exported zip.")
    if len(xlsx_paths) == 1:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(xlsx_paths[0], output_path)
        return output_path

    first_path = xlsx_paths[0]
    with zipfile.ZipFile(first_path) as target_zip:
        target_sheet_path = xlsx_first_sheet_path(target_zip)
        target_sheet = ET.fromstring(target_zip.read(target_sheet_path))
        target_shared = load_shared_strings(target_zip)
        existing_names = target_zip.namelist()
        sheet_data = target_sheet.find(ns("sheetData"))
        if sheet_data is None:
            raise ChatAnalysisError(f"No sheetData found in {first_path.name}")
        existing_rows = sheet_data.findall(ns("row"))
        next_row_number = max((row_number(row) for row in existing_rows), default=0) + 1
        max_col = 1
        for row in existing_rows:
            for cell in row.findall(ns("c")):
                col, _ = split_cell_ref(cell.attrib.get("r", "A1"))
                max_col = max(max_col, column_index(col))

        shared_append_count = 0
        for source_path in xlsx_paths[1:]:
            with zipfile.ZipFile(source_path) as source_zip:
                source_sheet_path = xlsx_first_sheet_path(source_zip)
                source_sheet = ET.fromstring(source_zip.read(source_sheet_path))
                source_sheet_data = source_sheet.find(ns("sheetData"))
                if source_sheet_data is None:
                    continue
                source_shared = load_shared_strings(source_zip)
                rows = source_sheet_data.findall(ns("row"))
                for source_row in rows[1:]:
                    new_row = copy.deepcopy(source_row)
                    target_shared, appended = remap_shared_strings(new_row, source_shared, target_shared)
                    shared_append_count += appended
                    max_col = max(max_col, update_row_refs(new_row, next_row_number))
                    sheet_data.append(new_row)
                    next_row_number += 1

        dimension = target_sheet.find(ns("dimension"))
        if dimension is not None:
            dimension.attrib["ref"] = f"A1:{column_name(max_col)}{max(1, next_row_number - 1)}"

        output_path.parent.mkdir(parents=True, exist_ok=True)
        temp_output = output_path.with_suffix(".tmp")
        with zipfile.ZipFile(temp_output, "w", compression=zipfile.ZIP_DEFLATED) as output_zip:
            for item in target_zip.infolist():
                if item.filename == target_sheet_path:
                    output_zip.writestr(item, ET.tostring(target_sheet, encoding="utf-8", xml_declaration=True))
                elif item.filename == "xl/sharedStrings.xml" and target_shared is not None:
                    target_shared.attrib["count"] = str(
                        int(target_shared.attrib.get("count", len(target_shared.findall(ns("si")))))
                        + shared_append_count
                    )
                    target_shared.attrib["uniqueCount"] = str(len(target_shared.findall(ns("si"))))
                    output_zip.writestr(item, ET.tostring(target_shared, encoding="utf-8", xml_declaration=True))
                else:
                    output_zip.writestr(item, target_zip.read(item.filename))

            if target_shared is not None and "xl/sharedStrings.xml" not in existing_names:
                target_shared.attrib["count"] = str(len(target_shared.findall(ns("si"))))
                target_shared.attrib["uniqueCount"] = str(len(target_shared.findall(ns("si"))))
                output_zip.writestr(
                    "xl/sharedStrings.xml",
                    ET.tostring(target_shared, encoding="utf-8", xml_declaration=True),
                )

        temp_output.replace(output_path)
    return output_path


def collect_xlsx_from_exports(zip_paths: Iterable[Path], temp_dir: Path) -> List[Path]:
    chunk_dir = temp_dir / "chunks"
    chunk_dir.mkdir(parents=True, exist_ok=True)
    xlsx_paths: List[Path] = []
    for zip_path in zip_paths:
        xlsx_paths.extend(safe_extract_xlsx(zip_path, chunk_dir))
    xlsx_paths.sort(key=lambda item: item.name)
    return xlsx_paths


def export_for_date(target_date: date, output_dir: Path, poll_interval: float, poll_timeout: float) -> Path:
    env_file = exporter.load_dotenv(exporter.default_env_path())
    client = ChatGroupAnalysisClient(env_file)
    total = client.fetch_detail_total(target_date)
    print(f"Date: {target_date:%Y-%m-%d}")
    print(f"Detail rows total: {total}")

    task_id = client.create_export(target_date)
    print(f"Export task id: {task_id}")
    files = client.wait_export_files(task_id, poll_interval, poll_timeout)
    print(f"Export files returned: {len(files)}")

    with tempfile.TemporaryDirectory(prefix="chat_group_analysis_") as temp_name:
        temp_dir = Path(temp_name)
        zip_dir = temp_dir / "zips"
        zip_dir.mkdir(parents=True, exist_ok=True)
        zip_paths = [client.download_file(item, zip_dir) for item in files]
        xlsx_paths = collect_xlsx_from_exports(zip_paths, temp_dir)
        if not xlsx_paths:
            raise ChatAnalysisError("Export zip did not contain any xlsx files.")

        final_name = exporter.sanitize_filename(xlsx_paths[0].name or f"{DEFAULT_OUTPUT_PREFIX}_001_000.xlsx")
        final_path = output_dir / final_name
        merge_xlsx_files(xlsx_paths, final_path)
        return final_path


def main(argv: List[str]) -> int:
    args = parse_args(argv)
    env_file = exporter.load_dotenv(exporter.default_env_path())
    target_date = resolve_target_date(args.date)
    output_dir = Path(args.output_dir or output_dir_for(target_date))
    poll_interval = args.poll_interval
    if poll_interval is None:
        poll_interval = float(exporter.env_value(env_file, "POLL_INTERVAL_SECONDS", "3"))
    poll_timeout = args.poll_timeout
    if poll_timeout is None:
        poll_timeout = float(
            exporter.env_value(env_file, "CHAT_ANALYSIS_POLL_TIMEOUT_SECONDS", "1800")
        )

    if args.dry_run:
        client = ChatGroupAnalysisClient(env_file)
        total = client.fetch_detail_total(target_date)
        print(f"Date: {target_date:%Y-%m-%d}")
        print(f"Detail rows total: {total}")
        return 0

    saved_path = export_for_date(target_date, output_dir, poll_interval, poll_timeout)
    print(f"Saved: {saved_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        raise SystemExit(130)
    except (exporter.ScrmError, ChatAnalysisError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
