from __future__ import annotations

import argparse
import contextlib
import ctypes
import io
import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Optional

import daily_export_scheduler
import runtime_paths


TASK_TIME = "09:40"
LOCK_STALE_SECONDS = 12 * 60 * 60


class Tee(io.TextIOBase):
    def __init__(self, *streams: io.TextIOBase) -> None:
        self.streams = streams

    def write(self, text: str) -> int:
        for stream in self.streams:
            stream.write(text)
            stream.flush()
        return len(text)

    def flush(self) -> None:
        for stream in self.streams:
            stream.flush()


class RunLock:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.fd: Optional[int] = None

    def __enter__(self) -> "RunLock":
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if self.path.exists():
            age = time.time() - self.path.stat().st_mtime
            if age > LOCK_STALE_SECONDS:
                self.path.unlink(missing_ok=True)
        try:
            self.fd = os.open(str(self.path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(self.fd, str(os.getpid()).encode("ascii", "ignore"))
        except FileExistsError as exc:
            raise RuntimeError("已有导出任务正在运行，请稍后再试。") from exc
        return self

    def __exit__(self, *_: object) -> None:
        if self.fd is not None:
            os.close(self.fd)
            self.fd = None
        self.path.unlink(missing_ok=True)


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="SCRM daily exporter.")
    parser.add_argument("command", nargs="?", choices=["run", "run-now", "login", "status", "install-task", "uninstall-task"])
    parser.add_argument("--config-dir", help="Runtime config directory.")
    parser.add_argument("--data-dir", help="Export output directory.")
    parser.add_argument("--today", help="Override today's date for tests, YYYY-MM-DD.")
    parser.add_argument("--plan-only", action="store_true", help="Print pending dates without exporting.")
    parser.add_argument("--run-now", action="store_true", help="Run pending exports now.")
    parser.add_argument("--login-only", action="store_true", help="Refresh login only.")
    parser.add_argument("--status", action="store_true", help="Print status.")
    parser.add_argument("--install-task", action="store_true", help="Install the scheduled task.")
    parser.add_argument("--uninstall-task", action="store_true", help="Uninstall the scheduled task.")
    return parser.parse_args(argv)


def configure_console() -> None:
    if os.name == "nt":
        try:
            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
            ctypes.windll.kernel32.SetConsoleCP(65001)
        except Exception:
            pass
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def resolve_runtime(args: argparse.Namespace) -> tuple[Path, Path]:
    config_dir = runtime_paths.resolve_dir(args.config_dir, runtime_paths.default_config_dir())
    data_dir = runtime_paths.resolve_dir(args.data_dir, runtime_paths.default_data_dir())
    runtime_paths.ensure_runtime_dirs(config_dir, data_dir)
    os.environ["SCRM_CONFIG_DIR"] = str(config_dir)
    os.environ["SCRM_DATA_DIR"] = str(data_dir)
    os.environ["SCRM_ENV_PATH"] = str(runtime_paths.env_path(config_dir))
    return config_dir, data_dir


def latest_log(logs_dir: Path) -> Optional[Path]:
    logs = sorted(logs_dir.glob("daily_export_*.log"), key=lambda item: item.stat().st_mtime, reverse=True)
    return logs[0] if logs else None


def read_status(config_dir: Path) -> dict[str, str]:
    status_path = runtime_paths.state_dir(config_dir) / daily_export_scheduler.STATUS_FILE_NAME
    result: dict[str, str] = {}
    if not status_path.exists():
        return result
    for line in status_path.read_text(encoding="utf-8", errors="replace").splitlines():
        if "=" in line:
            key, value = line.split("=", 1)
            result[key] = value
    return result


def print_status(config_dir: Path, data_dir: Path) -> int:
    status = read_status(config_dir)
    print(f"Config: {config_dir}")
    print(f"Data: {data_dir}")
    print(f"Status: {status.get('status', '(none)')}")
    print(f"Message: {status.get('message', '(none)')}")
    log_path = latest_log(runtime_paths.logs_dir(config_dir))
    print(f"Latest log: {log_path or '(none)'}")
    state_path = runtime_paths.state_dir(config_dir) / daily_export_scheduler.STATE_FILE_NAME
    if state_path.exists():
        state = json.loads(state_path.read_text(encoding="utf-8"))
        print(f"Last successful dates: {state.get('last_successful_dates') or {}}")
    return 0


def scheduler_args(args: argparse.Namespace, config_dir: Path, data_dir: Path, login_only: bool) -> List[str]:
    result = ["--config-dir", str(config_dir), "--data-dir", str(data_dir)]
    if args.today:
        result.extend(["--today", args.today])
    if args.plan_only:
        result.append("--plan-only")
    if login_only:
        result.append("--login-only")
    return result


def run_with_logging(args: argparse.Namespace, config_dir: Path, data_dir: Path, login_only: bool = False) -> int:
    log_dir = runtime_paths.logs_dir(config_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    prefix = "login" if login_only else "daily_export"
    log_path = log_dir / f"{prefix}_{stamp}.log"
    lock_path = runtime_paths.state_dir(config_dir) / "export.lock"

    with log_path.open("w", encoding="utf-8", buffering=1) as log_file:
        tee_stdout = Tee(sys.__stdout__, log_file)
        tee_stderr = Tee(sys.__stderr__, log_file)
        with contextlib.redirect_stdout(tee_stdout), contextlib.redirect_stderr(tee_stderr):
            print(f"Started: {datetime.now().astimezone().isoformat()}")
            print(f"Config directory: {config_dir}")
            print(f"Data directory: {data_dir}")
            print(f"Log file: {log_path}")
            try:
                if login_only:
                    exit_code = daily_export_scheduler.main(scheduler_args(args, config_dir, data_dir, True))
                else:
                    with RunLock(lock_path):
                        exit_code = daily_export_scheduler.main(scheduler_args(args, config_dir, data_dir, False))
            except Exception as exc:
                print(f"Error: {type(exc).__name__}: {exc}", file=sys.stderr)
                exit_code = 1
            print(f"Finished: {datetime.now().astimezone().isoformat()}, exit code: {exit_code}")
            return int(exit_code)


def executable_for_task() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()
    return Path(sys.executable).resolve()


def task_command(config_dir: Path, data_dir: Path) -> str:
    exe = executable_for_task()
    if getattr(sys, "frozen", False):
        return f'"{exe}" run --config-dir "{config_dir}" --data-dir "{data_dir}"'
    script = Path(__file__).resolve()
    return f'"{exe}" "{script}" run --config-dir "{config_dir}" --data-dir "{data_dir}"'


def run_subprocess(command: Iterable[str]) -> int:
    completed = subprocess.run(list(command), text=True, capture_output=True)
    if completed.stdout:
        print(completed.stdout.rstrip())
    if completed.stderr:
        print(completed.stderr.rstrip(), file=sys.stderr)
    return completed.returncode


def install_task(config_dir: Path, data_dir: Path) -> int:
    command = task_command(config_dir, data_dir)
    args = [
        "schtasks",
        "/Create",
        "/F",
        "/SC",
        "DAILY",
        "/ST",
        TASK_TIME,
        "/TN",
        runtime_paths.TASK_NAME,
        "/TR",
        command,
    ]
    code = run_subprocess(args)
    if code == 0:
        print(f"计划任务已安装：{runtime_paths.TASK_NAME}，每天 {TASK_TIME}。")
    return code


def uninstall_task() -> int:
    code = run_subprocess(["schtasks", "/Delete", "/F", "/TN", runtime_paths.TASK_NAME])
    if code == 0:
        print(f"计划任务已卸载：{runtime_paths.TASK_NAME}。")
    return code


def choose_command(args: argparse.Namespace) -> str:
    if args.command:
        return "run" if args.command == "run-now" else args.command
    if args.login_only:
        return "login"
    if args.status:
        return "status"
    if args.install_task:
        return "install-task"
    if args.uninstall_task:
        return "uninstall-task"
    return "run" if args.run_now else "status"


def main(argv: Optional[List[str]] = None) -> int:
    configure_console()
    args = parse_args(argv)
    config_dir, data_dir = resolve_runtime(args)
    command = choose_command(args)
    if command == "run":
        return run_with_logging(args, config_dir, data_dir)
    if command == "login":
        return run_with_logging(args, config_dir, data_dir, login_only=True)
    if command == "status":
        return print_status(config_dir, data_dir)
    if command == "install-task":
        return install_task(config_dir, data_dir)
    if command == "uninstall-task":
        return uninstall_task()
    print(f"Unknown command: {command}", file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
