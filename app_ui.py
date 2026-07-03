from __future__ import annotations

import json
import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any, Dict, List, Optional

import daily_export_scheduler
import runtime_paths


TASKS = [
    ("super_group_undelivered", "超级群发未送达"),
    ("chat_group_analysis_by_chat", "客户群分析-按群聊"),
    ("reach_customer_summary", "统计触达客户"),
    ("group_send_customer_group_export", "群发客户群导出"),
]

STATUS_COLORS = {
    "未开始": "#6b7280",
    "运行中": "#1d4ed8",
    "成功": "#047857",
    "失败": "#b91c1c",
    "等待扫码": "#b45309",
}


def latest_log(log_dir: Path) -> Optional[Path]:
    logs = sorted(log_dir.glob("*.log"), key=lambda item: item.stat().st_mtime, reverse=True)
    return logs[0] if logs else None


def tail_text(path: Optional[Path], limit: int = 120) -> str:
    if not path or not path.exists():
        return "暂无日志。"
    lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    return "\n".join(lines[-limit:])


def parse_status_file(path: Path) -> Dict[str, str]:
    result: Dict[str, str] = {}
    if not path.exists():
        return result
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        if "=" in line:
            key, value = line.split("=", 1)
            result[key] = value
    return result


def parse_log_events(path: Optional[Path]) -> List[Dict[str, Any]]:
    if not path or not path.exists():
        return []
    events: List[Dict[str, Any]] = []
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        if not line.startswith("@@SCRM_STATUS "):
            continue
        try:
            value = json.loads(line[len("@@SCRM_STATUS ") :])
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            events.append(value)
    return events


def load_export_state(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def exe_command() -> List[str]:
    if getattr(sys, "frozen", False):
        exe = Path(sys.executable).with_name("scrm-exporter.exe")
        return [str(exe)]
    return [sys.executable, str(Path(__file__).resolve().with_name("app_cli.py"))]


def open_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    os.startfile(str(path))


class ExporterUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("企微社群任务自动导出")
        self.geometry("980x680")
        self.minsize(860, 580)
        self.config_dir = runtime_paths.default_config_dir().resolve()
        self.data_dir = runtime_paths.default_data_dir().resolve()
        runtime_paths.ensure_runtime_dirs(self.config_dir, self.data_dir)
        self.process: Optional[subprocess.Popen[str]] = None
        self.task_vars: Dict[str, Dict[str, tk.StringVar]] = {}
        self._build()
        self.after(300, self.refresh)

    def _build(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        header = ttk.Frame(self, padding=(14, 12, 14, 8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)

        ttk.Label(header, text="实时运行状态", font=("Microsoft YaHei UI", 14, "bold")).grid(row=0, column=0, sticky="w")
        self.status_var = tk.StringVar(value="读取中")
        self.message_var = tk.StringVar(value="")
        ttk.Label(header, textvariable=self.status_var, font=("Microsoft YaHei UI", 12)).grid(row=0, column=1, sticky="e")
        ttk.Label(header, textvariable=self.message_var, foreground="#4b5563").grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        paths = ttk.Frame(self, padding=(14, 0, 14, 8))
        paths.grid(row=1, column=0, sticky="ew")
        paths.columnconfigure(1, weight=1)
        ttk.Label(paths, text=f"导出目录：{self.data_dir}").grid(row=0, column=0, sticky="w")
        ttk.Label(paths, text=f"运行目录：{self.config_dir}").grid(row=1, column=0, sticky="w")

        buttons = ttk.Frame(self, padding=(14, 0, 14, 10))
        buttons.grid(row=2, column=0, sticky="ew")
        self.login_button = ttk.Button(buttons, text="扫码登录/刷新登录态", command=self.login)
        self.run_button = ttk.Button(buttons, text="立即运行一次", command=self.run_now)
        self.output_button = ttk.Button(buttons, text="打开导出文件夹", command=lambda: open_dir(self.data_dir))
        self.logs_button = ttk.Button(buttons, text="打开日志文件夹", command=lambda: open_dir(runtime_paths.logs_dir(self.config_dir)))
        self.install_button = ttk.Button(buttons, text="重新安装计划任务", command=self.install_task)
        self.uninstall_button = ttk.Button(buttons, text="卸载计划任务", command=self.uninstall_task)
        for index, button in enumerate(
            [
                self.login_button,
                self.run_button,
                self.output_button,
                self.logs_button,
                self.install_button,
                self.uninstall_button,
            ]
        ):
            button.grid(row=0, column=index, padx=(0, 8), sticky="w")

        body = ttk.PanedWindow(self, orient=tk.VERTICAL)
        body.grid(row=3, column=0, sticky="nsew", padx=14, pady=(0, 14))

        checklist = ttk.LabelFrame(body, text="4 个导出任务 checklist", padding=(10, 8))
        checklist.columnconfigure(1, weight=1)
        for row, (task_id, label) in enumerate(TASKS):
            status_var = tk.StringVar(value="未开始")
            detail_var = tk.StringVar(value="")
            self.task_vars[task_id] = {"status": status_var, "detail": detail_var}
            ttk.Label(checklist, text=label, width=24).grid(row=row, column=0, sticky="w", pady=4)
            status_label = ttk.Label(checklist, textvariable=status_var, width=10)
            status_label.grid(row=row, column=1, sticky="w", pady=4)
            ttk.Label(checklist, textvariable=detail_var, foreground="#4b5563").grid(row=row, column=2, sticky="ew", pady=4)

        log_frame = ttk.LabelFrame(body, text="最近日志", padding=(8, 8))
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=18, wrap="word", font=("Consolas", 9))
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        body.add(checklist, weight=1)
        body.add(log_frame, weight=4)

    def command_args(self, command: str) -> List[str]:
        return [
            *exe_command(),
            command,
            "--config-dir",
            str(self.config_dir),
            "--data-dir",
            str(self.data_dir),
        ]

    def start_process(self, command: str) -> None:
        if self.process and self.process.poll() is None:
            messagebox.showinfo("正在运行", "已有任务正在运行，请稍后。")
            return
        self.process = subprocess.Popen(
            self.command_args(command),
            cwd=str(runtime_paths.app_dir()),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            text=True,
        )
        self.refresh()

    def run_now(self) -> None:
        self.start_process("run")

    def login(self) -> None:
        self.start_process("login")

    def run_admin_command(self, command: str, success: str) -> None:
        completed = subprocess.run(
            self.command_args(command),
            cwd=str(runtime_paths.app_dir()),
            text=True,
            capture_output=True,
        )
        if completed.returncode == 0:
            messagebox.showinfo("完成", success)
        else:
            messagebox.showerror("失败", (completed.stderr or completed.stdout or "命令执行失败").strip())

    def install_task(self) -> None:
        self.run_admin_command("install-task", "计划任务已安装。")

    def uninstall_task(self) -> None:
        self.run_admin_command("uninstall-task", "计划任务已卸载。")

    def refresh(self) -> None:
        log_path = latest_log(runtime_paths.logs_dir(self.config_dir))
        status = parse_status_file(runtime_paths.state_dir(self.config_dir) / daily_export_scheduler.STATUS_FILE_NAME)
        events = parse_log_events(log_path)
        state = load_export_state(runtime_paths.state_dir(self.config_dir) / daily_export_scheduler.STATE_FILE_NAME)

        state_text = status.get("status") or "暂无状态"
        if self.process and self.process.poll() is None:
            state_text = "RUNNING"
        self.status_var.set(state_text)
        self.message_var.set(status.get("message") or f"最新日志：{log_path.name if log_path else '暂无'}")

        checklist = self.build_checklist(events, state)
        for task_id, values in checklist.items():
            self.task_vars[task_id]["status"].set(values["status"])
            self.task_vars[task_id]["detail"].set(values["detail"])

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.insert("1.0", tail_text(log_path))
        self.log_text.configure(state="disabled")

        running = bool(self.process and self.process.poll() is None)
        button_state = "disabled" if running else "normal"
        self.run_button.configure(state=button_state)
        self.login_button.configure(state=button_state)
        self.after(1500, self.refresh)

    def build_checklist(self, events: List[Dict[str, Any]], state: Dict[str, Any]) -> Dict[str, Dict[str, str]]:
        result = {task_id: {"status": "未开始", "detail": ""} for task_id, _label in TASKS}
        last_success = state.get("last_successful_dates") if isinstance(state, dict) else {}
        if isinstance(last_success, dict):
            for task_id, date_value in last_success.items():
                if task_id in result:
                    result[task_id] = {"status": "成功", "detail": f"最近成功：{date_value}"}

        for event in events:
            task_id = str(event.get("task_id") or "")
            if task_id not in result:
                continue
            date_value = str(event.get("date") or "")
            message = str(event.get("message") or "")
            output_dir = str(event.get("output_dir") or "")
            if event.get("event") == "TASK_START":
                result[task_id] = {"status": "运行中", "detail": f"{date_value} {output_dir}".strip()}
            elif event.get("event") == "TASK_OK":
                result[task_id] = {"status": "成功", "detail": f"{date_value} {output_dir}".strip()}
            elif event.get("event") == "TASK_FAILED":
                result[task_id] = {"status": "失败", "detail": f"{date_value} {message}".strip()}
            elif event.get("event") == "WAITING_LOGIN":
                result[task_id] = {"status": "等待扫码", "detail": date_value}
        return result


def main() -> int:
    app = ExporterUI()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
