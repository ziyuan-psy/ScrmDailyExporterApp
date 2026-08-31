# SCRM Daily Exporter

This is an installable Windows GUI application that automatically exports daily WeCom community-related tasks, generates local statistics files, and can optionally append daily reach metrics to Feishu spreadsheets.

End users do not need to install Python. A Python development environment is required only to maintain the source code, rebuild the application, or modify its logic.

## 1. Using the Current GUI

### 1.1 What the application does

After installation, the application backfills unfinished dates and runs the following seven tasks in a fixed order:

1. Undelivered mass messages
2. Customer group analysis by group chat
3. Reach counts for mass customer messages and Moments posts
4. Customer group mass-message export
5. Reach summary
6. Store-group reach counts
7. Online spreadsheet sync

Tasks 1–4 export raw files and Word statistics documents from the WeCom admin console. Tasks 5–6 use the local exports to generate `社群任务触达人数日报.xlsx`. When Feishu sync is enabled, Task 7 appends dates from the local daily report that are newer than those in the online spreadsheets.

### 1.2 Installation and first-time setup

End users only need to run the installer:

```text
ScrmDailyExporterSetup.exe
```

The installer automatically:

- Installs the application to `%LOCALAPPDATA%\ScrmDailyExporter\app`
- Creates the Start menu shortcut `SCRM Daily Exporter`
- Registers the scheduled task `每日企微私域任务导出` for the current Windows user
- Opens the GUI and runs the application automatically at 09:40 every day by default

For first-time use, open `企微社群任务自动导出`, click `扫码登录/刷新登录态` (Scan to Sign In / Refresh Session), and scan the QR code in the Chrome window to sign in to the WeCom admin console.

The login session is stored only on the current computer. It is not distributed with the installer or source code. You must scan the QR code again when switching computers or accounts, when the session expires, or when the WeCom admin console requires a new login.

### 1.3 Default paths

| Item | Path |
| --- | --- |
| Application directory | `%LOCALAPPDATA%\ScrmDailyExporter\app` |
| Runtime configuration directory | `%LOCALAPPDATA%\ScrmDailyExporter` |
| Log directory | `%LOCALAPPDATA%\ScrmDailyExporter\logs` |
| State directory | `%LOCALAPPDATA%\ScrmDailyExporter\state` |
| GUI settings file | `%LOCALAPPDATA%\ScrmDailyExporter\app_settings.json` |
| Login browser profile | `%LOCALAPPDATA%\ScrmDailyExporter\chrome-profile` |
| Default export directory | `%USERPROFILE%\Documents\每日企微私域任务导出` |
| Daily reach report | `%USERPROFILE%\Documents\每日企微私域任务导出\社群任务触达人数日报.xlsx` |

### 1.4 GUI controls and inputs

| Control | Purpose |
| --- | --- |
| 扫码登录/刷新登录态 | Opens the dedicated Chrome login window and refreshes the local WeCom credentials in `.env` after QR-code login |
| 立即运行一次 | Immediately backfills all currently unfinished tasks |
| 打开导出文件夹 | Opens the local export directory |
| 打开日志文件夹 | Opens the runtime log directory |
| 飞书同步配置 | Configures the App ID, App Secret, and the links to two Feishu spreadsheet sheets |
| 重新安装计划任务 | Re-registers the task that runs automatically every day at 09:40 |
| 卸载计划任务 | Deletes the daily scheduled task without removing the application or historical exports |
| 从指定日期开始导出 | Runs from the specified date through yesterday and skips tasks that have already succeeded |
| 全局自动补导起始日期 | Controls the first date for automatic backfills; clearing it restores the default seven-day lookback |

### 1.5 Output files

The export directory contains a dated folder for each day as well as summary files maintained across dates. Raw Excel files are stored in dated folders similar to:

```text
%USERPROFILE%\Documents\每日企微私域任务导出\社群任务0825
```

Common outputs include:

| File | Location | Source |
| --- | --- | --- |
| Mass-message-related Excel files | Daily dated folder | Task 1, exported from the WeCom admin console |
| Customer group analysis by group chat Excel files | Daily dated folder | Task 2, exported from the WeCom admin console |
| `企微社群任务触达客户统计.docx` | Export directory root | Task 3, containing reach counts for mass customer messages and Moments posts |
| Customer group mass-message Excel files | Daily dated folder | Task 4, exported from the WeCom admin console |
| `社群任务触达人数日报.xlsx` | Export directory root | Tasks 5–6, containing the reach summary and store-group reach counts |

Rules for writing `企微社群任务触达客户统计.docx`:

- By default, the file is stored in the export directory root instead of being split into separate daily Word files.
- Each date has its own block with a heading such as `2026.8.25`.
- Two fixed rows appear below each date: the reach count for that day's `群发客户` task and the reach count for that day's `群发朋友圈` task.
- Each row records the number of matched tasks and the reach count for each task. If more than one task of the same type is found, the day's total reach is also included.
- Rerunning the same date replaces the existing block for that date to prevent duplicates. The file is not updated if the content has not changed.

`社群任务触达人数日报.xlsx` currently contains the following worksheets:

- `触达人数汇总`
- `门店分组触达人数`
- `说明`

Key columns in `触达人数汇总`:

| Column | Content |
| --- | --- |
| A | Date in `yyyy/mm/dd` format |
| B | Day of the week, from `星期一` through `星期日` |
| C | Welfare-account friends, based on the total reach for mass customer messages in the Word document |
| F | Community reach, based on the community-task reach count |
| I | Moments reach, based on the total reach for mass Moments posts in the Word document |

Key columns in `门店分组触达人数`:

| Column | Content |
| --- | --- |
| A | Date |
| B | Type |
| C | Store group |
| D | Coupon type |
| E | Reach count |

### 1.6 Business rules and metric definitions

#### Task-title filters

The following export tasks exclude titles containing any of these keywords by default:

- `测试`
- `海外`
- `境外`

This applies to:

- Undelivered mass messages
- Customer group mass-message exports
- Reach counts for mass customer messages and Moments posts

Related configuration variables:

| Task | Configuration variable |
| --- | --- |
| Undelivered mass messages | `EXCLUDE_TASK_KEYWORDS` |
| Customer group mass-message exports | `CUSTOMER_GROUP_EXCLUDE_KEYWORDS` |
| Reach counts for mass customer messages and Moments posts | `REACH_EXCLUDE_KEYWORDS` |

If the exclusion-keyword configuration is empty, no tasks are excluded by keyword. If inclusion keywords are configured, only tasks whose titles match an inclusion keyword are exported.

#### Community reach

Community reach is calculated from the task-export Excel files in each dated folder. The rules are:

- Count only rows where `送达状态=已送达`
- Match `客户群chatid（本应用）` in the task file to `客户群ID` in the customer-group statistics file
- After a successful match, add the corresponding `群客户总数`
- If the same group appears in multiple tasks, count every contact occurrence rather than deduplicating the group
- Include both `超级群发` and `群发客户群` tasks in community reach

#### Reach for mass customer messages and Moments posts

Task 3 generates `企微社群任务触达客户统计.docx`. The two totals in the Word document are written to the daily report as follows:

| Word metric | Daily report field |
| --- | --- |
| Total reach for mass customer messages | Welfare-account friends |
| Total reach for mass Moments posts | Moments |

#### Store-group reach

Store-group statistics include only tasks whose filenames contain a store-group marker:

| Filename marker | Store group |
| --- | --- |
| `AFD` | `社群-A档` |
| `BFD` | `社群-B档` |
| `CFD` | `社群-C档` |
| `SFD` | `社群-S档` |

The calculation includes:

- `超级群发` files whose names contain `AFD`, `BFD`, `CFD`, or `SFD`
- `群发客户群` files whose names contain `AFD`, `BFD`, `CFD`, or `SFD`

Regular `D` files and files without an `AFD/BFD/CFD/SFD` marker are excluded. The food-coupon row equals the sum of the A/B/C/S reach counts for that day.

#### Reruns and historical data

- The local daily report is updated by date. Recalculating a date overwrites the old rows for that date and prevents duplicate appends.
- By default, only the most recent days are recalculated on a rolling basis; older historical dates remain unchanged.
- Feishu sync only appends dates from the local report that are newer than the latest online date. It does not overwrite, clear, or modify historical online rows.

### 1.7 Feishu sync configuration

Feishu sync is disabled by default. To enable it, click `飞书同步配置` (Feishu Sync Settings) in the GUI and enter:

```text
App ID
App Secret
Reach summary sheet link
Store-group reach sheet link
```

After `测试连接` (Test Connection) succeeds, select `启用飞书同步` (Enable Feishu Sync) and save the settings.

`启用飞书同步` is the master switch for Task 7:

- Disabled: automatic runs complete only the first six tasks and neither connect to Feishu nor write to online spreadsheets.
- Enabled: Task 7 syncs the online spreadsheets after the first six tasks finish.

Sync is append-only:

- Read the dates already present in the Feishu spreadsheets
- Append only local rows whose dates are newer than the latest online date
- Do not overwrite, clear, or modify historical online rows
- Automatically apply centered alignment, date formatting, and thousands separators after appending

The Feishu app must have permission to read and write spreadsheets. The `App Secret` is stored in `%LOCALAPPDATA%\ScrmDailyExporter\.env` on the current computer. Do not commit it to Git or distribute it with the source package.

### 1.8 Troubleshooting

**The scheduled task did not run**

Confirm that the computer was powered on and the current Windows user was signed in at 09:40. You can click `重新安装计划任务` in the GUI to refresh the scheduled task.

**The login window does not open or keeps waiting for the QR-code scan**

The application uses Chrome debug port `9333` by default. If another application is using it, change the following value in `%LOCALAPPDATA%\ScrmDailyExporter\.env`:

```text
CHROME_DEBUG_PORT=9334
```

Then click `扫码登录/刷新登录态` again.

**Writing to the Excel file fails**

If the log reports a `PermissionError` or says that the file is in use, `社群任务触达人数日报.xlsx` is usually open in Excel. Close the file and rerun the application.

**Dates appear as numbers in the Feishu spreadsheet**

The current version automatically applies formatting after appending rows. If older data already appears as numbers, use the maintenance script's format-repair command for the affected date, as described in Section 2.8.

**Online spreadsheet sync fails**

First click `测试连接` in the GUI's `飞书同步配置` dialog. If the test fails, check the App ID, App Secret, spreadsheet links, app permissions, and whether the spreadsheets have been shared with the app.

## 2. Source Maintenance and Extension

### 2.1 Source structure

| File | Purpose |
| --- | --- |
| `app_ui.py` | Main GUI, task checklist, buttons, and Feishu sync settings dialog |
| `app_cli.py` | Command-line entry point for run, login, status, install-task, and uninstall-task |
| `daily_export_scheduler.py` | Daily scheduling and task orchestration, including task state, backfill dates, and login refresh |
| `export_super_group_undelivered.py` | Undelivered mass-message exports and core SCRM API client logic |
| `export_chat_group_analysis_by_chat.py` | Customer group analysis exports by group chat |
| `export_reach_customer_summary.py` | Calculates reach for mass customer messages and Moments posts and generates the Word document |
| `export_group_send_customer_group.py` | Customer group mass-message exports |
| `export_reach_daily_excel.py` | Generates the local `社群任务触达人数日报.xlsx` report |
| `sync_feishu_reach_workbook.py` | Appends local daily-report data to Feishu spreadsheets |
| `scrm_browser_fetch.py` | Fallback that sends same-origin requests through Chrome DevTools in an authenticated page |
| `runtime_paths.py` | Rules for application, configuration, log, and other runtime paths |
| `app_settings.py` | GUI settings such as the export directory and global backfill start date |
| `state_file_io.py` | State-file reading and writing |
| `requirements.txt` | Python dependencies |
| `ScrmDailyExporter.spec` | Production PyInstaller configuration |
| `build_release.ps1` | Generates `dist\ScrmDailyExporter` |
| `build_installer.ps1` | Builds the installer from `dist` |
| `install_local.ps1` | Quickly installs the application locally without generating an installer |
| `installer.iss` | Inno Setup installer configuration |

### 2.2 Local development environment

After obtaining a standard source package, maintainers should run the following commands in the source directory:

```powershell
python -m venv .venv
.\.venv\Scripts\pip.exe install -r requirements.txt
```

Start the GUI directly from source during development:

```powershell
.\.venv\Scripts\python.exe app_ui.py
```

You can also test the schedule plan without performing real exports:

```powershell
.\.venv\Scripts\python.exe app_cli.py run --test-mode --plan-only
```

Common commands:

```powershell
.\.venv\Scripts\python.exe app_cli.py status
.\.venv\Scripts\python.exe app_cli.py login
.\.venv\Scripts\python.exe app_cli.py run --start-date 2026-08-01
.\.venv\Scripts\python.exe app_cli.py install-task --test-mode
.\.venv\Scripts\python.exe app_cli.py uninstall-task --test-mode
```

`--test-mode` uses separate test configuration, output directories, and scheduled-task names so that the production installation is not affected.

### 2.3 Build and release

Generate the directly runnable executable directory:

```powershell
.\build_release.ps1
```

Output directory:

```text
dist\ScrmDailyExporter
```

Generate the production installer:

```powershell
.\build_installer.ps1
```

Output file:

```text
installer-output\ScrmDailyExporterSetup.exe
```

Send `ScrmDailyExporterSetup.exe` to end users. You may separately provide maintainers with a clean source package.

`dist`, `build`, `installer-output`, `.venv`, and `.env` are not committed to Git. It is normal for a source package not to include `dist`; run `build_release.ps1` to regenerate it when needed.

### 2.4 Configuration and runtime data

All runtime configuration is stored under `%LOCALAPPDATA%\ScrmDailyExporter` for the current Windows user.

| File or directory | Content |
| --- | --- |
| `.env` | Sensitive settings such as WeCom tokens/cookies, Feishu App ID/Secret, and the Chrome port |
| `logs` | Detailed logs for each run |
| `state\export_state.json` | Most recent successful date for each task |
| `state\latest_status.txt` | Current status summary shown in the GUI |
| `app_settings.json` | GUI settings for the export directory and global backfill start date |
| `chrome-profile` | Dedicated Chrome login profile |

Do not commit or distribute `.env`. Each colleague who needs to use the application independently should scan the QR code on their own computer and enter their own Feishu sync settings in the GUI.

### 2.5 If the WeCom admin API changes

First inspect the logs to identify the failed task:

```powershell
Get-Content -LiteralPath "$env:LOCALAPPDATA\ScrmDailyExporter\state\latest_status.txt" -Raw
```

Then open the latest log:

```powershell
Get-ChildItem "$env:LOCALAPPDATA\ScrmDailyExporter\logs" -Filter *.log |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 1
```

Use the failed task to locate the relevant source file:

| Failed task | File to inspect first |
| --- | --- |
| Undelivered mass messages | `export_super_group_undelivered.py` |
| Customer group analysis by group chat | `export_chat_group_analysis_by_chat.py` |
| Reach for mass customer messages and Moments posts | `export_reach_customer_summary.py` |
| Customer group mass-message exports | `export_group_send_customer_group.py` |
| Reach summary or store-group reach | `export_reach_daily_excel.py` |
| Online spreadsheet sync | `sync_feishu_reach_workbook.py` |

General process for investigating an API change:

1. Use `扫码登录/刷新登录态` in the GUI to open the dedicated Chrome window.
2. Manually perform the corresponding admin action in Chrome while recording it in the DevTools Network panel.
3. Compare the new request URL, payload, headers, and response fields.
4. Update the API path, request parameters, or field parsing in the relevant `export_*.py` file.
5. Check the schedule plan with `--test-mode --plan-only`.
6. Validate a small real export for a specified date.
7. After validation, rerun `build_release.ps1` and `build_installer.ps1`.

An expired WeCom session does not necessarily indicate an API change. Scan the QR code again first. If the log contains `WinError 10061`, `127.0.0.1:9333`, or `Chrome debug port`, the Chrome debugging window is usually not running or the port is already in use.

### 2.6 If the Feishu API or spreadsheet format changes

Inspect `sync_feishu_reach_workbook.py` first.

Key logic includes:

- Parsing Feishu spreadsheet links
- Obtaining a `tenant_access_token`
- Converting a wiki token to a spreadsheet token
- Reading dates already present in the online spreadsheets
- Determining whether only newer dates should be appended
- Writing to `触达人数汇总`
- Writing to `门店分组触达人数`
- Applying date formats, thousands separators, and centered alignment after appending

If a spreadsheet sheet is replaced, end users can update its link directly in the GUI's `飞书同步配置` dialog without changing the source code.

If the online spreadsheet's column structure changes, update all of the following:

- `SUMMARY_COLUMNS`
- `STORE_GROUP_COLUMNS`
- `read_local_summary_rows`
- `read_local_store_group_rows`
- `format_summary_rows`
- `format_store_group_rows`

### 2.7 Adding a task

Adding a task usually requires all of the following changes:

1. Add or extend the relevant export or statistics script.
2. Add the task ID, label, run function, and dependency order in `daily_export_scheduler.py`.
3. Add the task to the `TASKS` checklist in `app_ui.py`.
4. Add state migration or initialization logic if the task requires it.
5. Update the README's task list, output-file documentation, and maintenance instructions.
6. Check the schedule plan with `--test-mode --plan-only`.
7. Validate the task on a limited range of specified dates.

Avoid renaming a task ID after release because `state\export_state.json` records successful dates by task ID.

### 2.8 Debugging and maintenance commands

Syntax check:

```powershell
.\.venv\Scripts\python.exe -m compileall -q .
```

View the installed application's status:

```powershell
& "$env:LOCALAPPDATA\ScrmDailyExporter\app\scrm-exporter.exe" status
```

Inspect the installed application's schedule plan:

```powershell
& "$env:LOCALAPPDATA\ScrmDailyExporter\app\scrm-exporter.exe" run --plan-only
```

Backfill from a specified date:

```powershell
& "$env:LOCALAPPDATA\ScrmDailyExporter\app\scrm-exporter.exe" run --start-date 2026-08-01
```

Repair Feishu formatting for one date only:

```powershell
.\.venv\Scripts\python.exe .\sync_feishu_reach_workbook.py `
  --workbook "$env:USERPROFILE\Documents\每日企微私域任务导出\社群任务触达人数日报.xlsx" `
  --env "$env:LOCALAPPDATA\ScrmDailyExporter\.env" `
  --date 2026-08-16 `
  --format-only
```

Dry-run Feishu sync:

```powershell
.\.venv\Scripts\python.exe .\sync_feishu_reach_workbook.py `
  --workbook "$env:USERPROFILE\Documents\每日企微私域任务导出\社群任务触达人数日报.xlsx" `
  --env "$env:LOCALAPPDATA\ScrmDailyExporter\.env" `
  --date 2026-08-16 `
  --dry-run
```

---

<a id="english"></a>

# SCRM Daily Exporter

[English](#english) | [中文](#chinese)

An installable Windows GUI application that automatically exports daily WeCom SCRM community-task data, generates local summary reports, and can optionally append daily reach metrics to Feishu spreadsheets.

End users do not need Python installed. A Python development environment is required only to maintain the source code, rebuild the application, or change its logic.

## 1. Using the Current GUI

### 1.1 What the application does

After installation, the application catches up on incomplete dates and runs seven tasks in a fixed order:

1. Export undelivered Super Mass Messaging records
2. Export Customer Group Analytics by group chat
3. Calculate reach for customer mass messages and Moments posts
4. Export customer-group mass-message records
5. Generate the reach summary
6. Generate reach by store group
7. Sync reports to Feishu spreadsheets

Tasks 1–4 export raw files and Word summaries from the WeCom admin console. Tasks 5–6 use the local exports to generate <code>社群任务触达人数日报.xlsx</code> (Daily Community-Task Reach Report). When Feishu sync is enabled, Task 7 appends dates from the local report that are newer than those already in the online spreadsheets.

### 1.2 Installation and first-time setup

End users only need to run the installer:

~~~text
ScrmDailyExporterSetup.exe
~~~

The installer automatically:

- Installs the application to <code>%LOCALAPPDATA%\ScrmDailyExporter\app</code>
- Creates the Start menu shortcut <code>SCRM Daily Exporter</code>
- Registers the current Windows user's scheduled task <code>每日企微私域任务导出</code>
- Opens the GUI and runs the workflow once per day at 09:40 by default

On first use, open <code>企微社群任务自动导出</code> and click <code>扫码登录/刷新登录态</code> (Scan to Log In / Refresh Session). Scan the QR code in the Chrome window that opens to sign in to the WeCom admin console.

The login session is stored only on the current computer and is never included with the installer or source code. You must scan again after changing computers or accounts, after the session expires, or whenever the WeCom admin console requires a new login.

### 1.3 Default paths

| Item | Path |
| --- | --- |
| Application directory | <code>%LOCALAPPDATA%\ScrmDailyExporter\app</code> |
| Runtime configuration directory | <code>%LOCALAPPDATA%\ScrmDailyExporter</code> |
| Log directory | <code>%LOCALAPPDATA%\ScrmDailyExporter\logs</code> |
| State directory | <code>%LOCALAPPDATA%\ScrmDailyExporter\state</code> |
| GUI settings file | <code>%LOCALAPPDATA%\ScrmDailyExporter\app_settings.json</code> |
| Login browser profile | <code>%LOCALAPPDATA%\ScrmDailyExporter\chrome-profile</code> |
| Default export directory | <code>%USERPROFILE%\Documents\每日企微私域任务导出</code> |
| Daily reach report | <code>%USERPROFILE%\Documents\每日企微私域任务导出\社群任务触达人数日报.xlsx</code> |

### 1.4 GUI controls

| Control | Purpose |
| --- | --- |
| <code>扫码登录/刷新登录态</code> | Opens a dedicated Chrome login window and refreshes the local WeCom login information in <code>.env</code> after QR-code login |
| <code>立即运行一次</code> | Immediately runs all currently incomplete tasks |
| <code>打开导出文件夹</code> | Opens the local export directory |
| <code>打开日志文件夹</code> | Opens the log directory |
| <code>飞书同步配置</code> | Configures the App ID, App Secret, and links to the two Feishu spreadsheet sheets |
| <code>重新安装计划任务</code> | Re-registers the scheduled task that runs daily at 09:40 |
| <code>卸载计划任务</code> | Removes the daily scheduled task without deleting the application or historical exports |
| <code>从指定日期开始导出</code> | Runs from the entered date through yesterday and skips tasks that have already succeeded |
| <code>全局自动补导起始日期</code> | Sets the earliest date for automatic catch-up; clearing it restores the default seven-day catch-up window |

### 1.5 Output files

The export directory contains a dated folder for each day as well as cross-date summary files. Daily raw Excel files are stored in folders such as:

~~~text
%USERPROFILE%\Documents\每日企微私域任务导出\社群任务0825
~~~

Typical outputs include:

| File | Location | Source |
| --- | --- | --- |
| Super Mass Messaging Excel files | Daily dated folder | Task 1, exported from the WeCom admin console |
| Customer Group Analytics by group chat Excel files | Daily dated folder | Task 2, exported from the WeCom admin console |
| <code>企微社群任务触达客户统计.docx</code> | Export directory root | Task 3, summarizes reach for customer mass messages and Moments posts |
| Customer-group mass-message Excel files | Daily dated folder | Task 4, exported from the WeCom admin console |
| <code>社群任务触达人数日报.xlsx</code> | Export directory root | Tasks 5–6, contains the overall reach summary and reach by store group |

Rules for writing <code>企微社群任务触达客户统计.docx</code>:

- The file is stored in the export directory root rather than split into one Word file per day.
- Each date has its own section, with a heading formatted like <code>2026.8.25</code>.
- Each date section always contains two lines: reach for that day's <code>群发客户</code> tasks and reach for that day's <code>群发朋友圈</code> tasks.
- Each line records the number of matched tasks and the reach of every task. If there is more than one task of the same type, the daily total is also included.
- Running the same date again replaces that date's existing section instead of appending a duplicate. The file is not updated if its contents have not changed.

The current <code>社群任务触达人数日报.xlsx</code> workbook contains:

- <code>触达人数汇总</code> (Reach Summary)
- <code>门店分组触达人数</code> (Reach by Store Group)
- <code>说明</code> (Notes)

Key columns in <code>触达人数汇总</code>:

| Column | Contents |
| --- | --- |
| A | Date, formatted as <code>yyyy/mm/dd</code> |
| B | Day of week, formatted from <code>星期一</code> through <code>星期日</code> |
| C | Welfare-account friends, sourced from total customer mass-message reach in the Word report |
| F | Community groups, sourced from community-task reach |
| I | Moments, sourced from total Moments-post reach in the Word report |

Key columns in <code>门店分组触达人数</code>:

| Column | Contents |
| --- | --- |
| A | Date |
| B | Type |
| C | Store group |
| D | Coupon type |
| E | Reach |

### 1.6 Business rules and metric definitions

#### Task-title filters

By default, export tasks exclude titles containing any of the following keywords:

- <code>测试</code> (test)
- <code>海外</code> (overseas)
- <code>境外</code> (outside mainland China)

This applies to:

- Undelivered Super Mass Messaging records
- Customer-group mass-message exports
- Reach for customer mass messages and Moments posts

Related configuration variables:

| Task | Configuration variable |
| --- | --- |
| Undelivered Super Mass Messaging records | <code>EXCLUDE_TASK_KEYWORDS</code> |
| Customer-group mass-message exports | <code>CUSTOMER_GROUP_EXCLUDE_KEYWORDS</code> |
| Customer mass-message and Moments reach | <code>REACH_EXCLUDE_KEYWORDS</code> |

If the exclusion-keyword setting is empty, no tasks are excluded by keyword. If inclusion keywords are configured, only task titles matching an inclusion keyword are exported.

#### Community-group reach

Community-group reach is calculated from task-export Excel files in each daily folder:

- Only rows with <code>送达状态=已送达</code> (Delivery Status = Delivered) are counted.
- <code>客户群chatid（本应用）</code> in each task file is matched to <code>客户群ID</code> in the customer-group statistics file.
- After a successful match, the corresponding <code>群客户总数</code> (total group members) is added to reach.
- If the same group appears in multiple tasks, every exposure is counted; groups are not deduplicated across tasks.
- Both <code>超级群发</code> and <code>群发客户群</code> tasks contribute to community-group reach.

#### Customer mass-message and Moments reach

Task 3 generates <code>企微社群任务触达客户统计.docx</code>. Its two totals are written to the daily report as follows:

| Word metric | Daily-report field |
| --- | --- |
| Total reach of customer mass messages | Welfare-account friends |
| Total reach of Moments posts | Moments |

#### Reach by store group

Store-group reach includes only task files whose names contain a store-group marker:

| Filename marker | Store group |
| --- | --- |
| <code>AFD</code> | <code>社群-A档</code> |
| <code>BFD</code> | <code>社群-B档</code> |
| <code>CFD</code> | <code>社群-C档</code> |
| <code>SFD</code> | <code>社群-S档</code> |

Included files:

- <code>超级群发</code> files whose names contain <code>AFD/BFD/CFD/SFD</code>
- <code>群发客户群</code> files whose names contain <code>AFD/BFD/CFD/SFD</code>

Ordinary <code>D</code> files and files without an <code>AFD/BFD/CFD/SFD</code> marker are excluded. The food-coupon row equals the sum of reach across the A, B, C, and S tiers for that date.

#### Re-runs and historical data

- The local daily report is updated by date. Recalculating the same date replaces its old rows rather than appending duplicates.
- By default, only the most recent few days are recalculated on a rolling basis; older historical dates remain unchanged.
- Feishu sync only appends local dates newer than the latest online date. It never overwrites, clears, or modifies historical online rows.

### 1.7 Feishu sync configuration

Feishu sync is disabled by default. To enable it, click <code>飞书同步配置</code> in the GUI and enter:

~~~text
App ID
App Secret
Reach Summary sheet URL
Reach by Store Group sheet URL
~~~

After <code>测试连接</code> (Test Connection) succeeds, select <code>启用飞书同步</code> (Enable Feishu Sync) and save.

<code>启用飞书同步</code> is the master switch for Task 7:

- Disabled: automatic runs execute only the first six tasks, without connecting to Feishu or writing to online spreadsheets.
- Enabled: Task 7 syncs the online spreadsheets after the first six tasks finish.

Sync is append-only:

- It first reads the dates already present in the Feishu spreadsheets.
- It appends only local-report dates newer than the latest online date.
- It never overwrites, clears, or changes historical online rows.
- After appending, it automatically applies centered alignment, date formats, and thousands separators.

The Feishu app must have permission to read and write spreadsheets. The <code>App Secret</code> is stored in <code>%LOCALAPPDATA%\ScrmDailyExporter\.env</code> on the current computer. Do not commit it to Git or distribute it with the source package.

### 1.8 Troubleshooting

**The scheduled task did not run**

Make sure the computer was powered on and the current Windows user was signed in at 09:40. You can click <code>重新安装计划任务</code> in the GUI to refresh the scheduled task.

**The login window does not open or keeps waiting for QR-code login**

The application uses Chrome debug port <code>9333</code> by default. If another application is using that port, change it in <code>%LOCALAPPDATA%\ScrmDailyExporter\.env</code>:

~~~text
CHROME_DEBUG_PORT=9334
~~~

Then click <code>扫码登录/刷新登录态</code> again.

**Writing the Excel file fails**

If the log contains <code>PermissionError</code> or says that the file is in use, <code>社群任务触达人数日报.xlsx</code> is usually open in Excel. Close the workbook and run the application again.

**Dates appear as numbers in the Feishu spreadsheet**

The current version applies formatting automatically after appending. If older data is already displayed as numbers, use the maintenance script's format-repair option to reapply the date format; see Section 2.8.

**Online-spreadsheet sync fails**

First click <code>测试连接</code> under <code>飞书同步配置</code>. If the test fails, verify the App ID, App Secret, spreadsheet URLs, app permissions, and whether the spreadsheets have been shared with the app.

## 2. Source Maintenance and Extension

### 2.1 Source structure

| File | Purpose |
| --- | --- |
| <code>app_ui.py</code> | Main GUI, task checklist, buttons, and Feishu sync settings |
| <code>app_cli.py</code> | CLI entry point for <code>run</code>, <code>login</code>, <code>status</code>, <code>install-task</code>, and <code>uninstall-task</code> |
| <code>daily_export_scheduler.py</code> | Daily scheduling and task orchestration, including task state, catch-up dates, and login refresh |
| <code>export_super_group_undelivered.py</code> | Undelivered Super Mass Messaging export and base SCRM API-client logic |
| <code>export_chat_group_analysis_by_chat.py</code> | Customer Group Analytics by group chat export |
| <code>export_reach_customer_summary.py</code> | Calculates customer mass-message and Moments reach and generates the Word report |
| <code>export_group_send_customer_group.py</code> | Customer-group mass-message export |
| <code>export_reach_daily_excel.py</code> | Generates the local <code>社群任务触达人数日报.xlsx</code> workbook |
| <code>sync_feishu_reach_workbook.py</code> | Appends local daily-report data to Feishu spreadsheets |
| <code>scrm_browser_fetch.py</code> | Fallback that sends same-origin requests through Chrome DevTools in an authenticated page |
| <code>runtime_paths.py</code> | Rules for application, configuration, log, and other runtime paths |
| <code>app_settings.py</code> | Export directory and global catch-up start date saved by the GUI |
| <code>state_file_io.py</code> | State-file input/output |
| <code>requirements.txt</code> | Python dependencies |
| <code>ScrmDailyExporter.spec</code> | Production PyInstaller configuration |
| <code>build_release.ps1</code> | Builds <code>dist\ScrmDailyExporter</code> |
| <code>build_installer.ps1</code> | Builds the installer from <code>dist</code> |
| <code>install_local.ps1</code> | Performs a quick local installation without creating an installer |
| <code>installer.iss</code> | Inno Setup installer configuration |

### 2.2 Local development environment

After receiving a standard source package, maintainers should run the following commands in the source directory:

~~~powershell
python -m venv .venv
.\.venv\Scripts\pip.exe install -r requirements.txt
~~~

Launch the GUI directly from source during development:

~~~powershell
.\.venv\Scripts\python.exe app_ui.py
~~~

To inspect the schedule without performing an actual export:

~~~powershell
.\.venv\Scripts\python.exe app_cli.py run --test-mode --plan-only
~~~

Common commands:

~~~powershell
.\.venv\Scripts\python.exe app_cli.py status
.\.venv\Scripts\python.exe app_cli.py login
.\.venv\Scripts\python.exe app_cli.py run --start-date 2026-08-01
.\.venv\Scripts\python.exe app_cli.py install-task --test-mode
.\.venv\Scripts\python.exe app_cli.py uninstall-task --test-mode
~~~

<code>--test-mode</code> uses separate test configuration, output directories, and scheduled-task names so that the production installation is not affected.

### 2.3 Building and releasing

Build the directly runnable executable directory:

~~~powershell
.\build_release.ps1
~~~

Output directory:

~~~text
dist\ScrmDailyExporter
~~~

Build the production installer:

~~~powershell
.\build_installer.ps1
~~~

Output file:

~~~text
installer-output\ScrmDailyExporterSetup.exe
~~~

Send <code>ScrmDailyExporterSetup.exe</code> to end users. A clean source package can be provided separately to maintainers.

<code>dist</code>, <code>build</code>, <code>installer-output</code>, <code>.venv</code>, and <code>.env</code> are excluded from Git. It is normal for a source package not to contain <code>dist</code>; run <code>build_release.ps1</code> when it needs to be regenerated.

### 2.4 Configuration and runtime data

All runtime configuration is stored under <code>%LOCALAPPDATA%\ScrmDailyExporter</code> for the current Windows user.

| File or directory | Contents |
| --- | --- |
| <code>.env</code> | WeCom token/cookie, Feishu App ID/Secret, Chrome port, and other sensitive configuration |
| <code>logs</code> | Detailed logs for each run |
| <code>state\export_state.json</code> | Most recent successful date for each task |
| <code>state\latest_status.txt</code> | Current GUI status summary |
| <code>app_settings.json</code> | Export directory and global catch-up start date saved by the GUI |
| <code>chrome-profile</code> | Dedicated Chrome login profile |

Do not commit or distribute <code>.env</code>. Each colleague who needs to use the application independently should scan to log in on their own computer and enter their own Feishu sync settings in the GUI.

### 2.5 When the WeCom admin APIs change

First, use the logs to identify the failed task:

~~~powershell
Get-Content -LiteralPath "$env:LOCALAPPDATA\ScrmDailyExporter\state\latest_status.txt" -Raw
~~~

Then open the latest log:

~~~powershell
Get-ChildItem "$env:LOCALAPPDATA\ScrmDailyExporter\logs" -Filter *.log |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 1
~~~

Use the failed task to locate the relevant source file:

| Failed task | Check first |
| --- | --- |
| Undelivered Super Mass Messaging records | <code>export_super_group_undelivered.py</code> |
| Customer Group Analytics by group chat | <code>export_chat_group_analysis_by_chat.py</code> |
| Customer mass-message and Moments reach | <code>export_reach_customer_summary.py</code> |
| Customer-group mass-message export | <code>export_group_send_customer_group.py</code> |
| Reach Summary or Reach by Store Group | <code>export_reach_daily_excel.py</code> |
| Online-spreadsheet sync | <code>sync_feishu_reach_workbook.py</code> |

General procedure for investigating an API change:

1. Use <code>扫码登录/刷新登录态</code> in the GUI to open the dedicated Chrome window.
2. Manually perform the corresponding admin action in Chrome DevTools' Network panel.
3. Compare the new request URL, payload, headers, and response fields.
4. Update the endpoint, request parameters, or field parsing in the corresponding <code>export_*.py</code> file.
5. Check the task plan with <code>--test-mode --plan-only</code>.
6. Validate a real export for a small date range.
7. After validation, run <code>build_release.ps1</code> and <code>build_installer.ps1</code> again.

An export failure does not necessarily mean that an API changed. Try scanning to log in again first. Log messages containing <code>WinError 10061</code>, <code>127.0.0.1:9333</code>, or <code>Chrome debug port</code> usually indicate that the Chrome debugging window is not running or that its port conflicts with another application.

### 2.6 When the Feishu API or spreadsheet format changes

Check <code>sync_feishu_reach_workbook.py</code> first.

Key logic includes:

- Parsing Feishu spreadsheet URLs
- Obtaining a <code>tenant_access_token</code>
- Converting a wiki token to a spreadsheet token
- Reading dates already present in the online spreadsheets
- Determining whether only new dates should be appended
- Writing <code>触达人数汇总</code>
- Writing <code>门店分组触达人数</code>
- Applying date formats, thousands separators, and centered alignment after appending

If a spreadsheet sheet is replaced, end users can update its URL directly under <code>飞书同步配置</code> in the GUI without modifying source code.

If the online column structure changes, update all of the following:

- <code>SUMMARY_COLUMNS</code>
- <code>STORE_GROUP_COLUMNS</code>
- <code>read_local_summary_rows</code>
- <code>read_local_store_group_rows</code>
- <code>format_summary_rows</code>
- <code>format_store_group_rows</code>

### 2.7 Adding a task

Adding a task generally requires changes in all of the following places:

1. Add or extend the relevant export or aggregation script.
2. Add the task ID, label, execution function, and dependency order in <code>daily_export_scheduler.py</code>.
3. Add the task to the <code>TASKS</code> checklist in <code>app_ui.py</code>.
4. Add state migration or initialization logic if the task requires it.
5. Update the README task list, output documentation, and maintenance instructions.
6. Check the task plan with <code>--test-mode --plan-only</code>.
7. Validate the task on a small date range.

Avoid renaming a task ID after release because <code>state\export_state.json</code> records completion dates by task ID.

### 2.8 Debugging and maintenance commands

Syntax check:

~~~powershell
.\.venv\Scripts\python.exe -m compileall -q .
~~~

Check the installed application's status:

~~~powershell
& "$env:LOCALAPPDATA\ScrmDailyExporter\app\scrm-exporter.exe" status
~~~

Inspect the installed application's task plan:

~~~powershell
& "$env:LOCALAPPDATA\ScrmDailyExporter\app\scrm-exporter.exe" run --plan-only
~~~

Catch up from a specified date:

~~~powershell
& "$env:LOCALAPPDATA\ScrmDailyExporter\app\scrm-exporter.exe" run --start-date 2026-08-01
~~~

Repair Feishu formatting for one date only:

~~~powershell
.\.venv\Scripts\python.exe .\sync_feishu_reach_workbook.py \
  --workbook "$env:USERPROFILE\Documents\每日企微私域任务导出\社群任务触达人数日报.xlsx" \
  --env "$env:LOCALAPPDATA\ScrmDailyExporter\.env" \
  --date 2026-08-16 \
  --format-only
~~~

Dry-run Feishu sync:

~~~powershell
.\.venv\Scripts\python.exe .\sync_feishu_reach_workbook.py \
  --workbook "$env:USERPROFILE\Documents\每日企微私域任务导出\社群任务触达人数日报.xlsx" \
  --env "$env:LOCALAPPDATA\ScrmDailyExporter\.env" \
  --date 2026-08-16 \
  --dry-run
~~~

---

<a id="chinese"></a>

[English](#english) | [中文](#chinese)

# 企微社群任务自动导出工具

这是一个可安装的 Windows GUI 工具，用于每天自动导出企微社群相关任务、生成本地统计文件，并可选把触达人数日报追加同步到飞书在线表。

使用同事不需要提前安装 Python。维护源码、重新构建或修改逻辑时才需要 Python 开发环境。

## 1. 当前 GUI 使用说明

### 1.1 工具当前做什么

安装后，工具会按天补导未完成日期，并按固定顺序执行 7 个任务：

1. 超级群发未送达
2. 客户群分析-按群聊
3. 群发客户及朋友圈触达人数
4. 群发客户群导出
5. 触达人数汇总
6. 门店分组触达人数
7. 同步在线表

第 1-4 个任务负责从企微后台导出原始文件和 Word 统计文档；第 5-6 个任务基于本地导出结果生成 `社群任务触达人数日报.xlsx`；第 7 个任务在启用飞书同步后，把本地日报里比在线表更新的日期追加写入飞书。

### 1.2 安装和首次使用

普通使用同事只需要运行安装包：

```text
ScrmDailyExporterSetup.exe
```

安装后会自动完成：

- 程序安装到 `%LOCALAPPDATA%\ScrmDailyExporter\app`
- 创建开始菜单入口 `SCRM Daily Exporter`
- 注册当前 Windows 用户的计划任务 `每日企微私域任务导出`
- 默认每天 09:40 打开 GUI 并自动运行一次

首次使用时，打开 `企微社群任务自动导出`，点击 `扫码登录/刷新登录态`，在打开的 Chrome 窗口里扫码登录企微后台。

登录态只保存在当前电脑，不会随安装包或源码分发。换电脑、换账号、登录过期或企微后台要求重新登录时，都需要重新扫码。

### 1.3 默认路径

| 内容 | 路径 |
| --- | --- |
| 程序安装目录 | `%LOCALAPPDATA%\ScrmDailyExporter\app` |
| 运行配置目录 | `%LOCALAPPDATA%\ScrmDailyExporter` |
| 日志目录 | `%LOCALAPPDATA%\ScrmDailyExporter\logs` |
| 状态目录 | `%LOCALAPPDATA%\ScrmDailyExporter\state` |
| GUI 设置文件 | `%LOCALAPPDATA%\ScrmDailyExporter\app_settings.json` |
| 登录浏览器资料目录 | `%LOCALAPPDATA%\ScrmDailyExporter\chrome-profile` |
| 默认导出目录 | `%USERPROFILE%\Documents\每日企微私域任务导出` |
| 触达人数日报 | `%USERPROFILE%\Documents\每日企微私域任务导出\社群任务触达人数日报.xlsx` |

### 1.4 GUI 按钮和输入项

| 控件 | 用途 |
| --- | --- |
| 扫码登录/刷新登录态 | 打开专用 Chrome 登录窗口，扫码后刷新本机 `.env` 中的企微登录信息 |
| 立即运行一次 | 立即补导当前未完成任务 |
| 打开导出文件夹 | 打开本地导出结果目录 |
| 打开日志文件夹 | 打开运行日志目录 |
| 飞书同步配置 | 配置 App ID、App Secret 和两个飞书在线表 sheet 链接 |
| 重新安装计划任务 | 重新注册每日 09:40 自动运行任务 |
| 卸载计划任务 | 删除每日自动运行任务，不删除程序和历史导出文件 |
| 从指定日期开始导出 | 从输入日期跑到昨天，已成功任务会跳过 |
| 全局自动补导起始日期 | 控制自动补导从哪一天开始；清空后恢复默认最近 7 天补导 |

### 1.5 输出文件

导出目录里会包含每天的日期文件夹，以及跨日期维护的汇总文件。每天的原始 Excel 默认放在类似下面的日期文件夹中：

```text
%USERPROFILE%\Documents\每日企微私域任务导出\社群任务0825
```

常见输出包括：

| 文件 | 位置 | 来源 |
| --- | --- | --- |
| 超级群发相关 Excel | 每天日期文件夹 | 任务 1，从企微后台导出 |
| 客户群分析-按群聊相关 Excel | 每天日期文件夹 | 任务 2，从企微后台导出 |
| `企微社群任务触达客户统计.docx` | 导出目录根目录 | 任务 3，统计群发客户和朋友圈触达人数 |
| 群发客户群相关 Excel | 每天日期文件夹 | 任务 4，从企微后台导出 |
| `社群任务触达人数日报.xlsx` | 导出目录根目录 | 任务 5-6，写入触达人数汇总和门店分组触达人数 |

`企微社群任务触达客户统计.docx` 写入规则：

- 文件默认写在导出目录根目录，不按天拆成多个 Word。
- 每个日期写一个日期块，日期标题格式类似 `2026.8.25`。
- 日期块下面固定写两行：当天 `群发客户` 任务触达人数、当天 `群发朋友圈` 任务触达人数。
- 每行会写命中的任务数量和每个任务的触达人数；如果同类任务超过 1 个，会追加当天合计触达人数。
- 同一天重新运行时，会替换 Word 里该日期已有内容，避免重复追加；如果内容没有变化，则不更新文件。

`社群任务触达人数日报.xlsx` 当前包含：

- `触达人数汇总`
- `门店分组触达人数`
- `说明`

`触达人数汇总` 的核心列：

| 列 | 内容 |
| --- | --- |
| A | 日期，格式 `yyyy/mm/dd` |
| B | 星期，格式 `星期一` 到 `星期日` |
| C | 福利官好友，来自 Word 里的群发客户触达总人数 |
| F | 社群，来自社群任务触达人数 |
| I | 朋友圈，来自 Word 里的群发朋友圈触达总人数 |

`门店分组触达人数` 的核心列：

| 列 | 内容 |
| --- | --- |
| A | 日期 |
| B | 类型 |
| C | 门店分组 |
| D | 券类型 |
| E | 触达人数 |

### 1.6 业务规则和统计口径

#### 任务标题筛选

以下导出任务默认排除标题包含这些关键字的任务：

- `测试`
- `海外`
- `境外`

适用范围：

- 超级群发未送达
- 群发客户群导出
- 群发客户及朋友圈触达人数

对应配置项：

| 任务 | 配置项 |
| --- | --- |
| 超级群发未送达 | `EXCLUDE_TASK_KEYWORDS` |
| 群发客户群导出 | `CUSTOMER_GROUP_EXCLUDE_KEYWORDS` |
| 群发客户及朋友圈触达人数 | `REACH_EXCLUDE_KEYWORDS` |

如果排除关键字配置为空，则不按关键字排除。如果配置了包含关键字，则只导出标题命中包含关键字的任务。

#### 社群触达人数

社群触达人数来自每天日期文件夹里的任务导出 Excel。统计规则：

- 只统计 `送达状态=已送达`
- 用任务文件里的 `客户群chatid（本应用）` 匹配客户群统计表里的 `客户群ID`
- 匹配成功后，加和对应的 `群客户总数`
- 同一个群如果出现在多个任务中，按多次触达累计，不去重
- `超级群发` 和 `群发客户群` 两类任务都会计入社群触达人数

#### 群发客户及朋友圈触达人数

任务 3 会生成 `企微社群任务触达客户统计.docx`。Word 文档里的两个总数会写入日报：

| Word 统计项 | 日报字段 |
| --- | --- |
| 群发客户触达总人数 | 福利官好友 |
| 群发朋友圈触达总人数 | 朋友圈 |

#### 门店分组触达人数

门店分组只统计文件名包含门店分组标志的任务：

| 文件名标志 | 门店分组 |
| --- | --- |
| `AFD` | `社群-A档` |
| `BFD` | `社群-B档` |
| `CFD` | `社群-C档` |
| `SFD` | `社群-S档` |

统计范围：

- 文件名带 `AFD/BFD/CFD/SFD` 的 `超级群发`
- 文件名带 `AFD/BFD/CFD/SFD` 的 `群发客户群`

不统计普通 `D` 文件，也不统计不带 `AFD/BFD/CFD/SFD` 标志的文件。食品券行等于当天 A/B/C/S 四档触达人数之和。

#### 重复运行和历史数据

- 本地日报按日期更新，同一天重新计算时会覆盖本地日报里的当天旧行，避免重复追加。
- 默认只滚动重算最近几天，较早历史日期保留不变。
- 飞书同步只追加本地日报里比在线表更新的日期，不覆盖、不清空、不修改在线表历史行。

### 1.7 飞书同步配置

飞书同步默认关闭。需要同步时，在 GUI 里点击 `飞书同步配置`，填写：

```text
App ID
App Secret
触达人数汇总 sheet 链接
门店分组触达人数 sheet 链接
```

点击 `测试连接` 成功后，勾选 `启用飞书同步` 并保存。

`启用飞书同步` 是第 7 个任务的总开关：

- 不勾选：自动运行只做前 6 个任务，不连接飞书，也不写在线表。
- 勾选：前 6 个任务完成后，第 7 个任务会同步在线表。

同步规则是追加式：

- 先读取飞书在线表已有日期
- 只追加本地日报里比在线表最新日期更新的数据
- 不覆盖、不清空、不修改线上历史行
- 追加后自动设置居中、日期格式和数字千分位

飞书 App 需要具备读取和写入电子表格的权限。`App Secret` 会保存到当前电脑的 `%LOCALAPPDATA%\ScrmDailyExporter\.env`，不要提交到 Git，也不要随源码包转发。

### 1.8 常见问题

**自动任务没有运行**

确认电脑在 09:40 时已开机并登录当前 Windows 用户。可以在 GUI 里点 `重新安装计划任务` 刷新计划任务。

**登录窗口打不开或一直等待扫码**

工具默认使用 Chrome debug port `9333`。如果被其他程序占用，可以在 `%LOCALAPPDATA%\ScrmDailyExporter\.env` 里修改：

```text
CHROME_DEBUG_PORT=9334
```

改完后重新点击 `扫码登录/刷新登录态`。

**Excel 文件写入失败**

如果日志提示 `PermissionError` 或文件被占用，通常是 `社群任务触达人数日报.xlsx` 正在被 Excel 打开。关闭文件后重新运行即可。

**飞书在线表日期显示成数字**

当前版本会在追加后自动套格式。如果旧数据已经写成数字，可以用维护脚本的格式修复入口按日期补格式，见第 2.8 节。

**同步在线表失败**

先在 GUI 的 `飞书同步配置` 里点 `测试连接`。如果测试失败，检查 App ID、App Secret、在线表链接、应用权限、在线表是否授权给该应用。

## 2. 源码维护和拓展说明

### 2.1 源码结构

| 文件 | 作用 |
| --- | --- |
| `app_ui.py` | GUI 主界面，任务 checklist、按钮、飞书同步配置窗口 |
| `app_cli.py` | 命令行入口，负责 run、login、status、install-task、uninstall-task |
| `daily_export_scheduler.py` | 每日调度和任务编排，维护任务状态、补导日期、登录刷新 |
| `export_super_group_undelivered.py` | 超级群发未送达导出，以及 SCRM API 客户端基础逻辑 |
| `export_chat_group_analysis_by_chat.py` | 客户群分析-按群聊导出 |
| `export_reach_customer_summary.py` | 统计群发客户和朋友圈触达人数，生成 Word 文档 |
| `export_group_send_customer_group.py` | 群发客户群导出 |
| `export_reach_daily_excel.py` | 生成本地 `社群任务触达人数日报.xlsx` |
| `sync_feishu_reach_workbook.py` | 把本地日报追加同步到飞书在线表 |
| `scrm_browser_fetch.py` | 通过 Chrome DevTools 在已登录页面里发同源请求的 fallback |
| `runtime_paths.py` | 安装目录、配置目录、日志目录等路径规则 |
| `app_settings.py` | GUI 保存的导出目录、全局补导起始日期等设置 |
| `state_file_io.py` | 状态文件读写 |
| `requirements.txt` | Python 依赖 |
| `ScrmDailyExporter.spec` | PyInstaller 正式打包配置 |
| `build_release.ps1` | 生成 `dist\ScrmDailyExporter` |
| `build_installer.ps1` | 基于 `dist` 生成安装包 |
| `install_local.ps1` | 本机快速安装，不生成安装包 |
| `installer.iss` | Inno Setup 安装包配置 |

### 2.2 本地开发环境

维护源码的同事拿到普通源码包后，先在源码目录执行：

```powershell
python -m venv .venv
.\.venv\Scripts\pip.exe install -r requirements.txt
```

开发时可以直接从源码启动 GUI：

```powershell
.\.venv\Scripts\python.exe app_ui.py
```

也可以只测试调度计划，不实际导出：

```powershell
.\.venv\Scripts\python.exe app_cli.py run --test-mode --plan-only
```

常用命令：

```powershell
.\.venv\Scripts\python.exe app_cli.py status
.\.venv\Scripts\python.exe app_cli.py login
.\.venv\Scripts\python.exe app_cli.py run --start-date 2026-08-01
.\.venv\Scripts\python.exe app_cli.py install-task --test-mode
.\.venv\Scripts\python.exe app_cli.py uninstall-task --test-mode
```

`--test-mode` 会使用独立的测试配置目录、测试输出目录和测试计划任务名，避免影响正式安装版。

### 2.3 构建和发版

生成可直接运行的 exe 目录：

```powershell
.\build_release.ps1
```

输出目录：

```text
dist\ScrmDailyExporter
```

生成正式安装包：

```powershell
.\build_installer.ps1
```

输出文件：

```text
installer-output\ScrmDailyExporterSetup.exe
```

发给普通使用同事时，发 `ScrmDailyExporterSetup.exe`。发给维护同事时，可以另发一份干净源码包。

`dist`、`build`、`installer-output`、`.venv`、`.env` 都不进 Git。源码包里没有 `dist` 是正常的，需要时重新运行 `build_release.ps1` 生成。

### 2.4 配置和运行数据

运行配置都在当前 Windows 用户自己的 `%LOCALAPPDATA%\ScrmDailyExporter` 下。

| 文件或目录 | 内容 |
| --- | --- |
| `.env` | 企微 token/cookie、飞书 App ID/Secret、Chrome 端口等敏感配置 |
| `logs` | 每次运行的详细日志 |
| `state\export_state.json` | 每个任务最近成功日期 |
| `state\latest_status.txt` | GUI 当前状态摘要 |
| `app_settings.json` | GUI 保存的导出目录和全局补导起始日期 |
| `chrome-profile` | 登录专用 Chrome profile |

不要提交或转发 `.env`。如果同事需要独立使用，让同事在自己电脑上扫码登录，并在 GUI 中填写自己的飞书同步配置。

### 2.5 如果企微后台接口变化了

先看日志定位失败任务：

```powershell
Get-Content -LiteralPath "$env:LOCALAPPDATA\ScrmDailyExporter\state\latest_status.txt" -Raw
```

再打开最新日志：

```powershell
Get-ChildItem "$env:LOCALAPPDATA\ScrmDailyExporter\logs" -Filter *.log |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 1
```

根据失败任务找对应源码：

| 失败任务 | 优先检查文件 |
| --- | --- |
| 超级群发未送达 | `export_super_group_undelivered.py` |
| 客户群分析-按群聊 | `export_chat_group_analysis_by_chat.py` |
| 群发客户及朋友圈触达人数 | `export_reach_customer_summary.py` |
| 群发客户群导出 | `export_group_send_customer_group.py` |
| 触达人数汇总、门店分组触达人数 | `export_reach_daily_excel.py` |
| 同步在线表 | `sync_feishu_reach_workbook.py` |

排查接口变化的一般流程：

1. 用 GUI 的 `扫码登录/刷新登录态` 打开专用 Chrome。
2. 在 Chrome DevTools 的 Network 面板里手动执行对应后台操作。
3. 对比新的 request URL、payload、headers、response 字段。
4. 更新对应 `export_*.py` 里的接口路径、请求参数或字段解析。
5. 用 `--test-mode --plan-only` 检查调度计划。
6. 用指定日期小范围验证真实导出。
7. 通过后重新 `build_release.ps1` 和 `build_installer.ps1`。

如果只是企微登录态失效，不一定是接口变化。优先重新扫码；如果日志包含 `WinError 10061`、`127.0.0.1:9333`、`Chrome debug port`，通常是 Chrome 调试窗口没启动或端口冲突。

### 2.6 如果飞书接口或在线表格式变化了

优先检查 `sync_feishu_reach_workbook.py`。

重点逻辑包括：

- 解析飞书在线表链接
- 获取 `tenant_access_token`
- wiki token 转 spreadsheet token
- 读取在线表已有日期
- 判断是否只追加新日期
- 写入 `触达人数汇总`
- 写入 `门店分组触达人数`
- 追加后设置日期、千分位和居中格式

如果在线表 sheet 更换，普通使用同事可以直接在 GUI 的 `飞书同步配置` 里换链接，不需要改源码。

如果在线表列结构变化，需要同步修改：

- `SUMMARY_COLUMNS`
- `STORE_GROUP_COLUMNS`
- `read_local_summary_rows`
- `read_local_store_group_rows`
- `format_summary_rows`
- `format_store_group_rows`

### 2.7 如果要新增一个任务

新增任务通常需要同时改这些地方：

1. 新增或扩展对应的导出/统计脚本。
2. 在 `daily_export_scheduler.py` 中增加 task id、label、运行函数和依赖顺序。
3. 在 `app_ui.py` 的 `TASKS` 中增加 checklist 展示。
4. 如果任务有状态初始化逻辑，补充 state migration 或初始化函数。
5. 更新 README 的任务列表、输出文件说明和维护说明。
6. 用 `--test-mode --plan-only` 检查任务计划。
7. 小范围跑指定日期验证。

任务 id 一旦发布后尽量不要随意改名，因为 `state\export_state.json` 会按 task id 记录成功日期。

### 2.8 调试和维护命令

语法检查：

```powershell
.\.venv\Scripts\python.exe -m compileall -q .
```

查看安装版状态：

```powershell
& "$env:LOCALAPPDATA\ScrmDailyExporter\app\scrm-exporter.exe" status
```

安装版计划检查：

```powershell
& "$env:LOCALAPPDATA\ScrmDailyExporter\app\scrm-exporter.exe" run --plan-only
```

从指定日期补导：

```powershell
& "$env:LOCALAPPDATA\ScrmDailyExporter\app\scrm-exporter.exe" run --start-date 2026-08-01
```

只修复某天飞书线上格式：

```powershell
.\.venv\Scripts\python.exe .\sync_feishu_reach_workbook.py `
  --workbook "$env:USERPROFILE\Documents\每日企微私域任务导出\社群任务触达人数日报.xlsx" `
  --env "$env:LOCALAPPDATA\ScrmDailyExporter\.env" `
  --date 2026-08-16 `
  --format-only
```

飞书同步 dry-run：

```powershell
.\.venv\Scripts\python.exe .\sync_feishu_reach_workbook.py `
  --workbook "$env:USERPROFILE\Documents\每日企微私域任务导出\社群任务触达人数日报.xlsx" `
  --env "$env:LOCALAPPDATA\ScrmDailyExporter\.env" `
  --date 2026-08-16 `
  --dry-run
```
