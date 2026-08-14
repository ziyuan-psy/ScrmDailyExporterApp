# SCRM Daily Exporter

This is an installable Windows tool for daily SCRM community-task exports. It does not require Codex, and coworkers do not need to install Python before using it.

## Default Paths

- App installation directory: `%LOCALAPPDATA%\ScrmDailyExporter\app`
- Runtime configuration directory: `%LOCALAPPDATA%\ScrmDailyExporter`
- Logs directory: `%LOCALAPPDATA%\ScrmDailyExporter\logs`
- State directory: `%LOCALAPPDATA%\ScrmDailyExporter\state`
- App settings file: `%LOCALAPPDATA%\ScrmDailyExporter\app_settings.json`
- Login browser profile directory: `%LOCALAPPDATA%\ScrmDailyExporter\chrome-profile`
- Export output directory: `%USERPROFILE%\Documents\每日企微私域任务导出`
- Reach report workbook: `%USERPROFILE%\Documents\每日企微私域任务导出\社群任务触达人数日报.xlsx`

## Daily Use

1. Open "企微社群任务自动导出".
2. On first use, click "扫码登录/刷新登录态" and scan the QR code in the Chrome or Edge window.
3. To change the save location, click "选择导出目录". Manual exports and scheduled exports will both use the saved directory.
4. "全局自动补导起始日期" defaults to yesterday. Enter today or a future date to avoid backfilling historical data on first run, or clear it to restore the latest 7-day catch-up window.
5. Click "立即运行一次" to manually export unfinished tasks. The UI shows a 7-task checklist and the latest logs.

The 7 tasks run in this order: super-group undelivered export, customer-group analysis by chat, customer and Moments reach count, group-send customer-group export, reach summary, store-group reach summary, and Feishu online spreadsheet sync. Tasks 5 and 6 write to the local workbook `社群任务触达人数日报.xlsx`; task 7 appends only new dates from that workbook to Feishu and never overwrites existing online rows.

## Scheduled Run

The installer registers a scheduled task for the current Windows user: `每日企微私域任务导出`.

By default, the app UI opens and runs once every day at 09:40. If the computer is off, the user is not logged in, or a task fails, the next run will automatically catch up unfinished dates.

## Build And Install

Run this from the project directory:

```powershell
.\build_release.ps1
```

Build output:

```text
dist\ScrmDailyExporter
```

If Inno Setup is not installed, install locally with:

```powershell
.\install_local.ps1
```

This copies the app to `%LOCALAPPDATA%\ScrmDailyExporter\app`, creates Start Menu shortcuts, and registers the daily scheduled task.

If Inno Setup 6 is installed on the build machine, run:

```powershell
.\build_installer.ps1
```

The installer is generated at:

```text
installer-output\ScrmDailyExporterSetup.exe
```

When debugging the UI or scheduler logic, you usually do not need to rebuild the installer. Prefer running the source directly:

```powershell
.\.venv\Scripts\python.exe app_ui.py
.\.venv\Scripts\python.exe app_cli.py run --test-mode --plan-only
```

To verify executable behavior, run `build_release.ps1` first and test the programs in `dist\ScrmDailyExporter`. Only run `build_installer.ps1` when validating installation flow, scheduled task registration, or preparing a package for coworkers.

## Command Line

```powershell
scrm-exporter.exe run
scrm-exporter.exe login
scrm-exporter.exe status
scrm-exporter.exe install-task
scrm-exporter.exe uninstall-task
scrm-exporter.exe run --start-date 2026-07-01
scrm-exporter.exe install-task --test-mode
scrm-exporter.exe uninstall-task --test-mode
scrm-exporter-ui.exe --auto-run
scrm-exporter-ui.exe --test-mode
```

You can also specify directories:

```powershell
scrm-exporter.exe run --config-dir "%LOCALAPPDATA%\ScrmDailyExporter" --data-dir "%USERPROFILE%\Documents\每日企微私域任务导出"
```

## Feishu Sync

Feishu sync is disabled by default. Open "飞书同步配置" in the app, then fill the Feishu `App ID`, `App Secret`, and the two online sheet links. The app saves these values to `%LOCALAPPDATA%\ScrmDailyExporter\.env` for the current Windows user.

The app parses the document token and sheet IDs from the links automatically:

```text
App ID
App Secret
触达人数汇总 sheet link
门店分组触达人数 sheet link
```

Advanced/manual `.env` configuration is also supported:

```text
FEISHU_SYNC_ENABLED=1
FEISHU_APP_ID=cli_xxxxx
FEISHU_APP_SECRET=paste_app_secret_here
FEISHU_SUMMARY_SHEET_URL=https://example.feishu.cn/wiki/xxx?sheet=ba5a9a
FEISHU_STORE_GROUP_SHEET_URL=https://example.feishu.cn/wiki/xxx?sheet=6Dks8k
FEISHU_WIKI_TOKEN=
FEISHU_SUMMARY_SHEET_ID=ba5a9a
FEISHU_STORE_GROUP_SHEET_ID=6Dks8k
FEISHU_SCAN_ROWS=5000
```

The sync reads the local workbook and the online sheets first. For each target date, it appends rows only when the date is newer than the latest date already present in the corresponding online sheet. Existing local workbook rows and existing Feishu rows are not modified.

## Test Mode

Use test mode for local validation to avoid overwriting the production scheduled task:

```powershell
scrm-exporter.exe install-task --test-mode
scrm-exporter.exe run --test-mode --plan-only
scrm-exporter-ui.exe --test-mode
```

Test mode uses the scheduled task name `每日企微私域任务导出-App测试` and separate runtime/output directories.

## Export From A Specific Date

The UI supports exporting from a specified start date. The command line supports the same flow:

```powershell
scrm-exporter.exe run --start-date 2026-07-01
```

The exporter runs from the specified date through yesterday. Tasks already marked successful are skipped and are not forcibly re-exported.

## Notes

- Login state is stored only on the current user's computer and is not distributed with the installer.
- The tool does not bypass account permissions. Exportable data depends on the account used for QR login.
- Historical backups for `企微社群任务触达客户统计.docx` are disabled by default. To enable them, set `REACH_DOCX_BACKUP_ENABLED=1` in `.env`.
- Feishu app secrets must stay in the local `.env` file and should not be committed.
- If the SCRM backend API or login mechanism changes, the tool needs to be updated.

---

# 企微社群任务自动导出工具

这是可安装版工具，不依赖 Codex，也不要求同事电脑提前安装 Python。

## 默认目录

- 程序安装目录：`%LOCALAPPDATA%\ScrmDailyExporter\app`
- 运行配置目录：`%LOCALAPPDATA%\ScrmDailyExporter`
- 日志目录：`%LOCALAPPDATA%\ScrmDailyExporter\logs`
- 状态目录：`%LOCALAPPDATA%\ScrmDailyExporter\state`
- App 设置文件：`%LOCALAPPDATA%\ScrmDailyExporter\app_settings.json`
- 登录浏览器资料目录：`%LOCALAPPDATA%\ScrmDailyExporter\chrome-profile`
- 导出结果目录：`%USERPROFILE%\Documents\每日企微私域任务导出`
- 触达人数日报：`%USERPROFILE%\Documents\每日企微私域任务导出\社群任务触达人数日报.xlsx`

## 日常使用

1. 打开“企微社群任务自动导出”。
2. 首次使用点击“扫码登录/刷新登录态”，在打开的 Chrome 或 Edge 窗口扫码登录。
3. 如需更换保存位置，点击“选择导出目录”；保存后手动导出和计划任务都会使用新目录。
4. “全局自动补导起始日期”默认是昨天；可填今天或未来日期来避免首次运行补导历史，清空后恢复最近 7 天补导。
5. 点击“立即运行一次”可手动补导未完成任务；控制台会显示 7 个任务的 checklist 和最近日志。

7 个任务依次为：超级群发未送达、客户群分析-按群聊、群发客户及朋友圈触达人数、群发客户群导出、触达人数汇总、门店分组触达人数、同步在线表。第 5、6 个任务会写入本地 `社群任务触达人数日报.xlsx`，第 7 个任务只把本地表里比在线表更新的日期追加到飞书，不覆盖在线表已有行。

## 自动运行

安装包会注册当前 Windows 用户的计划任务：`每日企微私域任务导出`。

默认每天 09:40 打开 App UI 并自动执行一次导出。电脑关机、用户未登录或任务失败时，下次运行会自动补导未成功日期。

## 构建和安装

在项目目录运行：

```powershell
.\build_release.ps1
```

构建结果在：

```text
dist\ScrmDailyExporter
```

没有 Inno Setup 时，可以直接运行：

```powershell
.\install_local.ps1
```

它会复制程序到 `%LOCALAPPDATA%\ScrmDailyExporter\app`，创建开始菜单入口，并注册每日计划任务。

如果构建机安装了 Inno Setup 6，可以运行：

```powershell
.\build_installer.ps1
```

生成正式安装包：

```text
installer-output\ScrmDailyExporterSetup.exe
```

调试 UI 或调度逻辑时，不需要每次重新打安装包。优先直接运行源码：

```powershell
.\.venv\Scripts\python.exe app_ui.py
.\.venv\Scripts\python.exe app_cli.py run --test-mode --plan-only
```

需要验证 exe 行为时先运行 `build_release.ps1`，直接测试 `dist\ScrmDailyExporter` 里的程序；只有验证安装流程、计划任务注册或发给同事时，再运行 `build_installer.ps1`。

## 命令行

```powershell
scrm-exporter.exe run
scrm-exporter.exe login
scrm-exporter.exe status
scrm-exporter.exe install-task
scrm-exporter.exe uninstall-task
scrm-exporter.exe run --start-date 2026-07-01
scrm-exporter.exe install-task --test-mode
scrm-exporter.exe uninstall-task --test-mode
scrm-exporter-ui.exe --auto-run
scrm-exporter-ui.exe --test-mode
```

可以指定目录：

```powershell
scrm-exporter.exe run --config-dir "%LOCALAPPDATA%\ScrmDailyExporter" --data-dir "%USERPROFILE%\Documents\每日企微私域任务导出"
```

## 飞书同步

飞书同步默认关闭。同事打开 app 里的“飞书同步配置”，填写飞书 `App ID`、`App Secret` 和两条在线表链接即可；程序会把这些配置保存到当前 Windows 用户的 `%LOCALAPPDATA%\ScrmDailyExporter\.env`。

```text
App ID
App Secret
触达人数汇总 sheet 链接
门店分组触达人数 sheet 链接
```

程序会自动从链接里解析文档 token 和两个 sheet id。也可以手动在 `.env` 里配置：

```text
FEISHU_SYNC_ENABLED=1
FEISHU_APP_ID=cli_xxxxx
FEISHU_APP_SECRET=paste_app_secret_here
FEISHU_SUMMARY_SHEET_URL=https://example.feishu.cn/wiki/xxx?sheet=ba5a9a
FEISHU_STORE_GROUP_SHEET_URL=https://example.feishu.cn/wiki/xxx?sheet=6Dks8k
FEISHU_WIKI_TOKEN=
FEISHU_SUMMARY_SHEET_ID=ba5a9a
FEISHU_STORE_GROUP_SHEET_ID=6Dks8k
FEISHU_SCAN_ROWS=5000
```

同步任务会先读取本地 workbook 和飞书在线表。对每个目标日期，只有当这个日期比对应在线 sheet 里已有的最大日期更新时才追加写入；本地旧行和飞书旧行都不会被修改。

## 测试模式

本机测试时可使用测试模式，避免覆盖已有生产计划任务：

```powershell
scrm-exporter.exe install-task --test-mode
scrm-exporter.exe run --test-mode --plan-only
scrm-exporter-ui.exe --test-mode
```

测试模式使用计划任务名 `每日企微私域任务导出-App测试`，运行目录和导出目录也与正式模式分开。

## 指定日期导出

控制台支持“从指定日期开始导出”，命令行也可使用：

```powershell
scrm-exporter.exe run --start-date 2026-07-01
```

导出会从指定日期跑到昨天，已成功的任务会跳过，不会强制重导。

## 注意事项

- 登录态只保存在当前用户电脑上，不随安装包分发。
- 工具不会绕过账号权限，能导出什么取决于当前扫码账号权限。
- 默认不再为 `企微社群任务触达客户统计.docx` 生成历史备份；如需恢复备份，可在 `.env` 设置 `REACH_DOCX_BACKUP_ENABLED=1`。
- 飞书应用密钥只应保存在本机 `.env` 文件里，不要提交到仓库。
- 如果 SCRM 后台接口或登录机制变化，需要更新工具版本。



