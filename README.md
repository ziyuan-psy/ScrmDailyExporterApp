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
| 登录浏览器资料目录 | `%LOCALAPPDATA%\ScrmDailyExporter\chrome-login-profile` |
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

每天会生成一个日期文件夹，默认命名类似：

```text
%USERPROFILE%\Documents\每日企微私域任务导出\社群任务0825
```

常见输出包括：

| 文件 | 来源 |
| --- | --- |
| 超级群发相关 Excel | 任务 1，从企微后台导出 |
| 客户群分析-按群聊相关 Excel | 任务 2，从企微后台导出 |
| `企微社群任务触达客户统计.docx` | 任务 3，统计群发客户和朋友圈触达人数 |
| 群发客户群相关 Excel | 任务 4，从企微后台导出 |
| `社群任务触达人数日报.xlsx` | 任务 5-6，写入触达人数汇总和门店分组触达人数 |

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

### 1.6 飞书同步配置

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

### 1.7 常见问题

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
| `chrome-login-profile` | 登录专用 Chrome profile |

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
