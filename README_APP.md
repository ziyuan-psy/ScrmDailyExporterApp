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

## 日常使用

1. 打开“企微社群任务自动导出”。
2. 首次使用点击“扫码登录/刷新登录态”，在打开的 Chrome 或 Edge 窗口扫码登录。
3. 如需更换保存位置，点击“选择导出目录”；保存后手动导出和计划任务都会使用新目录。
4. “全局自动补导起始日期”默认是昨天；可填今天或未来日期来避免首次运行补导历史，清空后恢复最近 7 天补导。
5. 点击“立即运行一次”可手动补导未完成任务；控制台会显示 4 个任务的 checklist 和最近日志。

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
- 如果 SCRM 后台接口或登录机制变化，需要更新工具版本。
