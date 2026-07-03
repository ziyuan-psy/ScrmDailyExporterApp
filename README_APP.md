# 企微社群任务自动导出工具

这是可安装版工具，不依赖 Codex，也不要求同事电脑提前安装 Python。

## 默认目录

- 程序安装目录：`%LOCALAPPDATA%\ScrmDailyExporter\app`
- 运行配置目录：`%LOCALAPPDATA%\ScrmDailyExporter`
- 日志目录：`%LOCALAPPDATA%\ScrmDailyExporter\logs`
- 状态目录：`%LOCALAPPDATA%\ScrmDailyExporter\state`
- 登录浏览器资料目录：`%LOCALAPPDATA%\ScrmDailyExporter\chrome-profile`
- 导出结果目录：`%USERPROFILE%\Documents\每日企微社群任务导出`

## 日常使用

1. 打开“企微社群任务自动导出”。
2. 首次使用点击“扫码登录/刷新登录态”，在打开的 Chrome 或 Edge 窗口扫码登录。
3. 点击“立即运行一次”可手动补导昨天及之前未完成的任务。
4. 控制台会显示 4 个任务的 checklist 和最近日志。

## 自动运行

安装包会注册当前 Windows 用户的计划任务：`每日企微社群任务导出`。

默认每天 09:40 执行后台导出。电脑关机或任务失败时，下次运行会自动补导未成功日期。

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

## 命令行

```powershell
scrm-exporter.exe run
scrm-exporter.exe login
scrm-exporter.exe status
scrm-exporter.exe install-task
scrm-exporter.exe uninstall-task
```

可以指定目录：

```powershell
scrm-exporter.exe run --config-dir "%LOCALAPPDATA%\ScrmDailyExporter" --data-dir "%USERPROFILE%\Documents\每日企微社群任务导出"
```

## 注意事项

- 登录态只保存在当前用户电脑上，不随安装包分发。
- 工具不会绕过账号权限，能导出什么取决于当前扫码账号权限。
- 如果 SCRM 后台接口或登录机制变化，需要更新工具版本。
