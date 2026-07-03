#define MyAppName "SCRM Daily Exporter"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "SCRM Automation"
#define MyAppExeName "scrm-exporter-ui.exe"

[Setup]
AppId={{8A86A55A-D7AF-43C4-9A59-873CE92E33C9}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\ScrmDailyExporter\app
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer-output
OutputBaseFilename=ScrmDailyExporterSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional tasks:"; Flags: unchecked

[Files]
Source: "dist\ScrmDailyExporter\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\scrm-exporter.exe"; Parameters: "install-task"; StatusMsg: "正在注册每日自动导出计划任务..."; Flags: runhidden waituntilterminated
Filename: "{app}\{#MyAppExeName}"; Description: "启动 {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{app}\scrm-exporter.exe"; Parameters: "uninstall-task"; Flags: runhidden waituntilterminated
