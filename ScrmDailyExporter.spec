# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules


hiddenimports = collect_submodules("docx") + collect_submodules("openpyxl")
common_datas = [(".env.example", "."), ("README_APP.md", ".")]

cli = Analysis(
    ["scrm_exporter.py"],
    pathex=[],
    binaries=[],
    datas=common_datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
cli_pyz = PYZ(cli.pure)
cli_exe = EXE(
    cli_pyz,
    cli.scripts,
    [],
    exclude_binaries=True,
    name="scrm-exporter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
)

ui = Analysis(
    ["scrm_exporter_ui.py"],
    pathex=[],
    binaries=[],
    datas=common_datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
ui_pyz = PYZ(ui.pure)
ui_exe = EXE(
    ui_pyz,
    ui.scripts,
    [],
    exclude_binaries=True,
    name="scrm-exporter-ui",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
)

coll = COLLECT(
    cli_exe,
    ui_exe,
    cli.binaries,
    ui.binaries,
    cli.datas,
    ui.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="ScrmDailyExporter",
)
