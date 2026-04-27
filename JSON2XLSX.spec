# -*- mode: python ; coding: utf-8 -*-
# Сборка: pyinstaller JSON2XLSX.spec
# Результат: один файл dist/JSON2XLSX.exe (onefile; при старте распаковка во временную папку).

from pathlib import Path

spec_dir = Path(SPECPATH)

a = Analysis(
    ["ConvertorJ2X.py"],
    pathex=[str(spec_dir)],
    binaries=[],
    datas=[(str(spec_dir / "schemas"), "schemas")],
    hiddenimports=["design", "taskmarks_aggregation"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="JSON2XLSX",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
