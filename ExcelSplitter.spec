# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec file for Excel Splitter.

Build with (from the project root, on Windows, inside your venv):
    pyinstaller ExcelSplitter.spec

...or just run build.bat, which does a clean build using this file.

This produces a single ExcelSplitter.exe (--onefile-equivalent) with no
console window (--windowed-equivalent), per the project spec. If you'd
rather have an app that starts faster at the cost of being a folder full
of files instead of one .exe, change `onefile=True` below to `onefile=False`
-- see the note near EXE/COLLECT.
"""

from PyInstaller.utils.hooks import collect_submodules

block_cipher = None
onefile = True

# PyInstaller's static import analysis can miss some of pandas' and
# openpyxl's more dynamic submodules; collecting them explicitly avoids a
# "ModuleNotFoundError" that only shows up once the .exe is actually run on
# a clean machine (this exact failure mode is why the app is also run from
# source and smoke-tested before packaging -- see README.md).
hidden_imports = (
    collect_submodules("pandas")
    + collect_submodules("openpyxl")
    + collect_submodules("xlsxwriter")
    + [
        "xlrd",
        "PySide6.QtCore",
        "PySide6.QtGui",
        "PySide6.QtWidgets",
    ]
)

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "tkinter", "matplotlib", "PyQt5", "PyQt6", "PySide2",
        "IPython", "notebook", "pytest",
    ],
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe_kwargs = dict(
    name="ExcelSplitter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,  # --windowed: no command-prompt window on launch
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="resources/icons/app_icon.ico",
)

if onefile:
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.zipfiles,
        a.datas,
        [],
        runtime_tmpdir=None,
        **exe_kwargs,
    )
else:
    # --onedir alternative: faster startup, ships as ExcelSplitter/ folder
    # instead of a single file. Swap `onefile = False` above to use this.
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        **exe_kwargs,
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.zipfiles,
        a.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name="ExcelSplitter",
    )
