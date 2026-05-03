# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import sys

ROOT = Path(SPECPATH)
ICON = ROOT / "assets" / "hwp_to_pdf_final.ico"
PYTHON_ROOT = Path(sys.base_prefix)
PYTHON_DLLS = PYTHON_ROOT / "DLLs"


a = Analysis(
    ["src/hwp2pdf/__main__.py"],
    pathex=[str(ROOT / "src")],
    binaries=[
        (str(PYTHON_DLLS / "_tkinter.pyd"), "."),
        (str(PYTHON_DLLS / "tcl86t.dll"), "."),
        (str(PYTHON_DLLS / "tk86t.dll"), "."),
    ],
    datas=[
        (str(PYTHON_ROOT / "Lib" / "tkinter"), "tkinter"),
        (str(PYTHON_ROOT / "tcl" / "tcl8.6"), "_tcl_data"),
        (str(PYTHON_ROOT / "tcl" / "tk8.6"), "_tk_data"),
    ],
    hiddenimports=["_tkinter", "pythoncom", "pywintypes", "win32com", "win32com.client"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[str(ROOT / "scripts" / "pyi_rth_tkinter_paths.py")],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    exclude_binaries=False,
    name="hwp2pdf",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ICON) if ICON.exists() else None,
)
