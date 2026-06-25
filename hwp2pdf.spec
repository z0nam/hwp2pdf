# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import sys

ROOT = Path(SPECPATH)
ICON = ROOT / "assets" / "hwp_to_pdf_final.ico"
PYTHON_ROOT = Path(sys.base_prefix)
PYTHON_DLLS = PYTHON_ROOT / "DLLs"


SECURITY_DLL_X86 = ROOT / "vendor" / "x86" / "FilePathCheckerModule.dll"
SECURITY_DLL_X64 = ROOT / "vendor" / "x64" / "FilePathCheckerModule.dll"
SECURITY_DLL_DATAS = []
if SECURITY_DLL_X86.exists():
    SECURITY_DLL_DATAS.append((str(SECURITY_DLL_X86), "vendor/x86"))
if SECURITY_DLL_X64.exists():
    SECURITY_DLL_DATAS.append((str(SECURITY_DLL_X64), "vendor/x64"))


a_gui = Analysis(
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
        *SECURITY_DLL_DATAS,
    ],
    hiddenimports=[
        "_tkinter", "pythoncom", "pywintypes",
        "win32com", "win32com.client",
        "win32gui", "win32con", "win32process",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[str(ROOT / "scripts" / "pyi_rth_tkinter_paths.py")],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz_gui = PYZ(a_gui.pure)

gui_exe = EXE(
    pyz_gui,
    a_gui.scripts,
    a_gui.binaries,
    a_gui.datas,
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

a_cli = Analysis(
    ["src/hwp2pdf/cli.py"],
    pathex=[str(ROOT / "src")],
    binaries=[],
    datas=SECURITY_DLL_DATAS,
    hiddenimports=[
        "pythoncom", "pywintypes",
        "win32com", "win32com.client",
        "win32gui", "win32con", "win32process",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz_cli = PYZ(a_cli.pure)

cli_exe = EXE(
    pyz_cli,
    a_cli.scripts,
    a_cli.binaries,
    a_cli.datas,
    [],
    exclude_binaries=False,
    name="hwp2pdf-cli",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
