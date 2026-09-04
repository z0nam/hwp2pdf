# -*- mode: python ; coding: utf-8 -*-
#
# Linux build. Kept separate from hwp2pdf.spec and hwp2pdf-macos.spec because
# Windows bundles tcl/tk by hand and macOS produces a .app bundle. On Linux,
# PyInstaller bundles the shared libraries and produces standalone ELF binaries.

import sys
from pathlib import Path

ROOT = Path(SPECPATH)
ICON = ROOT / "assets" / "hwp_to_pdf_final.ico"
RHWP_BINARY = ROOT / "vendor" / "rhwp" / "rhwp"
RHWP_BINARIES = [(str(RHWP_BINARY), "vendor/rhwp")] if RHWP_BINARY.exists() else []

sys.path.insert(0, str(ROOT / "src"))
from hwp2pdf.version import __version__  # noqa: E402

LAZY_IMPORTS = [
    # Imported inside a function so the CA bundle travels with the build; see
    # hwp2pdf/certs.py for why a frozen app cannot use the platform's path.
    "certifi",
    # Reached through create_backend() and the `serve` subcommand.
    "hwp2pdf.backends.remote_http",
    "hwp2pdf.backends.windows_com",
    "hwp2pdf.serve",
    "hwp2pdf.server.http_server",
    "hwp2pdf.server.jobs",
]

a_gui = Analysis(
    ["src/hwp2pdf/__main__.py"],
    pathex=[str(ROOT / "src")],
    binaries=RHWP_BINARIES,
    datas=[],
    hiddenimports=["_tkinter", "tkinterdnd2", *LAZY_IMPORTS],
    hookspath=[str(ROOT)],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pythoncom", "pywintypes", "win32com", "win32gui", "win32con", "win32process"],
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
    upx=False,
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
    binaries=RHWP_BINARIES,
    datas=[],
    hiddenimports=LAZY_IMPORTS,
    hookspath=[str(ROOT)],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pythoncom", "pywintypes", "win32com", "win32gui", "win32con", "win32process"],
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
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
