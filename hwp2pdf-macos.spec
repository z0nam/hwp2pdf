# -*- mode: python ; coding: utf-8 -*-
#
# macOS build. Kept separate from hwp2pdf.spec because the Windows spec bundles
# tcl/tk by hand from sys.base_prefix/DLLs, ships the vendored security DLLs and
# produces two console/windowed exes -- none of which applies here. PyInstaller
# handles macOS tkinter on its own, so there is no runtime hook either.

import sys
from pathlib import Path

ROOT = Path(SPECPATH)
ICON = ROOT / "assets" / "hwp_to_pdf_final.icns"

sys.path.insert(0, str(ROOT / "src"))
from hwp2pdf.version import __version__  # noqa: E402

LAZY_IMPORTS = [
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
    binaries=[],
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
    [],
    exclude_binaries=True,
    name="hwp2pdf",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=True,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    gui_exe,
    a_gui.binaries,
    a_gui.datas,
    strip=False,
    upx=False,
    name="hwp2pdf",
)

app = BUNDLE(
    coll,
    name="hwp2pdf.app",
    icon=str(ICON) if ICON.exists() else None,
    bundle_identifier="io.github.z0nam.hwp2pdf",
    version=__version__,
    info_plist={
        "CFBundleShortVersionString": __version__,
        "CFBundleVersion": __version__,
        "LSMinimumSystemVersion": "11.0",
        "NSHighResolutionCapable": True,
        "NSHumanReadableCopyright": "MIT License",
        "CFBundleDocumentTypes": [
            {
                "CFBundleTypeName": "Hancom Office Hangul Document",
                "CFBundleTypeRole": "Viewer",
                "LSHandlerRank": "Alternate",
                "CFBundleTypeExtensions": ["hwp", "hwpx"],
            }
        ],
    },
)

a_cli = Analysis(
    ["src/hwp2pdf/cli.py"],
    pathex=[str(ROOT / "src")],
    binaries=[],
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
