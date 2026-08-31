"""Platform-specific locations and file-manager integration.

Centralizes the paths that used to be hard-coded to Windows environment
variables in ``app.py``. On Windows every location resolves to exactly what the
previous code produced, so existing installs keep using the same files.
"""

import os
import subprocess
import sys
from pathlib import Path

IS_WINDOWS = os.name == "nt"
IS_MACOS = sys.platform == "darwin"

APP_DIR_NAME = "hwp2pdf"
SETTINGS_FILE_NAME = "settings.json"
UPDATE_STATE_FILE_NAME = "update_state.json"
SERVER_TOKEN_FILE_NAME = "server_token.txt"
WINDOWS_TEMP_WORKDIR = Path(r"C:\temp\hwp_convert_safe")


def app_data_dir() -> Path:
    """Per-user directory holding settings, update state and downloads."""
    if IS_WINDOWS:
        base = os.environ.get("LOCALAPPDATA") or Path.home()
        return Path(base) / APP_DIR_NAME
    if IS_MACOS:
        return Path.home() / "Library" / "Application Support" / APP_DIR_NAME
    base = os.environ.get("XDG_CONFIG_HOME") or (Path.home() / ".config")
    return Path(base) / APP_DIR_NAME


def settings_path() -> Path:
    return app_data_dir() / SETTINGS_FILE_NAME


def update_state_path() -> Path:
    return app_data_dir() / UPDATE_STATE_FILE_NAME


def update_download_dir() -> Path:
    return app_data_dir() / "updates"


def security_install_dir() -> Path:
    return app_data_dir() / "security"


def server_state_dir() -> Path:
    return app_data_dir() / "server"


def server_token_path() -> Path:
    return app_data_dir() / SERVER_TOKEN_FILE_NAME


def temp_workdir() -> Path:
    """Staging folder for the safe temporary conversion mode."""
    if IS_WINDOWS:
        return WINDOWS_TEMP_WORKDIR
    if IS_MACOS:
        return Path.home() / "Library" / "Caches" / APP_DIR_NAME / "convert"
    base = os.environ.get("XDG_CACHE_HOME") or (Path.home() / ".cache")
    return Path(base) / APP_DIR_NAME / "convert"


def resource_root() -> Path:
    """Root that bundled data files (``vendor/``) are resolved against."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base)
    return Path(__file__).resolve().parent.parent.parent


def reveal_in_file_manager(path) -> None:
    """Open ``path`` in Explorer / Finder / the desktop file manager."""
    target = str(path)
    if IS_WINDOWS:
        os.startfile(target)  # noqa: S606 - Windows only, same as before
        return
    opener = "open" if IS_MACOS else "xdg-open"
    subprocess.Popen([opener, target])
