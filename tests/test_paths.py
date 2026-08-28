import os
import sys
from pathlib import Path

import pytest

from hwp2pdf import paths


def test_app_data_dir_is_under_home_or_localappdata():
    directory = paths.app_data_dir()
    assert directory.name == paths.APP_DIR_NAME
    assert directory.is_absolute()


def test_derived_paths_live_in_app_data_dir():
    root = paths.app_data_dir()
    assert paths.settings_path().parent == root
    assert paths.update_state_path().parent == root
    assert paths.update_download_dir().parent == root
    assert paths.security_install_dir().parent == root
    assert paths.server_token_path().parent == root


@pytest.mark.skipif(os.name != "nt", reason="Windows layout")
def test_windows_layout_matches_previous_hardcoded_paths():
    base = Path(os.environ.get("LOCALAPPDATA") or Path.home())
    assert paths.app_data_dir() == base / "hwp2pdf"
    assert paths.temp_workdir() == Path(r"C:\temp\hwp_convert_safe")


@pytest.mark.skipif(sys.platform != "darwin", reason="macOS layout")
def test_macos_layout():
    assert paths.app_data_dir() == Path.home() / "Library" / "Application Support" / "hwp2pdf"
    assert paths.temp_workdir() == Path.home() / "Library" / "Caches" / "hwp2pdf" / "convert"


def test_resource_root_contains_vendor_dir_when_running_from_source():
    if getattr(sys, "_MEIPASS", None):
        pytest.skip("frozen build")
    assert (paths.resource_root() / "src" / "hwp2pdf").is_dir()
