"""Tests for updater asset matching across platforms."""

import sys
from hwp2pdf.updater import latest_release_download_url

RELEASE = {
    "assets": [
        {"name": "hwp2pdf-setup-2026.09.02.3.exe", "browser_download_url": "https://example.com/setup.exe"},
        {"name": "hwp2pdf-windows-2026.09.02.3.zip", "browser_download_url": "https://example.com/windows.zip"},
        {"name": "hwp2pdf-macos-arm64-2026.09.02.3.zip", "browser_download_url": "https://example.com/macos-arm64.zip"},
        {"name": "hwp2pdf-macos-x86_64-2026.09.02.3.zip", "browser_download_url": "https://example.com/macos-x86_64.zip"},
        {"name": "hwp2pdf-linux-x86_64-2026.09.02.3.tar.gz", "browser_download_url": "https://example.com/linux.tar.gz"},
    ]
}


def test_pick_release_asset_linux(monkeypatch):
    monkeypatch.setattr(sys, "platform", "linux")
    assert latest_release_download_url(RELEASE) == "https://example.com/linux.tar.gz"


def test_pick_release_asset_macos(monkeypatch):
    monkeypatch.setattr(sys, "platform", "darwin")
    assert latest_release_download_url(RELEASE).startswith("https://example.com/macos")


def test_pick_release_asset_windows(monkeypatch):
    monkeypatch.setattr(sys, "platform", "win32")
    assert latest_release_download_url(RELEASE) == "https://example.com/setup.exe"
