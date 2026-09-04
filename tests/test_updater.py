"""Choosing an update to install, and deciding whether we may install it.

The failure these guard against is silent and total: handing a Mac a build for
the wrong architecture, or a Windows archive, produces a download that simply
never runs.
"""

import sys

import pytest

from hwp2pdf import updater

RELEASE = {
    "tag_name": "v2026.09.02.3",
    "assets": [
        {"name": "hwp2pdf-2026.09.02.3.exe", "browser_download_url": "u/app.exe"},
        {"name": "hwp2pdf-cli-2026.09.02.3.exe", "browser_download_url": "u/cli.exe"},
        {"name": "hwp2pdf-macos-arm64-2026.09.02.3.zip", "browser_download_url": "u/arm64.zip"},
        {"name": "hwp2pdf-macos-x86_64-2026.09.02.3.zip", "browser_download_url": "u/x86_64.zip"},
        {"name": "hwp2pdf-serve-2026.09.02.3.exe", "browser_download_url": "u/serve.exe"},
        {"name": "hwp2pdf-setup-2026.09.02.3.exe", "browser_download_url": "u/setup.exe"},
        {"name": "hwp2pdf-windows-2026.09.02.3.zip", "browser_download_url": "u/windows.zip"},
        {"name": "hwp2pdf-linux-x86_64-2026.09.02.3.tar.gz", "browser_download_url": "u/linux-x86_64.tar.gz"},
    ],
}


def on_mac(monkeypatch, machine="arm64"):
    monkeypatch.setattr(updater.sys, "platform", "darwin")
    monkeypatch.setattr(updater.platform, "machine", lambda: machine)


def on_linux(monkeypatch, machine="x86_64"):
    monkeypatch.setattr(updater.sys, "platform", "linux")
    monkeypatch.setattr(updater.platform, "machine", lambda: machine)


@pytest.mark.parametrize("machine, expected", [
    ("arm64", "u/arm64.zip"),
    ("aarch64", "u/arm64.zip"),
    ("x86_64", "u/x86_64.zip"),
    ("amd64", "u/x86_64.zip"),
])
def test_each_mac_gets_the_build_for_its_own_architecture(monkeypatch, machine, expected):
    on_mac(monkeypatch, machine)
    assert updater.latest_release_download_url(RELEASE) == expected


def test_apple_silicon_can_fall_back_to_the_intel_build(monkeypatch):
    # Rosetta runs it. The reverse is impossible, which the next test pins.
    on_mac(monkeypatch, "arm64")
    intel_only = {"assets": [a for a in RELEASE["assets"] if "arm64" not in a["name"]]}
    assert updater.latest_release_download_url(intel_only) == "u/x86_64.zip"


def test_an_intel_mac_is_never_handed_the_arm64_build(monkeypatch):
    on_mac(monkeypatch, "x86_64")
    arm_only = {"assets": [a for a in RELEASE["assets"] if "x86_64" not in a["name"]]}
    assert updater.latest_release_download_url(arm_only) == ""


def test_a_mac_is_never_handed_a_windows_asset(monkeypatch):
    on_mac(monkeypatch)
    windows_only = {"assets": [a for a in RELEASE["assets"] if "macos" not in a["name"]]}
    assert updater.latest_release_download_url(windows_only) == ""


def test_an_unrecognised_mac_takes_whatever_macos_build_exists(monkeypatch):
    on_mac(monkeypatch, "riscv64")
    assert updater.latest_release_download_url(RELEASE) in ("u/arm64.zip", "u/x86_64.zip")


def test_linux_gets_the_tarball_for_its_own_architecture(monkeypatch):
    on_linux(monkeypatch, "x86_64")
    assert updater.latest_release_download_url(RELEASE) == "u/linux-x86_64.tar.gz"


def test_linux_spells_arm_the_way_uname_does(monkeypatch):
    # build_linux.sh names the tarball after `uname -m`, which reports
    # "aarch64" on Linux -- not the "arm64" macOS uses for the same silicon.
    on_linux(monkeypatch, "aarch64")
    arm = {"assets": RELEASE["assets"] + [{
        "name": "hwp2pdf-linux-aarch64-2026.09.02.3.tar.gz",
        "browser_download_url": "u/linux-aarch64.tar.gz",
    }]}
    assert updater.latest_release_download_url(arm) == "u/linux-aarch64.tar.gz"


def test_an_arm_linux_box_is_never_handed_the_x86_tarball(monkeypatch):
    on_linux(monkeypatch, "aarch64")
    assert updater.latest_release_download_url(RELEASE) == ""


def test_linux_is_never_handed_a_windows_or_macos_asset(monkeypatch):
    on_linux(monkeypatch)
    others = {"assets": [a for a in RELEASE["assets"] if "linux" not in a["name"]]}
    assert updater.latest_release_download_url(others) == ""


def test_linux_installs_nothing_by_itself(monkeypatch):
    # A tarball has no install-location convention, so the Linux build only
    # ever offers the download; the auto-update button stays hidden.
    monkeypatch.setattr(updater.sys, "platform", "linux")
    assert updater.is_updatable_asset_url("https://x/hwp2pdf-linux-x86_64-1.tar.gz") is False
    assert updater.can_auto_update() is False


def test_windows_still_prefers_the_installer(monkeypatch):
    monkeypatch.setattr(updater.sys, "platform", "win32")
    assert updater.latest_release_download_url(RELEASE) == "u/setup.exe"


def test_no_assets_at_all_is_not_a_crash(monkeypatch):
    on_mac(monkeypatch)
    assert updater.latest_release_download_url({"assets": []}) == ""
    assert updater.latest_release_download_url({}) == ""


# -- what the app may install by itself -----------------------------------

@pytest.mark.parametrize("url, expected", [
    ("https://x/hwp2pdf-macos-arm64-2026.09.02.3.zip", True),
    ("https://x/hwp2pdf-macos-x86_64-2026.09.02.3.zip", True),
    ("https://x/hwp2pdf-windows-2026.09.02.3.zip", False),
    ("https://x/hwp2pdf-setup-2026.09.02.3.exe", False),
    ("", False),
])
def test_macos_installs_only_a_macos_bundle(monkeypatch, url, expected):
    monkeypatch.setattr(updater.sys, "platform", "darwin")
    assert updater.is_updatable_asset_url(url) is expected


@pytest.mark.parametrize("url, expected", [
    ("https://x/hwp2pdf-setup-2026.09.02.3.exe", True),
    ("https://x/hwp2pdf-windows-2026.09.02.3.zip", False),
    ("https://x/hwp2pdf-macos-arm64-2026.09.02.3.zip", False),
])
def test_windows_installs_only_the_setup_exe(monkeypatch, url, expected):
    monkeypatch.setattr(updater.sys, "platform", "win32")
    assert updater.is_updatable_asset_url(url) is expected


# -- finding the bundle to replace ----------------------------------------

def test_the_bundle_is_found_from_the_running_executable(monkeypatch, tmp_path):
    bundle = tmp_path / "hwp2pdf.app"
    executable = bundle / "Contents" / "MacOS" / "hwp2pdf"
    executable.parent.mkdir(parents=True)
    executable.touch()

    monkeypatch.setattr(updater.sys, "platform", "darwin")
    monkeypatch.setattr(updater.sys, "executable", str(executable))
    monkeypatch.setattr(updater.sys, "_MEIPASS", str(bundle), raising=False)
    assert updater.app_bundle_path() == bundle.resolve()


def test_a_dev_run_has_no_bundle_to_swap(monkeypatch):
    monkeypatch.setattr(updater.sys, "platform", "darwin")
    monkeypatch.delattr(updater.sys, "_MEIPASS", raising=False)
    assert updater.app_bundle_path() is None
    assert updater.can_auto_update() is False


def test_a_binary_outside_a_bundle_has_none_either(monkeypatch, tmp_path):
    binary = tmp_path / "hwp2pdf-cli"
    binary.touch()
    monkeypatch.setattr(updater.sys, "platform", "darwin")
    monkeypatch.setattr(updater.sys, "executable", str(binary))
    monkeypatch.setattr(updater.sys, "_MEIPASS", str(tmp_path), raising=False)
    assert updater.app_bundle_path() is None


@pytest.mark.skipif(sys.platform == "win32", reason="POSIX permission bits")
def test_a_bundle_the_user_cannot_replace_is_not_auto_updatable(monkeypatch, tmp_path):
    holder = tmp_path / "Applications"
    bundle = holder / "hwp2pdf.app"
    bundle.mkdir(parents=True)
    monkeypatch.setattr(updater, "app_bundle_path", lambda: bundle)
    monkeypatch.setattr(updater.sys, "platform", "darwin")

    assert updater.can_auto_update() is True
    holder.chmod(0o500)          # readable and traversable, but not writable
    try:
        assert updater.can_auto_update() is False
    finally:
        holder.chmod(0o700)


# -- urgent releases -------------------------------------------------------

@pytest.mark.parametrize("body", [
    "hwp2pdf-priority: critical",
    "Security fix.\n\nhwp2pdf-priority: critical\n\nDetails below.",
    "HWP2PDF-PRIORITY: CRITICAL",          # the marker is not case-sensitive
])
def test_a_release_can_declare_itself_urgent(body):
    assert updater.release_is_critical({"body": body}) is True


@pytest.mark.parametrize("body", ["", None, "Just an ordinary release.", "critical"])
def test_an_ordinary_release_is_not_urgent(body):
    assert updater.release_is_critical({"body": body}) is False


def test_a_release_with_no_body_at_all_is_not_urgent():
    assert updater.release_is_critical({}) is False


def test_an_outstanding_urgent_update_is_chased_harder(monkeypatch):
    now = 1_000_000.0
    monkeypatch.setattr(updater.time, "time", lambda: now)
    two_hours_ago = {"checked_at": now - 2 * 3600}

    assert updater.should_check_updates(dict(two_hours_ago)) is False
    assert updater.should_check_updates(dict(two_hours_ago, priority="critical")) is True


def test_the_ordinary_cadence_still_waits(monkeypatch):
    now = 1_000_000.0
    monkeypatch.setattr(updater.time, "time", lambda: now)
    assert updater.should_check_updates({"checked_at": now - 5 * 3600}) is False
    assert updater.should_check_updates({"checked_at": now - 7 * 3600}) is True


def test_a_state_with_no_timestamp_checks_immediately():
    assert updater.should_check_updates({}) is True
    assert updater.should_check_updates({"checked_at": "nonsense"}) is True
