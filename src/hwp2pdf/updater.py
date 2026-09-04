"""GitHub release check and update-state cache.

Extracted verbatim from ``app.py``; the ``ConverterApp`` methods that drive the
download and relaunch still live there.
"""

import json
import os
import platform
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

from hwp2pdf import paths
from hwp2pdf.version import __version__

GITHUB_RELEASES_API_URL = "https://api.github.com/repos/z0nam/hwp2pdf/releases/latest"


GITHUB_RELEASES_PAGE_URL = "https://github.com/z0nam/hwp2pdf/releases/latest"


#: A release published in the afternoon used to stay invisible until the next
#: day. Six hours keeps the request count trivial and a same-day release
#: reachable; clicking the version line asks immediately regardless.
UPDATE_CHECK_INTERVAL_SECONDS = 6 * 60 * 60

#: While an urgent release is outstanding the app stops waiting politely.
CRITICAL_CHECK_INTERVAL_SECONDS = 60 * 60

#: A release says it is urgent by carrying this line in its notes. GitHub has
#: no severity field, and the marker reads sensibly to a human scrolling the
#: release page, so it lives in the text rather than in a side channel.
CRITICAL_MARKER = "hwp2pdf-priority: critical"


def parse_version(value: str):
    parts = []
    for part in value.strip().lstrip("vV").split("."):
        try:
            parts.append(int(part))
        except ValueError:
            break
    return tuple(parts)


def latest_release_version(release: dict):
    for key in ("tag_name", "name"):
        value = str(release.get(key) or "").strip()
        parsed = parse_version(value)
        if parsed:
            return value.lstrip("vV")
    return ""


def latest_release_download_url(release: dict):
    assets = release.get("assets") or []
    if not isinstance(assets, list):
        return ""

    candidates = []
    for asset in assets:
        if not isinstance(asset, dict):
            continue
        name = str(asset.get("name") or "").lower()
        url = str(asset.get("browser_download_url") or "").strip()
        if name and url:
            candidates.append((name, url))

    if sys.platform == "darwin":
        # Never fall through to a bare ".zip" here: the Windows archive would
        # match it. A Mac with no recognised architecture gets whatever macOS
        # build exists; a known one gets its own.
        arch = macos_arch()
        if arch == "arm64":
            # Apple silicon runs the Intel build under Rosetta if it must; an
            # Intel Mac cannot run the arm64 build at all, so it never falls back.
            preferred_patterns = (("macos", "arm64", ".zip"), ("macos", "x86_64", ".zip"))
        elif arch:
            preferred_patterns = (("macos", arch, ".zip"),)
        else:
            preferred_patterns = (("macos", ".zip"),)
    elif sys.platform.startswith("linux"):
        # Same rule: an archive built for another architecture -- or for
        # another platform entirely -- is worse than offering no download.
        arch = linux_arch()
        preferred_patterns = (("linux", arch, ".tar.gz"),) if arch else (("linux", ".tar.gz"),)
    else:
        preferred_patterns = (
            ("setup", ".exe"),
            ("windows", ".zip"),
            (".exe",),
            (".zip",),
        )
    for pattern in preferred_patterns:
        for name, url in candidates:
            if all(part in name for part in pattern):
                return url
    # Guessing means handing over a binary for the wrong architecture or
    # another platform entirely; better to send the user to the release page.
    # Windows keeps its historical last-resort behaviour.
    if sys.platform == "darwin" or sys.platform.startswith("linux"):
        return ""
    return candidates[0][1] if candidates else ""


def linux_arch() -> str:
    """This machine's architecture as the Linux release assets spell it.

    ``build_linux.sh`` names its tarball after ``uname -m``, so the value is
    taken verbatim: an arm64 Linux archive is "aarch64", never "arm64".
    """
    return platform.machine().lower()


def macos_arch() -> str:
    """This Mac's architecture as the macOS release assets spell it."""
    machine = platform.machine().lower()
    if machine in ("arm64", "aarch64"):
        return "arm64"
    if machine in ("x86_64", "amd64"):
        return "x86_64"
    return ""


def release_is_critical(release: dict) -> bool:
    """Whether the release asks to be installed now rather than eventually."""
    return CRITICAL_MARKER.lower() in str(release.get("body") or "").lower()


def fetch_latest_release():
    request = urllib.request.Request(
        GITHUB_RELEASES_API_URL,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": f"hwp2pdf/{__version__}",
        },
    )
    with urllib.request.urlopen(request, timeout=10) as response:
        return json.loads(response.read().decode("utf-8"))


def load_update_state():
    try:
        with UPDATE_STATE_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_update_state(state: dict):
    try:
        UPDATE_STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
        with UPDATE_STATE_PATH.open("w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def should_check_updates(state: dict):
    try:
        checked_at = float(state.get("checked_at", 0))
    except (TypeError, ValueError):
        checked_at = 0
    interval = (
        CRITICAL_CHECK_INTERVAL_SECONDS
        if state.get("priority") == "critical"
        else UPDATE_CHECK_INTERVAL_SECONDS
    )
    return time.time() - checked_at >= interval


def is_installed_build() -> bool:
    """True when this exe was placed by the Inno Setup installer (its
    unins000.exe / .dat marker sits next to the exe). PyInstaller portable
    builds and dev runs return False."""
    if getattr(sys, "_MEIPASS", None) is None:
        return False
    try:
        exe_dir = Path(sys.executable).resolve().parent
    except Exception:
        return False
    return (exe_dir / "unins000.exe").exists() or (exe_dir / "unins000.dat").exists()


def app_bundle_path():
    """The ``.app`` this build runs from, or None when it is not inside one.

    A dev run and the bare CLI binary both return None: there is no bundle to
    swap, so they are told to download instead.
    """
    if sys.platform != "darwin" or getattr(sys, "_MEIPASS", None) is None:
        return None
    try:
        executable = Path(sys.executable).resolve()
    except OSError:
        return None
    for parent in executable.parents:
        if parent.suffix == ".app":
            return parent
    return None


def is_updatable_asset_url(url: str) -> bool:
    """Whether this asset is one the app knows how to install by itself."""
    if not url:
        return False
    name = url.rsplit("/", 1)[-1].lower()
    if sys.platform == "darwin":
        return name.startswith("hwp2pdf-macos-") and name.endswith(".zip")
    return name.startswith("hwp2pdf-setup-") and name.endswith(".exe")


def can_auto_update() -> bool:
    """Whether this build can replace itself in place.

    Windows needs the Inno Setup installer that put it there. macOS needs a
    ``.app`` sitting in a directory this user can write, since the update is a
    bundle swap -- an app in /Applications on a machine where the user is not
    an admin has to go through the download page instead.
    """
    if sys.platform == "darwin":
        bundle = app_bundle_path()
        return bundle is not None and os.access(bundle.parent, os.W_OK)
    return is_installed_build()


UPDATE_STATE_PATH = paths.update_state_path()
UPDATE_DOWNLOAD_DIR = paths.update_download_dir()
