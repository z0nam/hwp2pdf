"""Persisted user settings (``settings.json``).

The app previously persisted nothing except the update-check cache. The remote
conversion server needs an address and a token to survive restarts, so this
module owns a small versioned JSON document next to the update state.

Reads never raise: a missing or corrupt file yields the defaults, mirroring how
``updater.load_update_state`` already behaves.
"""

import copy
import json
import os
from pathlib import Path

from hwp2pdf import paths

SCHEMA_VERSION = 1

TRANSPORT_AUTO = "auto"
TRANSPORT_UPLOAD = "upload"
TRANSPORT_SHARE = "share"
TRANSPORTS = (TRANSPORT_AUTO, TRANSPORT_UPLOAD, TRANSPORT_SHARE)

ENV_SERVER_URL = "HWP2PDF_SERVER_URL"
ENV_SERVER_TOKEN = "HWP2PDF_TOKEN"

DEFAULTS = {
    "version": SCHEMA_VERSION,
    "language": "ko",
    "server": {
        "url": "",
        "token": "",
        "transport": TRANSPORT_AUTO,
        # [{"name": "work", "local_mount": "/Volumes/work"}]
        "shares": [],
        "store_token_in_keychain": False,
    },
    "options": {
        "recursive": True,
        "overwrite": True,
        "safe_temp": True,
        "force_one_page": True,
        "formats": ["PDF"],
        "job_timeout_enabled": False,
        # Conservative opt-in recovery limit for a single local conversion.
        # Existing settings are preserved by _merge when loading.
        "job_timeout_minutes": 10,
        # Approximate local rendering when the preferred Hancom engine cannot start.
        "rhwp_fallback": False,
    },
    "last_target": "",
    #: Path to the rhwp executable; empty means "discover it".
    "rhwp_path": "",
}


def default_settings() -> dict:
    return copy.deepcopy(DEFAULTS)


def _merge(defaults, loaded):
    """Recursively fill missing keys, ignoring unknown ones."""
    result = copy.deepcopy(defaults)
    if not isinstance(loaded, dict):
        return result
    for key, default_value in defaults.items():
        if key not in loaded:
            continue
        value = loaded[key]
        if isinstance(default_value, dict):
            result[key] = _merge(default_value, value)
        elif isinstance(default_value, list):
            result[key] = list(value) if isinstance(value, list) else default_value
        elif isinstance(default_value, bool):
            result[key] = bool(value) if isinstance(value, bool) else default_value
        elif isinstance(default_value, str):
            result[key] = value if isinstance(value, str) else default_value
        else:
            result[key] = value
    return result


def load(path: Path | None = None) -> dict:
    """Load settings, falling back to defaults for anything missing or broken."""
    target = Path(path) if path else paths.settings_path()
    try:
        raw = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return default_settings()
    settings = _merge(DEFAULTS, raw)
    if settings["server"]["transport"] not in TRANSPORTS:
        settings["server"]["transport"] = TRANSPORT_AUTO
    settings["version"] = SCHEMA_VERSION
    return settings


def save(settings: dict, path: Path | None = None) -> bool:
    """Atomically write settings. Returns False instead of raising on failure."""
    target = Path(path) if path else paths.settings_path()
    payload = _merge(DEFAULTS, settings)
    payload["version"] = SCHEMA_VERSION
    tmp = target.with_name(target.name + ".tmp")
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        if os.name != "nt":
            os.chmod(tmp, 0o600)
        os.replace(tmp, target)
        return True
    except OSError:
        try:
            tmp.unlink()
        except OSError:
            pass
        return False


def server_settings(settings: dict | None = None) -> dict:
    """Server config with ``HWP2PDF_SERVER_URL`` / ``HWP2PDF_TOKEN`` applied.

    Environment variables win over the file so CI, the smoke script and the CLI
    can point at a server without touching the user's saved settings.
    """
    data = copy.deepcopy((settings or load())["server"])
    env_url = os.environ.get(ENV_SERVER_URL)
    env_token = os.environ.get(ENV_SERVER_TOKEN)
    if env_url:
        data["url"] = env_url
    if env_token:
        data["token"] = env_token
    return data
