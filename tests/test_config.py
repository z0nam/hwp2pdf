import json
import os

import pytest

from hwp2pdf import config


@pytest.fixture
def settings_file(tmp_path):
    return tmp_path / "settings.json"


def test_missing_file_returns_defaults(settings_file):
    loaded = config.load(settings_file)
    assert loaded == config.default_settings()


def test_roundtrip(settings_file):
    settings = config.default_settings()
    settings["server"]["url"] = "http://host:8765"
    settings["server"]["token"] = "abc"
    settings["server"]["transport"] = config.TRANSPORT_SHARE
    settings["options"]["formats"] = ["PDF", "DOCX"]
    settings["language"] = "en"

    assert config.save(settings, settings_file) is True
    loaded = config.load(settings_file)

    assert loaded["server"]["url"] == "http://host:8765"
    assert loaded["server"]["transport"] == config.TRANSPORT_SHARE
    assert loaded["options"]["formats"] == ["PDF", "DOCX"]
    assert loaded["language"] == "en"


def test_corrupt_file_falls_back_to_defaults(settings_file):
    settings_file.write_text("{ not json", encoding="utf-8")
    assert config.load(settings_file) == config.default_settings()


def test_partial_file_is_filled_in(settings_file):
    settings_file.write_text(json.dumps({"server": {"url": "http://x"}}), encoding="utf-8")
    loaded = config.load(settings_file)
    assert loaded["server"]["url"] == "http://x"
    assert loaded["server"]["transport"] == config.TRANSPORT_AUTO
    assert loaded["options"]["recursive"] is True


def test_unknown_transport_is_reset(settings_file):
    settings_file.write_text(json.dumps({"server": {"transport": "carrier-pigeon"}}), encoding="utf-8")
    assert config.load(settings_file)["server"]["transport"] == config.TRANSPORT_AUTO


@pytest.mark.skipif(os.name == "nt", reason="POSIX permissions")
def test_settings_file_is_owner_only(settings_file):
    config.save(config.default_settings(), settings_file)
    assert settings_file.stat().st_mode & 0o777 == 0o600


def test_env_overrides_file(settings_file, monkeypatch):
    settings = config.default_settings()
    settings["server"]["url"] = "http://file:1"
    settings["server"]["token"] = "file-token"
    config.save(settings, settings_file)

    monkeypatch.setenv(config.ENV_SERVER_URL, "http://env:2")
    monkeypatch.setenv(config.ENV_SERVER_TOKEN, "env-token")
    resolved = config.server_settings(config.load(settings_file))

    assert resolved["url"] == "http://env:2"
    assert resolved["token"] == "env-token"


def test_env_absent_keeps_file_values(settings_file, monkeypatch):
    settings = config.default_settings()
    settings["server"]["url"] = "http://file:1"
    config.save(settings, settings_file)

    monkeypatch.delenv(config.ENV_SERVER_URL, raising=False)
    monkeypatch.delenv(config.ENV_SERVER_TOKEN, raising=False)

    assert config.server_settings(config.load(settings_file))["url"] == "http://file:1"
