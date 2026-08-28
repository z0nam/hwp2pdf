"""The windowless build has no stdout, so every message must survive that."""

import sys

import pytest

from hwp2pdf import paths
from hwp2pdf.serve import DEFAULT_LOG_MAX_BYTES, ServerLog, build_parser, resolve_log_path


def test_writes_to_a_file_with_timestamps(tmp_path):
    log = ServerLog(tmp_path / "server.log")
    log("first")
    log("second")

    lines = (tmp_path / "server.log").read_text(encoding="utf-8").strip().splitlines()
    assert len(lines) == 2
    assert lines[0].endswith("  first")
    assert lines[1].endswith("  second")


def test_survives_a_windowless_build(tmp_path, monkeypatch):
    """sys.stdout is None in a PyInstaller windowed exe; print() would raise."""
    monkeypatch.setattr(sys, "stdout", None)
    log = ServerLog(tmp_path / "server.log")
    assert log.echo is False
    log("still recorded")
    assert "still recorded" in (tmp_path / "server.log").read_text(encoding="utf-8")


def test_rotates_when_the_file_grows(tmp_path):
    target = tmp_path / "server.log"
    log = ServerLog(target, max_bytes=200)
    for index in range(60):
        log(f"line {index}")

    assert target.exists()
    assert target.with_suffix(".log.1").exists()
    assert target.stat().st_size <= 200 + 200


def test_an_unwritable_path_never_raises(tmp_path):
    blocked = tmp_path / "afile"
    blocked.write_text("not a directory", encoding="utf-8")
    log = ServerLog(blocked / "nested" / "server.log")
    log("must not raise")  # falls back to console-only


def test_no_log_file_when_a_console_is_present(monkeypatch):
    monkeypatch.setattr(sys, "stdout", sys.__stdout__)
    assert resolve_log_path("") is None


def test_windowless_defaults_to_the_app_data_dir(monkeypatch):
    monkeypatch.setattr(sys, "stdout", None)
    assert resolve_log_path("") == paths.app_data_dir() / "server.log"


def test_explicit_log_file_wins(tmp_path, monkeypatch):
    monkeypatch.setattr(sys, "stdout", None)
    assert resolve_log_path(str(tmp_path / "custom.log")) == tmp_path / "custom.log"


def test_log_file_is_a_documented_option():
    args = build_parser().parse_args(["--log-file", "x.log"])
    assert args.log_file == "x.log"
    assert build_parser().parse_args([]).log_file == ""


@pytest.mark.parametrize("size", [DEFAULT_LOG_MAX_BYTES])
def test_default_rotation_size_is_sane(size):
    assert 256 * 1024 <= size <= 16 * 1024 * 1024


# -- --bind tailscale must survive a slow Tailscale at logon --------------
def test_bind_passes_through_a_literal_address():
    from hwp2pdf.serve import resolve_bind

    assert resolve_bind("0.0.0.0") == "0.0.0.0"
    assert resolve_bind("127.0.0.1") == "127.0.0.1"


def test_tailscale_bind_resolves_immediately_when_up(monkeypatch):
    from hwp2pdf import serve

    monkeypatch.setattr(serve, "tailscale_address", lambda: "100.64.0.1")
    assert serve.resolve_bind("tailscale") == "100.64.0.1"


def test_tailscale_bind_waits_for_a_late_start(monkeypatch):
    """At logon the server can beat Tailscale to the punch."""
    from hwp2pdf import serve

    attempts = {"n": 0}

    def late():
        attempts["n"] += 1
        return "100.64.0.9" if attempts["n"] >= 3 else None

    notes = []
    monkeypatch.setattr(serve, "tailscale_address", late)
    monkeypatch.setattr(serve, "TAILSCALE_POLL_SECONDS", 0)

    assert serve.resolve_bind("tailscale", wait_seconds=5, notify=notes.append) == "100.64.0.9"
    assert any("waiting" in n for n in notes)
    assert any("100.64.0.9" in n for n in notes)


def test_tailscale_bind_gives_up_eventually(monkeypatch):
    from hwp2pdf import serve

    monkeypatch.setattr(serve, "tailscale_address", lambda: None)
    monkeypatch.setattr(serve, "TAILSCALE_POLL_SECONDS", 0)

    with pytest.raises(SystemExit) as excinfo:
        serve.resolve_bind("tailscale", wait_seconds=0.05)
    assert "Tailscale" in str(excinfo.value)


def test_wait_is_configurable():
    from hwp2pdf.serve import TAILSCALE_WAIT_SECONDS, build_parser

    assert build_parser().parse_args([]).tailscale_wait == TAILSCALE_WAIT_SECONDS
    assert build_parser().parse_args(["--tailscale-wait", "30"]).tailscale_wait == 30
