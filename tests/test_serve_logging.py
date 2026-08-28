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
