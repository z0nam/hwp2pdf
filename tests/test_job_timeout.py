"""A stuck conversion must not end the batch or leak Hangul.

known-issues.md #1: docs/fixtures/docx-failure-repro.hwp still hangs Hangul for
13+ minutes on a Hancom dialog the watcher does not dismiss. These tests cover
the recovery path around that, without needing Hangul.
"""

import threading
import time

import pytest

from hwp2pdf.backends.base import JobSpec
from hwp2pdf.backends.windows_com import WindowsComBackend
from hwp2pdf.i18n import translate

from fakes import RecordingSink


def make_job(tmp_path, lang="ko"):
    src = tmp_path / "big.hwp"
    src.write_bytes(b"x")
    return JobSpec(
        index=1, src_path=src, open_path=src, save_path=tmp_path / "big.pdf",
        output_format="PDF", force_one_page=True, safe_temp=False, lang=lang,
    )


class StuckHwp:
    """Stands in for an HwpObject wedged behind a modal dialog."""

    def __init__(self, released):
        self.released = released
        self.cleared = False

    def Open(self, *_args):
        self.released.wait(5)
        raise OSError("RPC server unavailable")

    def Clear(self, _flag):
        self.cleared = True


def test_timeout_kills_hangul_and_reports_it(tmp_path, monkeypatch):
    released = threading.Event()
    killed = []
    monkeypatch.setattr(
        "hwp2pdf.backends.windows_com.kill_hwp",
        lambda: (killed.append(True), released.set()) and True,
    )

    backend = WindowsComBackend(job_timeout=0.2)
    backend._sink = RecordingSink()
    backend.hwp = StuckHwp(released)

    job = make_job(tmp_path)
    started = time.monotonic()
    result = backend.convert(job)
    elapsed = time.monotonic() - started

    assert killed, "the watchdog must force-close Hangul"
    assert elapsed < 4, "the batch must not stay blocked"
    assert result.ok is False
    assert result.message == translate("ko", "job_timeout", seconds=0)
    assert backend._engine_broken is True
    warned = [t for t in backend._sink.logs()]
    assert any("big.hwp" in text for text in warned)


def test_next_file_restarts_the_engine(tmp_path, monkeypatch):
    restarts = []

    def fake_start(sink, lang):
        restarts.append(lang)
        backend.hwp = object()

    backend = WindowsComBackend(job_timeout=None)
    backend._sink = RecordingSink()
    backend._engine_broken = True
    monkeypatch.setattr(backend, "_start_engine", fake_start)
    monkeypatch.setattr(backend, "_convert_document", lambda job: None, raising=False)

    assert backend._restart_engine("ko") is True
    assert restarts == ["ko"]
    assert any(translate("ko", "engine_restarted") in t for t in backend._sink.logs())


def test_a_failed_restart_is_reported_not_raised(tmp_path, monkeypatch):
    backend = WindowsComBackend()
    backend._sink = RecordingSink()
    backend._engine_broken = True

    def boom(sink, lang):
        raise OSError("no engine")

    monkeypatch.setattr(backend, "_start_engine", boom)
    result = backend.convert(make_job(tmp_path))

    assert result.ok is False
    assert backend._sink.logs()


def test_timeout_defaults_to_disabled_locally():
    assert WindowsComBackend().job_timeout is None


def test_server_enables_a_timeout_by_default():
    from hwp2pdf.serve import build_parser
    from hwp2pdf.server import protocol

    assert build_parser().parse_args([]).job_timeout == protocol.DEFAULT_JOB_TIMEOUT_SECONDS
    assert protocol.DEFAULT_JOB_TIMEOUT_SECONDS > 0
    assert build_parser().parse_args(["--job-timeout", "0"]).job_timeout == 0


def test_cli_timeout_flag_is_opt_in():
    from hwp2pdf.cli import build_parser

    assert build_parser().parse_args(["t"]).timeout == 0
    assert build_parser().parse_args(["t", "--timeout", "300"]).timeout == 300
