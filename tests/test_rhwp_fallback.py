"""Local rhwp fallback: engaged only when the server cannot start, never silent."""

import os
import shutil
import subprocess
from pathlib import Path

import pytest

from fakes import FakeBackend, RecordingSink

from hwp2pdf import jobs
from hwp2pdf.backends import create_backend
from hwp2pdf.backends.base import BackendUnavailable
from hwp2pdf.backends.fallback import FallbackBackend
from hwp2pdf.backends.local_rhwp import ACTUAL_FORMAT, RHWP_ENV_VAR, RhwpBackend, find_rhwp

FIXTURE = Path(__file__).resolve().parent.parent / "docs" / "fixtures" / "docx-failure-repro.hwp"
REAL_RHWP = find_rhwp()


# -- discovery -----------------------------------------------------------
def test_explicit_path_wins(tmp_path):
    binary = tmp_path / "rhwp"
    binary.write_text("#!/bin/sh\n")
    assert find_rhwp(str(binary)) == binary


def test_missing_explicit_path_is_not_silently_replaced(tmp_path):
    assert find_rhwp(str(tmp_path / "nope")) is None


def test_env_var_is_honoured(tmp_path, monkeypatch):
    binary = tmp_path / "rhwp"
    binary.write_text("#!/bin/sh\n")
    monkeypatch.setenv(RHWP_ENV_VAR, str(binary))
    assert find_rhwp() == binary


def test_preflight_without_a_binary_is_a_clear_refusal(monkeypatch):
    monkeypatch.setattr("hwp2pdf.backends.local_rhwp.find_rhwp", lambda explicit="": None)
    with pytest.raises(BackendUnavailable) as excinfo:
        RhwpBackend().preflight("ko")
    assert RHWP_ENV_VAR in str(excinfo.value)


# -- format limits -------------------------------------------------------
def test_docx_is_refused_with_a_reason(tmp_path):
    backend = RhwpBackend()
    reason = backend.blocked_reason(tmp_path / "a.hwp", "DOCX", "ko")
    assert reason and "PDF" in reason


def test_pdf_is_accepted(tmp_path):
    assert RhwpBackend().blocked_reason(tmp_path / "a.hwp", "PDF", "ko") is None


# -- fallback wiring -----------------------------------------------------
def test_fallback_only_engages_when_the_primary_cannot_start():
    primary = FakeBackend()
    fallback = FakeBackend()
    wrapper = FallbackBackend(primary, fallback)
    wrapper.preflight("ko")
    assert wrapper.active is primary


def test_fallback_engages_and_says_so():
    primary = FakeBackend(unavailable="서버에 연결하지 못했습니다")
    fallback = FakeBackend()
    wrapper = FallbackBackend(primary, fallback)
    wrapper.preflight("ko")
    assert wrapper.active is fallback

    sink = RecordingSink()
    wrapper.open_session(sink, "ko", None)
    assert any("서버에 연결하지 못했습니다" in text for text in sink.logs())


def test_both_unavailable_reports_the_primary_problem():
    wrapper = FallbackBackend(
        FakeBackend(unavailable="서버 문제"), FakeBackend(unavailable="rhwp 없음")
    )
    with pytest.raises(BackendUnavailable) as excinfo:
        wrapper.preflight("ko")
    assert "서버 문제" in str(excinfo.value)


def test_capabilities_follow_the_active_backend():
    primary = FakeBackend(unavailable="down")
    fallback = FakeBackend()
    wrapper = FallbackBackend(primary, fallback)
    wrapper.preflight("ko")
    assert wrapper.capabilities.name == fallback.capabilities.name


def test_create_backend_wraps_only_when_asked():
    server = {"url": "http://host:17650", "token": "t", "transport": "auto", "shares": []}
    assert not isinstance(create_backend(server, "ko"), FallbackBackend)
    assert isinstance(create_backend(server, "ko", rhwp_fallback=True), FallbackBackend)


def test_no_server_configured_falls_back_to_whatever_is_local():
    """On Windows the local COM engine is still the primary; elsewhere rhwp is all there is."""
    backend = create_backend({"url": ""}, "ko", rhwp_fallback=True)
    if os.name == "nt":
        assert isinstance(backend, FallbackBackend)
        assert isinstance(backend.fallback, RhwpBackend)
    else:
        assert isinstance(backend, RhwpBackend)


def test_no_server_and_no_fallback_is_refused_off_windows():
    if os.name == "nt":
        pytest.skip("Windows has a local engine")
    with pytest.raises(BackendUnavailable):
        create_backend({"url": ""}, "ko")


# -- the real binary -----------------------------------------------------
@pytest.mark.skipif(REAL_RHWP is None, reason="rhwp not installed")
def test_real_rhwp_renders_a_pdf_and_is_labelled(tmp_path):
    """End to end with the actual binary, so the label cannot drift from reality."""
    source = tmp_path / "report.hwp"
    shutil.copy2(FIXTURE, source)

    sink = RecordingSink()
    jobs.run_batch(
        sink,
        RhwpBackend(),
        target=str(tmp_path), recursive=False, overwrite=True, use_safe_copy=False,
        force_one_page=True, output_formats=("PDF",), lang="ko",
    )

    pdf = tmp_path / "report.pdf"
    assert pdf.exists()
    assert pdf.read_bytes().startswith(b"%PDF-")
    # The engine is named in the log, so an approximate PDF stays identifiable.
    assert any(ACTUAL_FORMAT in text for text in sink.logs())
    assert any("한컴 엔진이 아니라" in text for text in sink.logs())


@pytest.mark.skipif(REAL_RHWP is None, reason="rhwp not installed")
def test_real_rhwp_refuses_docx(tmp_path):
    source = tmp_path / "report.hwp"
    shutil.copy2(FIXTURE, source)

    sink = RecordingSink()
    jobs.run_batch(
        sink,
        RhwpBackend(),
        target=str(tmp_path), recursive=False, overwrite=True, use_safe_copy=False,
        force_one_page=True, output_formats=("DOCX",), lang="ko",
    )
    assert not (tmp_path / "report.docx").exists()
    assert sink.done()[:3] == (0, 1, 0)
    assert any("PDF" in text for text in sink.logs())


def test_a_failing_rhwp_is_reported_not_raised(tmp_path, monkeypatch):
    backend = RhwpBackend()
    backend.binary = Path("/bin/false")

    from hwp2pdf.backends.base import JobSpec

    src = tmp_path / "a.hwp"
    src.write_bytes(b"x")
    result = backend.convert(JobSpec(
        index=1, src_path=src, open_path=src, save_path=tmp_path / "a.pdf",
        output_format="PDF", force_one_page=True, safe_temp=False, lang="ko",
    ))
    assert result.ok is False
    assert result.message


def test_a_hanging_rhwp_times_out(tmp_path, monkeypatch):
    backend = RhwpBackend(timeout=0.1)
    backend.binary = Path("/bin/sleep")

    def fake_run(*a, **kw):
        raise subprocess.TimeoutExpired(cmd="rhwp", timeout=0.1)

    monkeypatch.setattr(subprocess, "run", fake_run)
    from hwp2pdf.backends.base import JobSpec

    src = tmp_path / "a.hwp"
    src.write_bytes(b"x")
    result = backend.convert(JobSpec(
        index=1, src_path=src, open_path=src, save_path=tmp_path / "a.pdf",
        output_format="PDF", force_one_page=True, safe_temp=False, lang="ko",
    ))
    assert result.ok is False
