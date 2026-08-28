"""The COM backend must stay importable off Windows so the GUI can load."""

import os
from pathlib import Path

import pytest

from hwp2pdf.backends.base import BackendUnavailable
from hwp2pdf.backends.windows_com import (
    WindowsComBackend,
    blocked_conversion_reason,
    ensure_pywin32,
    get_hwp_processes,
    read_hwp_file_flags,
)

FIXTURE = Path(__file__).resolve().parent.parent / "docs" / "fixtures" / "docx-failure-repro.hwp"


def test_backend_capabilities():
    caps = WindowsComBackend.capabilities
    assert caps.name == "windows_com"
    assert caps.remote is False
    assert caps.local_staging is True
    assert caps.manages_hwp_process is True
    assert caps.local_preflight is True


@pytest.mark.skipif(os.name == "nt", reason="non-Windows degradation")
def test_preflight_raises_backend_unavailable_off_windows():
    with pytest.raises(BackendUnavailable):
        WindowsComBackend().preflight("ko")


@pytest.mark.skipif(os.name == "nt", reason="non-Windows degradation")
def test_helpers_degrade_quietly_off_windows():
    assert ensure_pywin32()[0] is False
    assert get_hwp_processes() == []
    assert read_hwp_file_flags(FIXTURE) is None
    assert blocked_conversion_reason(FIXTURE, "PDF", "ko") is None


@pytest.mark.skipif(os.name != "nt", reason="Windows only")
def test_file_header_flags_are_readable_on_windows():
    assert FIXTURE.exists()
    assert read_hwp_file_flags(FIXTURE) is not None


def test_console_output_is_decoded_with_the_oem_code_page(monkeypatch):
    """Regression: PYTHONUTF8=1 on Korean Windows broke process detection.

    tasklist writes its localized "no matching task" message in the console OEM
    code page. Under UTF-8 mode a bare text=True decoded it as UTF-8, crashed
    subprocess's reader thread and returned stdout=None.
    """
    from hwp2pdf.backends import windows_com

    seen = {}

    class Result:
        returncode = 0
        stdout = '"Hwp.exe","1234","Console","1","61,884 K"\n'

    def fake_run(args, **kwargs):
        seen.update(kwargs)
        return Result()

    monkeypatch.setattr(windows_com.subprocess, "run", fake_run)
    monkeypatch.setattr(windows_com, "IS_WINDOWS", True)

    assert windows_com.get_hwp_processes() == [{"name": "Hwp.exe", "pid": "1234"}]
    assert seen["encoding"] == "oem"
    assert seen["errors"] == "replace"


def test_missing_stdout_does_not_raise(monkeypatch):
    from hwp2pdf.backends import windows_com

    class Result:
        returncode = 0
        stdout = None

    monkeypatch.setattr(windows_com.subprocess, "run", lambda args, **kw: Result())
    assert windows_com.get_hwp_processes() == []
