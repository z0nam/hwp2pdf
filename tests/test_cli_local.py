"""``--local`` / ``$HWP2PDF_LOCAL``: the CLI as a strictly local engine."""

from hwp2pdf import config
from hwp2pdf.cli import build_parser, resolve_backend_settings


def _settings(monkeypatch, argv, saved_url="http://saved:17650"):
    monkeypatch.setattr(config, "server_settings", lambda: {"url": saved_url, "token": "t",
                                                            "transport": "auto"})
    return resolve_backend_settings(build_parser().parse_args(argv))


def test_saved_server_is_used_by_default(monkeypatch):
    monkeypatch.delenv("HWP2PDF_LOCAL", raising=False)
    assert _settings(monkeypatch, ["a.hwp"])["url"] == "http://saved:17650"


def test_local_flag_clears_saved_and_explicit_server(monkeypatch):
    monkeypatch.delenv("HWP2PDF_LOCAL", raising=False)
    assert _settings(monkeypatch, ["a.hwp", "--local"])["url"] == ""
    assert _settings(monkeypatch, ["a.hwp", "--local", "--server", "host"])["url"] == ""


def test_local_env(monkeypatch):
    monkeypatch.setenv("HWP2PDF_LOCAL", "1")
    assert _settings(monkeypatch, ["a.hwp"])["url"] == ""
    monkeypatch.setenv("HWP2PDF_LOCAL", "0")
    assert _settings(monkeypatch, ["a.hwp"])["url"] == "http://saved:17650"
