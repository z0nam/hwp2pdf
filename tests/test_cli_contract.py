"""The CLI keeps importing engine helpers from ``hwp2pdf.app`` after the split."""

import pytest

tk = pytest.importorskip("tkinter")
pytest.importorskip("tkinterdnd2")

from hwp2pdf import app, jobs  # noqa: E402
from hwp2pdf.cli import CliConversionContext, CliEventSink, build_parser, selected_formats  # noqa: E402


def test_app_still_exports_the_names_cli_imports():
    for name in ("APP_NAME", "ConverterApp", "enabled_extensions", "get_hwp_processes",
                 "kill_hwp", "translate"):
        assert hasattr(app, name), name


def test_converter_app_collect_files_is_the_shared_implementation(tmp_path):
    (tmp_path / "a.hwp").write_bytes(b"x")
    assert app.ConverterApp.collect_files(str(tmp_path), False) == jobs.collect_files(str(tmp_path), False)


def test_cli_context_duck_types_the_run_batch_caller(tmp_path):
    context = CliConversionContext()
    assert isinstance(context.log_queue, CliEventSink)
    assert context.stop_requested is False
    (tmp_path / "a.hwp").write_bytes(b"x")
    assert [p.name for p in context.collect_files(str(tmp_path), False)] == ["a.hwp"]


def test_parser_defaults_to_pdf():
    args = build_parser().parse_args(["target"])
    assert selected_formats(args) == ("PDF",)
    assert selected_formats(build_parser().parse_args(["t", "--docx"])) == ("DOCX",)
