"""The backend split must not change what ``hwp2pdf.app`` exposes.

``cli.py``, ``src/hwp_pdf_converter_app_safe.py`` and ``scripts/check_windows.ps1``
all import engine helpers from this module. A missing re-export only shows up at
runtime on Windows, so pin the whole surface here instead.
"""

import pytest

pytest.importorskip("tkinter")
pytest.importorskip("tkinterdnd2")

from hwp2pdf import app  # noqa: E402


@pytest.mark.parametrize("name", app.LEGACY_EXPORTS)
def test_legacy_export_resolves(name):
    assert getattr(app, name, None) is not None, name


def test_all_matches_the_legacy_surface():
    assert set(app.__all__) == set(app.LEGACY_EXPORTS)


def test_check_windows_ps1_import_still_works():
    # scripts/check_windows.ps1 runs exactly this.
    from hwp2pdf.app import output_extension

    assert output_extension("DOCX") == ".docx"
    assert output_extension("PDF") == ".pdf"


def test_compat_entrypoint_module_imports():
    import importlib
    import sys
    from pathlib import Path

    src = Path(app.__file__).resolve().parent.parent
    if str(src) not in sys.path:
        sys.path.insert(0, str(src))
    module = importlib.import_module("hwp_pdf_converter_app_safe")
    assert module.ConverterApp is app.ConverterApp
    assert module.main is app.main
