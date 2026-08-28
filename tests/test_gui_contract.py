"""GUI wiring checks that do not need a human in front of the window."""

import pytest

pytest.importorskip("tkinter")
pytest.importorskip("tkinterdnd2")

import tkinter as tk  # noqa: E402
from tkinterdnd2 import TkinterDnD  # noqa: E402

from hwp2pdf import config  # noqa: E402
from hwp2pdf.app import ConverterApp  # noqa: E402
from hwp2pdf.i18n import TEXT  # noqa: E402
from hwp2pdf.paths import IS_WINDOWS  # noqa: E402


@pytest.fixture
def app(tmp_path, monkeypatch):
    monkeypatch.setattr("hwp2pdf.paths.app_data_dir", lambda: tmp_path)
    try:
        root = TkinterDnD.Tk()
    except (tk.TclError, RuntimeError) as e:
        pytest.skip(f"no display: {e}")
    root.withdraw()
    instance = ConverterApp(root)
    root.update_idletasks()
    try:
        yield instance
    finally:
        root.destroy()


def test_server_panel_visibility_follows_the_backend_mode(app):
    assert app.use_remote_backend() is (not IS_WINDOWS or bool(app.server_url_var.get()))
    packed = app.ui["server_frame"].winfo_manager() == "pack"
    assert packed is bool(app.use_remote_backend())
    # The local/remote choice only exists where a local engine could run.
    assert bool(app.ui["server_remote_check"].winfo_manager()) is IS_WINDOWS

    if IS_WINDOWS:
        app.use_remote_var.set(False)
        app._apply_backend_mode()
        assert app.ui["server_frame"].winfo_manager() != "pack"
        app.use_remote_var.set(True)
        app._apply_backend_mode()
        assert app.ui["server_frame"].winfo_manager() == "pack"


def test_language_switch_relabels_the_server_panel(app):
    app.language_var.set("English")
    app._apply_language()
    assert app.ui["server_frame"].cget("text") == TEXT["en"]["server_section"]
    assert app.ui["server_test_btn"].cget("text") == TEXT["en"]["server_test"]
    assert app.ui["server_transport_combo"].cget("values") == (
        TEXT["en"]["transport_auto"], TEXT["en"]["transport_upload"], TEXT["en"]["transport_share"],
    )


def test_transport_combobox_maps_labels_back_to_codes(app):
    app.transport_label_var.set(app.tr("transport_share"))
    app._on_transport_changed()
    assert app.server_transport_var.get() == config.TRANSPORT_SHARE


def test_backend_settings_is_a_plain_dict_safe_to_hand_to_a_worker(app):
    app.use_remote_var.set(True)
    app.server_url_var.set("http://host:8765")
    app.server_token_var.set("t")
    settings = app.backend_settings()
    assert settings == {
        "url": "http://host:8765", "token": "t",
        "transport": app.server_transport_var.get(), "shares": [],
    }
    # Reading Tk variables off the main thread raises, so nothing in the value
    # handed to the conversion thread may be a Tk variable.
    assert all(not isinstance(v, tk.Variable) for v in settings.values())


def test_settings_round_trip_through_the_config_file(app, tmp_path):
    app.server_url_var.set("http://saved:8765")
    app.output_docx_var.set(True)
    app.language_var.set("English")
    app._save_settings()

    saved = config.load(tmp_path / "settings.json")
    assert saved["server"]["url"] == "http://saved:8765"
    assert saved["language"] == "en"
    assert saved["options"]["formats"] == ["PDF", "DOCX"]


def test_connection_test_without_an_address_reports_instead_of_hanging(app):
    app.server_url_var.set("")
    app.test_server_connection()
    assert app.server_test_running is False
    assert app.server_status_var.get()
