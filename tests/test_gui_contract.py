"""GUI wiring checks that do not need a human in front of the window."""

import pytest

pytest.importorskip("tkinter")
pytest.importorskip("tkinterdnd2")

import tkinter as tk  # noqa: E402
from tkinter import ttk  # noqa: E402
from tkinterdnd2 import TkinterDnD  # noqa: E402

import time  # noqa: E402

from hwp2pdf import config, discovery  # noqa: E402
from hwp2pdf.version import __version__  # noqa: E402
from hwp2pdf.server.protocol import DEFAULT_PORT  # noqa: E402
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


# -- multi-file selection alongside the server panel ---------------------
def test_file_selection_shows_the_list_and_counts(app, tmp_path):
    files = []
    for name in ("a.hwp", "b.hwpx"):
        path = tmp_path / name
        path.write_bytes(b"x")
        files.append(path)
    (tmp_path / "c.txt").write_bytes(b"x")

    assert app.ui["file_list_frame"].winfo_manager() != "grid"
    app._set_file_targets(files + [tmp_path / "c.txt"], append=False)

    assert [str(p) for p in app.selected_files] == [str(p) for p in files]
    assert app.file_listbox.size() == 2
    assert app.ui["file_list_frame"].winfo_manager() == "grid"
    assert "2" in app.file_count_var.get()


def test_dropping_more_files_appends_instead_of_replacing(app, tmp_path):
    first = tmp_path / "a.hwp"
    second = tmp_path / "b.hwp"
    first.write_bytes(b"x")
    second.write_bytes(b"x")

    app._handle_dropped_paths([str(first)])
    app._handle_dropped_paths([str(second)])

    assert [p.name for p in app.selected_files] == ["a.hwp", "b.hwp"]
    # The same file twice must not duplicate.
    app._handle_dropped_paths([str(second)])
    assert len(app.selected_files) == 2


def test_dropping_a_folder_replaces_the_file_selection(app, tmp_path):
    picked = tmp_path / "a.hwp"
    picked.write_bytes(b"x")
    folder = tmp_path / "docs"
    folder.mkdir()

    app._handle_dropped_paths([str(picked)])
    assert app.selected_files
    app._handle_dropped_paths([str(folder)])

    assert app.selected_files == []
    assert app.folder_var.get() == str(folder)


def test_clear_and_remove_selected(app, tmp_path):
    for name in ("a.hwp", "b.hwp", "c.hwp"):
        (tmp_path / name).write_bytes(b"x")
    app._set_file_targets([tmp_path / n for n in ("a.hwp", "b.hwp", "c.hwp")], append=False)

    app.file_listbox.selection_set(1)
    app.remove_selected_files()
    assert [p.name for p in app.selected_files] == ["a.hwp", "c.hwp"]

    app.clear_selected_files()
    assert app.selected_files == []
    assert app.folder_var.get() == ""
    assert app.ui["file_list_frame"].winfo_manager() != "grid"


def test_successful_file_target_is_removed_but_other_targets_remain(app, tmp_path):
    first = tmp_path / "a.hwp"
    second = tmp_path / "b.hwp"
    first.write_bytes(b"x")
    second.write_bytes(b"x")
    app._set_file_targets([first, second], append=False)

    app._remove_completed_file_target(first)

    assert app.selected_files == [second]
    assert app.folder_var.get() == str(second)
    assert app.file_listbox.size() == 1


def test_completion_no_longer_opens_a_blocking_dialog(app, monkeypatch):
    """The done handler writes to the log instead of a modal messagebox."""
    import tkinter.messagebox as messagebox

    opened = []
    monkeypatch.setattr(messagebox, "showinfo", lambda *a, **k: opened.append(a))

    app.is_running = True
    app.log_queue.put(("done", (2, 1, 0, "/tmp/hwp2pdf_log.csv", False)))
    app._poll_log_queue()

    assert opened == []
    log = app.log_text.get("1.0", "end")
    assert "1" in log and "2" in log
    assert app.is_running is False


def test_multi_file_selection_and_server_panel_coexist(app, tmp_path):
    """The merge must not have either feature clobber the other."""
    picked = tmp_path / "a.hwp"
    picked.write_bytes(b"x")
    app._set_file_targets([picked], append=False)

    # Windows defaults to the local engine, so opt into remote the way a user
    # would; elsewhere remote is the only option.
    app.use_remote_var.set(True)
    app._apply_backend_mode()
    app.server_url_var.set("http://host:8765")

    assert app.selected_files
    assert app.backend_settings()["url"] == "http://host:8765"
    assert app.ui["server_frame"].winfo_manager() == "pack"


def test_windows_defaults_to_the_local_engine(app):
    if not IS_WINDOWS:
        pytest.skip("Windows-only default")
    assert app.use_remote_var.get() is False
    assert app.backend_settings()["url"] == ""


# -- local conversion timeout (opt-in) -----------------------------------
def test_timeout_is_off_by_default(app):
    assert app.job_timeout_var.get() is False
    assert app.job_timeout_seconds() is None


def test_timeout_seconds_come_from_the_minutes_box(app):
    app.use_remote_var.set(False)
    if not IS_WINDOWS:
        # Remote-only platforms always defer to the server.
        assert app.job_timeout_seconds() is None
        return
    app.job_timeout_var.set(True)
    app.job_timeout_minutes_var.set("15")
    assert app.job_timeout_seconds() == 900


def test_remote_conversion_defers_to_the_server(app):
    app.use_remote_var.set(True)
    app._apply_backend_mode()
    app.job_timeout_var.set(True)
    app.job_timeout_minutes_var.set("5")
    assert app.job_timeout_seconds() is None
    assert str(app.ui["job_timeout_check"].cget("state")) == "disabled"
    assert app.ui["job_timeout_note"].cget("text")


def test_windows_engine_status_shows_running_only_when_hancom_is_installed(app, monkeypatch):
    monkeypatch.setattr("hwp2pdf.app.IS_WINDOWS", True)

    app._apply_engine_status({"installed": True, "running": ["123"]})
    assert app.hancom_install_status_var.get() == app.tr("hancom_install_ready")
    assert app.hwp_running_status_var.get() == app.tr("hwp_status_running")
    assert app.ui["hwp_running_status"].winfo_manager() == "grid"

    app._apply_engine_status({"installed": True, "running": []})
    assert app.hwp_running_status_var.get() == app.tr("hwp_status_not_running")

    app._apply_engine_status({"installed": False, "running": []})
    assert app.hancom_install_status_var.get() == app.tr("hancom_install_missing")
    assert app.ui["hwp_running_status"].winfo_manager() == ""


def test_macos_engine_status_hides_the_local_process_row(app, monkeypatch):
    monkeypatch.setattr("hwp2pdf.app.IS_WINDOWS", False)
    app._apply_engine_status({"installed": True, "running": ["123"]})

    assert app.hancom_install_status_var.get() == app.tr("hancom_install_unsupported")
    assert app.ui["hwp_running_status"].winfo_manager() == ""


def test_engine_status_probe_result_is_sent_to_the_gui_queue(app, monkeypatch):
    expected = {"installed": True, "detail": "registered", "running": ["321"]}
    monkeypatch.setattr("hwp2pdf.app.probe_hwp", lambda: expected)

    app._engine_status_worker()

    assert app.log_queue.get_nowait() == ("engine_status", expected)


def test_rhwp_option_explains_mode_and_availability(app, monkeypatch, tmp_path):
    binary = tmp_path / "rhwp.exe"
    binary.write_bytes(b"exe")
    monkeypatch.setattr("hwp2pdf.app.find_rhwp", lambda _path="": binary)

    if IS_WINDOWS:
        # use_remote_backend() is pinned to True off Windows -- there is no
        # local Hancom engine to fall back from -- so only Windows sees this.
        app.use_remote_var.set(False)
        app._refresh_rhwp_ui()
        assert app.rhwp_help_var.get() == app.tr("rhwp_fallback_help_local")

    app.use_remote_var.set(False)
    app._refresh_rhwp_ui()
    assert app.rhwp_status_var.get() == app.tr("rhwp_status_ready")
    assert str(app.ui["rhwp_fallback_check"].cget("state")) == "normal"

    app.use_remote_var.set(True)
    app._refresh_rhwp_ui()
    assert app.rhwp_help_var.get() == app.tr("rhwp_fallback_help_remote")


def test_missing_rhwp_disables_the_fallback_option(app, monkeypatch):
    monkeypatch.setattr("hwp2pdf.app.find_rhwp", lambda _path="": None)
    app.rhwp_fallback_var.set(True)
    app._refresh_rhwp_ui()

    assert app.rhwp_fallback_var.get() is False
    assert str(app.ui["rhwp_fallback_check"].cget("state")) == "disabled"
    assert app.tr("rhwp_status_missing") in app.rhwp_status_var.get()


def test_running_hwp_can_use_rhwp_for_this_run(app, monkeypatch, tmp_path):
    binary = tmp_path / "rhwp.exe"
    binary.write_bytes(b"exe")
    monkeypatch.setattr("hwp2pdf.app.get_hwp_processes", lambda: [{"pid": "123", "name": "Hwp.exe"}])
    monkeypatch.setattr("hwp2pdf.app.find_rhwp", lambda _path="": binary)
    monkeypatch.setattr(app, "_ask_hwp_running_action", lambda *a, **k: "rhwp")

    assert app._confirm_local_engine_ready(("PDF", "DOCX")) == "rhwp"


def test_one_run_rhwp_choice_omits_docx_without_recording_a_failure(app, monkeypatch, tmp_path):
    if not IS_WINDOWS:
        pytest.skip("local HWP process choice is Windows-only")
    source = tmp_path / "a.hwp"
    source.write_bytes(b"x")
    app._set_file_targets([source], append=False)
    app.output_pdf_var.set(True)
    app.output_docx_var.set(True)
    monkeypatch.setattr(app, "_confirm_local_engine_ready", lambda *a, **k: "rhwp")

    captured = {}

    class PendingThread:
        def __init__(self, *, target, args, daemon):
            captured.update(target=target, args=args, daemon=daemon)

        def start(self):
            captured["started"] = True

    monkeypatch.setattr("hwp2pdf.app.threading.Thread", PendingThread)
    app.start_conversion()

    assert captured["started"] is True
    assert captured["args"][5] == ("PDF",)
    assert captured["args"][9]["only"] is True


def test_a_nonsense_minutes_value_disables_rather_than_crashes(app):
    app.use_remote_var.set(False)
    app.job_timeout_var.set(True)
    for bad in ("", "abc", "0", "-5"):
        app.job_timeout_minutes_var.set(bad)
        assert app.job_timeout_seconds() in (None, 0) or app.job_timeout_seconds() > 0


def test_timeout_settings_persist(app, tmp_path):
    app.job_timeout_var.set(True)
    app.job_timeout_minutes_var.set("45")
    app._save_settings()

    saved = config.load(tmp_path / "settings.json")
    assert saved["options"]["job_timeout_enabled"] is True
    assert saved["options"]["job_timeout_minutes"] == 45


# -- reaching a server ---------------------------------------------------

def test_a_bare_host_name_becomes_a_full_address(app):
    app.use_remote_var.set(True)
    app.server_url_var.set("namun-ji")
    assert app.backend_settings()["url"] == f"http://namun-ji:{DEFAULT_PORT}"
    # The field itself is rewritten, so the user sees what will be used.
    assert app.server_url_var.get() == f"http://namun-ji:{DEFAULT_PORT}"


def test_pasting_an_invite_fills_in_the_token_too(app):
    app.use_remote_var.set(True)
    app.server_url_var.set(discovery.make_invite("namun-ji", "s3cr3t"))
    app._on_address_entered()

    assert app.server_url_var.get() == f"http://namun-ji:{DEFAULT_PORT}"
    assert app.server_token_var.get() == "s3cr3t"
    assert app.server_status_var.get() == app.tr("server_invite_applied")


def test_an_address_that_makes_no_sense_is_reported_and_left_alone(app):
    app.use_remote_var.set(True)
    app.server_url_var.set("ftp://nope")
    app._on_address_entered()

    # Left in place: the user has to see the typo to fix it.
    assert app.server_url_var.get() == "ftp://nope"
    assert app.server_status_var.get() == app.tr("server_address_invalid", value="ftp://nope")


class _NoThread:
    """Stands in for the worker thread so a test can drive it step by step."""

    def __init__(self, *_args, **_kwargs):
        pass

    def start(self):
        pass


@pytest.fixture
def find_steps(app, monkeypatch):
    """Run a search as three visible steps: the click, the worker, the UI pump."""
    monkeypatch.setattr("hwp2pdf.app.threading.Thread", _NoThread)

    def run(wide=False):
        app.find_servers(wide=wide)
        assert str(app.ui["server_find_btn"].cget("state")) == "disabled"
        app._find_servers_worker(wide)
        app._poll_log_queue()

    return run


def test_finding_nothing_offers_to_search_wider(app, monkeypatch, find_steps):
    app.use_remote_var.set(True)
    monkeypatch.setattr("hwp2pdf.discovery.discover", lambda **_kw: [])

    find_steps()

    assert app.server_find_running is False
    assert app.server_status_var.get() == app.tr("server_find_none")
    assert str(app.ui["server_find_btn"].cget("state")) == "normal"
    # The escalation appears only now -- never before a quiet first search.
    assert app.ui["server_wide_btn"].winfo_manager() == "grid"


def test_the_wider_search_is_not_what_the_first_press_does(app, monkeypatch, find_steps):
    app.use_remote_var.set(True)
    asked = []
    monkeypatch.setattr(
        "hwp2pdf.discovery.discover", lambda **kw: asked.append(kw.get("wide")) or []
    )

    assert app.ui["server_wide_btn"].winfo_manager() == ""
    find_steps()
    find_steps(wide=True)

    assert asked == [False, True]
    assert app.server_status_var.get() == app.tr(
        "server_find_wide_none", port=DEFAULT_PORT
    )
    # Nothing further to escalate to, so the button stands down again.
    assert app.ui["server_wide_btn"].winfo_manager() == ""


def test_a_wider_search_that_finds_something_hides_the_escalation(app, monkeypatch, find_steps):
    app.use_remote_var.set(True)
    monkeypatch.setattr("hwp2pdf.discovery.discover", lambda **_kw: [])
    find_steps()
    assert app.ui["server_wide_btn"].winfo_manager() == "grid"

    monkeypatch.setattr("hwp2pdf.discovery.discover", lambda **_kw: list(FOUND))
    monkeypatch.setattr(app, "_choose_server", lambda servers: None)
    find_steps(wide=True)

    assert app.ui["server_wide_btn"].winfo_manager() == ""


def test_discovery_failing_outright_is_not_fatal(app, monkeypatch):
    def explode(**_kw):
        raise OSError("no network")

    monkeypatch.setattr("hwp2pdf.discovery.discover", explode)
    app._find_servers_worker()
    assert app.log_queue.get_nowait() == ("server_find", (False, []))


FOUND = [
    {"url": "http://100.64.0.1:17650", "name": "namun-ji", "version": "2026.09.02.2",
     "api": 1, "auth_required": True, "compatible": True, "via": discovery.VIA_TAILSCALE},
    {"url": "http://192.168.0.5:17650", "name": "old-box", "version": "2025.01.01.1",
     "api": 0, "auth_required": False, "compatible": False, "via": discovery.VIA_LAN},
]


def _found_tree(dialog):
    return next(
        child for child in dialog.winfo_children() if isinstance(child, ttk.Treeview)
    )


def _dialog_button(dialog, label):
    for frame in dialog.winfo_children():
        for child in frame.winfo_children():
            if isinstance(child, ttk.Button) and child.cget("text") == label:
                return child
    raise AssertionError(f"no {label!r} button in the dialog")


def test_the_found_list_shows_why_each_server_may_not_work(app):
    dialog = app._choose_server(FOUND)
    try:
        rows = [_found_tree(dialog).item(item, "values")
                for item in _found_tree(dialog).get_children()]
        assert [row[1] for row in rows] == [server["url"] for server in FOUND]
        assert app.tr("server_find_via_tailscale") in rows[0][3]
        assert app.tr("server_find_needs_token") in rows[0][3]
        assert app.tr("server_find_incompatible") in rows[1][3]
    finally:
        dialog.destroy()


def test_picking_a_server_that_needs_a_token_stops_to_ask_for_one(app):
    app.use_remote_var.set(True)
    app.server_token_var.set("")
    dialog = app._choose_server(FOUND)
    _found_tree(dialog).selection_set("0")
    _dialog_button(dialog, app.tr("server_find_select")).invoke()

    assert app.server_url_var.get() == FOUND[0]["url"]
    # No connection test fires while the token box is still empty.
    assert app.server_test_running is False
    assert app.server_status_var.get() == app.tr("server_find_needs_token")


# -- being told about updates ---------------------------------------------

def test_clicking_the_version_line_asks_again_right_away(app, monkeypatch):
    monkeypatch.setattr("hwp2pdf.app.threading.Thread", _NoThread)
    # A check that just happened: the ordinary cadence would say "not yet".
    monkeypatch.setattr("hwp2pdf.app.load_update_state", lambda: {"checked_at": time.time()})
    monkeypatch.setattr("hwp2pdf.app.should_check_updates", lambda _state: False)

    app.check_for_updates_if_due()
    assert app.update_check_running is False

    app.ui["update_status_label"].event_generate("<Button-1>")
    assert app.update_check_running is True
    assert app.update_status_var.get() == app.tr("update_status_checking")


def test_an_urgent_release_says_so_on_the_status_line(app):
    ordinary = {"status": "newer", "latest": "9999.01.01.1", "priority": ""}
    app._apply_update_state(ordinary)
    assert app.update_status_var.get() == app.tr(
        "update_status_available", current=__version__, latest="9999.01.01.1"
    )

    app._apply_update_state(dict(ordinary, priority="critical"))
    assert app.update_status_var.get() == app.tr(
        "update_status_critical", current=__version__, latest="9999.01.01.1"
    )


def test_an_urgent_release_prompts_and_an_ordinary_one_does_not(app, monkeypatch):
    asked = []
    monkeypatch.setattr("hwp2pdf.app.messagebox.askyesno",
                        lambda *a, **k: asked.append(a[1]) or False)

    app._warn_if_update_is_urgent({"priority": "", "latest": "9999.01.01.1"})
    assert asked == []

    app._warn_if_update_is_urgent({"priority": "critical", "latest": "9999.01.01.1"})
    assert asked == [app.tr("update_critical_prompt", latest="9999.01.01.1")]


def test_an_urgent_release_we_already_have_does_not_nag(app, monkeypatch):
    asked = []
    monkeypatch.setattr("hwp2pdf.app.messagebox.askyesno", lambda *a, **k: asked.append(a) or False)
    app._warn_if_update_is_urgent({"priority": "critical", "latest": __version__})
    assert asked == []
