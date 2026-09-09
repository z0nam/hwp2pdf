"""Finder "Open with" must reach the GUI, not the headless converter."""

import os
import stat

import pytest

pytest.importorskip("tkinter")
pytest.importorskip("tkinterdnd2")

from hwp2pdf import __main__ as entry  # noqa: E402


def test_documents_are_recognised(tmp_path):
    a = tmp_path / "a.hwp"
    b = tmp_path / "b.hwpx"
    a.write_bytes(b"x")
    b.write_bytes(b"x")
    assert entry.looks_like_documents([str(a), str(b)]) is True


def test_cli_invocations_are_not_mistaken_for_documents(tmp_path):
    folder = tmp_path / "docs"
    folder.mkdir()
    doc = tmp_path / "a.hwp"
    doc.write_bytes(b"x")

    assert entry.looks_like_documents([]) is False
    assert entry.looks_like_documents([str(folder)]) is False          # a folder
    assert entry.looks_like_documents([str(doc), "--pdf"]) is False    # a flag
    assert entry.looks_like_documents([str(tmp_path / "gone.hwp")]) is False
    assert entry.looks_like_documents(["serve"]) is False


def test_opening_documents_starts_the_gui_with_them_selected(tmp_path, monkeypatch):
    doc = tmp_path / "a.hwp"
    doc.write_bytes(b"x")
    seen = {}
    monkeypatch.setattr(entry, "gui_main", lambda initial_paths=(): seen.update(paths=list(initial_paths)))
    monkeypatch.setattr(entry, "cli_main", lambda argv: pytest.fail("CLI must not run"))

    assert entry.main([str(doc)]) == 0
    assert seen["paths"] == [str(doc)]


def test_flags_still_go_to_the_cli(tmp_path, monkeypatch):
    monkeypatch.setattr(entry, "gui_main", lambda **kw: pytest.fail("GUI must not run"))
    monkeypatch.setattr(entry, "cli_main", lambda argv: 7)
    assert entry.main([str(tmp_path), "--pdf"]) == 7


def test_no_arguments_opens_a_plain_gui(monkeypatch):
    seen = {}
    monkeypatch.setattr(entry, "gui_main", lambda **kw: seen.update(kw))
    assert entry.main([]) == 0
    assert seen == {}


# -- the Tcl console Tk builds when launched from Finder -------------------

def test_a_zero_block_character_device_is_replaced(monkeypatch):
    """launchd hands a bundled app /dev/null on fd 0, which is what makes Tk
    build a Tcl console window -- and building that window's menu bar aborts
    intermittently before any of this app runs."""
    monkeypatch.setattr(entry.sys, "platform", "darwin")
    monkeypatch.setattr(entry.sys, "frozen", True, raising=False)
    monkeypatch.setattr(entry.os, "isatty", lambda _fd: False)
    monkeypatch.setattr(entry.os, "fstat", lambda _fd: os.stat_result(
        (stat.S_IFCHR | 0o666, 0, 0, 1, 0, 0, 0, 0, 0, 0), {"st_blocks": 0}
    ))
    replaced = []
    monkeypatch.setattr(entry.os, "pipe", lambda: (7, 8))
    monkeypatch.setattr(entry.os, "dup2", lambda src, dst: replaced.append((src, dst)))
    monkeypatch.setattr(entry.os, "close", lambda _fd: None)

    entry._keep_tk_from_building_a_console()
    assert replaced == [(7, 0)], "stdin was left as the device Tk consoles for"


def test_a_real_stdin_is_left_alone(monkeypatch):
    # A pipe or a file on fd 0 is someone's actual input, not launchd's stub.
    monkeypatch.setattr(entry.sys, "platform", "darwin")
    monkeypatch.setattr(entry.sys, "frozen", True, raising=False)
    monkeypatch.setattr(entry.os, "isatty", lambda _fd: False)
    monkeypatch.setattr(entry.os, "fstat", lambda _fd: os.stat_result(
        (stat.S_IFIFO | 0o666, 0, 0, 1, 0, 0, 0, 0, 0, 0), {"st_blocks": 0}
    ))
    monkeypatch.setattr(entry.os, "pipe", lambda: pytest.fail("replaced a real stdin"))

    entry._keep_tk_from_building_a_console()


def test_a_terminal_keeps_its_stdin(monkeypatch):
    monkeypatch.setattr(entry.sys, "platform", "darwin")
    monkeypatch.setattr(entry.sys, "frozen", True, raising=False)
    monkeypatch.setattr(entry.os, "isatty", lambda _fd: True)
    monkeypatch.setattr(entry.os, "pipe", lambda: pytest.fail("replaced a tty"))

    entry._keep_tk_from_building_a_console()


def test_a_dev_run_is_left_alone(monkeypatch):
    # Only the frozen bundle has the problem; a test runner's stdin is its own.
    monkeypatch.setattr(entry.sys, "platform", "darwin")
    monkeypatch.delattr(entry.sys, "frozen", raising=False)
    monkeypatch.setattr(entry.os, "pipe", lambda: pytest.fail("touched a dev run"))

    entry._keep_tk_from_building_a_console()


@pytest.mark.parametrize("platform", ["win32", "linux"])
def test_only_macos_needs_this(monkeypatch, platform):
    monkeypatch.setattr(entry.sys, "platform", platform)
    monkeypatch.setattr(entry.sys, "frozen", True, raising=False)
    monkeypatch.setattr(entry.os, "pipe", lambda: pytest.fail(f"touched {platform}"))

    entry._keep_tk_from_building_a_console()
