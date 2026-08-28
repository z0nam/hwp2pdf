"""Finder "Open with" must reach the GUI, not the headless converter."""

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
