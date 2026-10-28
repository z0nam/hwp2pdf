"""First-run usability-data consent: pre-answer "decline" only where no answer exists."""

from pathlib import Path

from hwp2pdf.hancom_options import decline_usability_consent, option_info_path

AGREED = """<?xml version="1.0" encoding="utf-8"?>
<root>
  <Common>
    <OptionInfo>
      <MakeDumpFile value="1" />
      <EnableUsabilityLog value="1" />
      <EnableUsabilityLogSending value="1" />
    </OptionInfo>
  </Common>
</root>
"""


def _values(path: Path) -> dict:
    import xml.etree.ElementTree as ET

    opts = ET.parse(path).getroot().find("Common/OptionInfo")
    return {el.tag: el.get("value") for el in opts}


def test_path_is_per_user_under_appdata(monkeypatch, tmp_path):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    assert option_info_path() == tmp_path / "Hnc" / "User" / "Common" / "120" / "optionInfo.xml"
    monkeypatch.delenv("APPDATA")
    assert option_info_path() is None and decline_usability_consent() == "skipped"


def test_fresh_account_gets_declined_answer(tmp_path):
    path = tmp_path / "Common" / "120" / "optionInfo.xml"
    assert decline_usability_consent(path) == "created"
    assert path.read_bytes().startswith(b"\xef\xbb\xbf")  # UTF-8 BOM, as Hangul writes it
    assert _values(path) == {"EnableUsabilityLog": "0", "EnableUsabilityLogSending": "0"}


def test_existing_answer_is_never_changed(tmp_path):
    path = tmp_path / "optionInfo.xml"
    path.write_text(AGREED, encoding="utf-8-sig")
    before = path.read_bytes()
    assert decline_usability_consent(path) == "kept"
    assert path.read_bytes() == before


def test_missing_answer_added_as_decline_keeping_other_options(tmp_path):
    path = tmp_path / "optionInfo.xml"
    path.write_text(AGREED.replace('      <EnableUsabilityLog value="1" />\n', "")
                    .replace('      <EnableUsabilityLogSending value="1" />\n', ""),
                    encoding="utf-8-sig")
    assert decline_usability_consent(path) == "added"
    assert _values(path) == {"MakeDumpFile": "1", "EnableUsabilityLog": "0",
                             "EnableUsabilityLogSending": "0"}
    assert path.read_bytes().startswith(b"\xef\xbb\xbf")
    assert b"EnableUsabilityLog" not in (tmp_path / "optionInfo.xml.bak").read_bytes()
    assert not (tmp_path / "optionInfo.xml.tmp").exists()


def test_failed_write_leaves_the_original_intact(tmp_path, monkeypatch):
    import os

    path = tmp_path / "optionInfo.xml"
    original = AGREED.replace('      <EnableUsabilityLog value="1" />\n', "")
    path.write_text(original, encoding="utf-8-sig")
    before = path.read_bytes()

    def boom(*_a, **_k):
        raise OSError("locked")

    monkeypatch.setattr(os, "replace", boom)
    try:
        decline_usability_consent(path)
    except OSError:
        pass
    assert path.read_bytes() == before
    assert not (tmp_path / "optionInfo.xml.tmp").exists()


def test_partial_answer_only_fills_the_gap(tmp_path):
    path = tmp_path / "optionInfo.xml"
    path.write_text(AGREED.replace('      <EnableUsabilityLogSending value="1" />\n', ""),
                    encoding="utf-8-sig")
    assert decline_usability_consent(path) == "added"
    assert _values(path)["EnableUsabilityLog"] == "1"  # the user's own answer stays
    assert _values(path)["EnableUsabilityLogSending"] == "0"


def test_unparsable_file_is_left_alone(tmp_path):
    path = tmp_path / "optionInfo.xml"
    path.write_bytes(b"not xml")
    assert decline_usability_consent(path) == "skipped"
    assert path.read_bytes() == b"not xml"
