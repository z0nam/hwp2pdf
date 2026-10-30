"""First-run usability-data consent: write "decline" only where Hangul has no options file yet."""

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


# jaix-win 이 한컴 2022 에서 창을 막는 것을 확인한 파일 그대로(UTF-8 BOM, 209바이트, 2026-10-30)
MEASURED = b"\xef\xbb\xbf" + b"\r\n".join([
    b'<?xml version="1.0" encoding="utf-8"?>',
    b"<root>",
    b"  <Common>",
    b"    <OptionInfo>",
    b'      <EnableUsabilityLog value="0" />',
    b'      <EnableUsabilityLogSending value="0" />',
    b"    </OptionInfo>",
    b"  </Common>",
    b"</root>",
]) + b"\r\n"


def test_fresh_account_gets_the_measured_minimal_file(tmp_path):
    path = tmp_path / "Common" / "120" / "optionInfo.xml"
    assert decline_usability_consent(path) == "created"
    assert path.read_bytes() == MEASURED and len(MEASURED) == 209
    assert _values(path) == {"EnableUsabilityLog": "0", "EnableUsabilityLogSending": "0"}
    assert not path.with_name("optionInfo.xml.tmp").exists()


def test_existing_answer_is_never_changed(tmp_path):
    path = tmp_path / "optionInfo.xml"
    path.write_text(AGREED, encoding="utf-8-sig")
    before = path.read_bytes()
    assert decline_usability_consent(path) == "kept"
    assert path.read_bytes() == before


def test_existing_file_without_an_answer_is_left_alone(tmp_path):
    # 방침(사용자 결정 2026-10-30): 파일이 있으면 무엇이 들었든 건드리지 않는다
    path = tmp_path / "optionInfo.xml"
    for body in (AGREED.replace('      <EnableUsabilityLog value="1" />\n', "")
                 .replace('      <EnableUsabilityLogSending value="1" />\n', "").encode("utf-8-sig"),
                 b"not xml", b""):
        path.write_bytes(body)
        assert decline_usability_consent(path) == "kept"
        assert path.read_bytes() == body
    assert not list(tmp_path.glob("optionInfo.xml.*"))  # no .bak / .tmp


def test_failed_write_leaves_no_half_file(tmp_path, monkeypatch):
    import os

    path = tmp_path / "optionInfo.xml"

    def boom(*_a, **_k):
        raise OSError("locked")

    monkeypatch.setattr(os, "link", boom)
    try:
        decline_usability_consent(path)
    except OSError:
        pass
    assert not path.exists() and not (tmp_path / "optionInfo.xml.tmp").exists()


def test_start_failure_message_points_at_a_waiting_window():
    # 실측(2026-10-30): 동의창이 뜬 채 COM 이 "서버 실행에 실패했습니다"로 끝나면 이 메시지만 남는다
    from hwp2pdf.i18n import translate

    for lang, word in (("ko", "사용성 데이터 수집 동의"), ("en", "usability data")):
        msg = translate(lang, "local_hwp_start_failed", detail="(-2146959355, ...)")
        assert word in msg and msg.strip().splitlines()[-1].count(word) == 1


def test_never_replaces_a_file_hangul_made_meanwhile(tmp_path, monkeypatch):
    import os

    path = tmp_path / "optionInfo.xml"
    real_link = os.link

    def link_after_hangul(src, dst):
        Path(dst).write_bytes(b"hangul wrote this")
        return real_link(src, dst)

    monkeypatch.setattr(os, "link", link_after_hangul)
    assert decline_usability_consent(path) == "kept"
    assert path.read_bytes() == b"hangul wrote this"
    assert not (tmp_path / "optionInfo.xml.tmp").exists()


def test_only_the_measured_hancom_2022(monkeypatch, tmp_path):
    from hwp2pdf import hancom_options

    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setattr(hancom_options, "installed_version", lambda: "130")
    assert decline_usability_consent() == "skipped" and not any(tmp_path.rglob("optionInfo.xml"))
    monkeypatch.setattr(hancom_options, "installed_version", lambda: "120")
    assert decline_usability_consent() == "created"
    assert hancom_options.version_from_server(
        r"C:\Program Files (x86)\Hnc\Office 2022\HOffice120\bin\hwp.exe -Automation") == "120"
    assert hancom_options.version_from_server(r"C:\Hwp\hwp.exe") is None
