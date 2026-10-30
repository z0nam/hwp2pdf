"""Pre-answer Hancom Office 2022's first-run "usability data" consent as *declined*.

The first time Hangul starts under a Windows account it shows a "사용성 데이터 수집
동의" window. Driven through COM nobody sees it, so the first conversion on a fresh
account hangs until someone clicks a button. Hangul keeps the answer per user in
``%APPDATA%\\Hnc\\User\\Common\\120\\optionInfo.xml``; "동의 안 함" is stored as::

    <EnableUsabilityLog value="0" />
    <EnableUsabilityLogSending value="0" />

(an account that agreed has ``1``/``1``). **Only when that file does not exist yet**
— Hangul has never run under this account — we write a minimal file holding just
those two ``0`` values (user decision 2026-10-30). An existing file is never
touched, whatever it holds: never ``1``, never an edit to a file Hangul wrote.

Measured on Hancom Office 2022 (2026-10-30): with :data:`MINIMAL` in place the
first conversion ran with no window in 8 s, and Hangul kept both values (it only
appended ``<Activation value="1" />``).
"""

from __future__ import annotations

import os
from pathlib import Path

#: Hancom Office 2022 keeps per-user options under this version folder
OPTION_VERSION = "120"

#: the file that was measured to suppress the window — written with a UTF-8 BOM and
#: CRLF line breaks, as Hangul writes it (209 bytes)
MINIMAL = "\r\n".join([
    '<?xml version="1.0" encoding="utf-8"?>',
    "<root>",
    "  <Common>",
    "    <OptionInfo>",
    '      <EnableUsabilityLog value="0" />',
    '      <EnableUsabilityLogSending value="0" />',
    "    </OptionInfo>",
    "  </Common>",
    "</root>",
]) + "\r\n"


def option_info_path(version: str = OPTION_VERSION) -> Path | None:
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return None
    return Path(appdata) / "Hnc" / "User" / "Common" / version / "optionInfo.xml"


def decline_usability_consent(path: Path | None = None) -> str:
    """Record "동의 안 함" for an account Hangul has never run under. Returns:

    ``created`` (no options file: wrote :data:`MINIMAL`),
    ``kept`` (the file exists — left exactly as it is),
    ``skipped`` (no %APPDATA%).
    """
    if path is None:
        path = option_info_path()
    if path is None:
        return "skipped"
    if path.exists():
        return "kept"
    path.parent.mkdir(parents=True, exist_ok=True)
    # written before Hangul starts, so nothing races us; tmp + replace keeps a failed
    # write from leaving a half file Hangul would then read
    tmp = path.with_name(path.name + ".tmp")
    try:
        tmp.write_bytes(MINIMAL.encode("utf-8-sig"))
        os.replace(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)
    return "created"
