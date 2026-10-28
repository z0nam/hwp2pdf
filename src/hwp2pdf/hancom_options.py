"""Pre-answer Hancom Office 2022's first-run "usability data" consent as *declined*.

The first time Hangul starts under a Windows account it shows a "사용성 데이터 수집
동의" window. Driven through COM nobody sees it, so the first conversion on a fresh
account hangs until someone clicks a button. Hangul keeps the answer per user in
``%APPDATA%\\Hnc\\User\\Common\\120\\optionInfo.xml``; "동의 안 함" is stored as::

    <EnableUsabilityLog value="0" />
    <EnableUsabilityLogSending value="0" />

(measured on a test account that clicked "동의 안 함"; an account that agreed has
``1``/``1``). Before starting Hangul we write those two ``0`` values **only where no
answer exists yet** — never ``1``, and an existing answer is never changed.
"""

from __future__ import annotations

import io
import os
import shutil
import xml.etree.ElementTree as ET
from pathlib import Path

#: Hancom Office 2022 keeps per-user options under this version folder
OPTION_VERSION = "120"
USABILITY_KEYS = ("EnableUsabilityLog", "EnableUsabilityLogSending")
DECLINED = "0"

_MINIMAL = (
    '<?xml version="1.0" encoding="utf-8"?>\n'
    "<root>\n"
    "  <Common>\n"
    "    <OptionInfo>\n"
    + "".join(f'      <{key} value="{DECLINED}" />\n' for key in USABILITY_KEYS)
    + "    </OptionInfo>\n"
    "  </Common>\n"
    "</root>\n"
)


def option_info_path(version: str = OPTION_VERSION) -> Path | None:
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return None
    return Path(appdata) / "Hnc" / "User" / "Common" / version / "optionInfo.xml"


def decline_usability_consent(path: Path | None = None) -> str:
    """Record "동의 안 함" where Hangul has no answer yet. Returns what happened:

    ``created`` (no options file: wrote one holding just the two answers),
    ``added`` (file existed without them: added the missing ones as 0; the original
    is kept once as ``optionInfo.xml.bak``),
    ``kept`` (an answer is already there — left as the user chose),
    ``skipped`` (no %APPDATA%, or a file we cannot parse: left untouched).
    """
    if path is None:
        path = option_info_path()
    if path is None:
        return "skipped"
    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        # Hangul writes this file as UTF-8 with a BOM
        _atomic_write(path, _MINIMAL.encode("utf-8-sig"))
        return "created"
    try:
        tree = ET.parse(path)
    except (ET.ParseError, OSError):
        return "skipped"
    root = tree.getroot()
    common = root.find("Common")
    if common is None:
        common = ET.SubElement(root, "Common")
    options = common.find("OptionInfo")
    if options is None:
        options = ET.SubElement(common, "OptionInfo")
    missing = [key for key in USABILITY_KEYS if options.find(key) is None]
    if not missing:
        return "kept"
    for key in missing:
        ET.SubElement(options, key, value=DECLINED)
    ET.indent(tree, space="  ")
    buf = io.BytesIO()
    buf.write(b"\xef\xbb\xbf")
    tree.write(buf, encoding="utf-8", xml_declaration=True)
    # the file holds the user's other Hancom options too: keep the original once,
    # and swap the new one in atomically so a failed write cannot truncate it
    backup = path.with_name(path.name + ".bak")
    if not backup.exists():
        shutil.copy2(path, backup)
    _atomic_write(path, buf.getvalue())
    return "added"


def _atomic_write(path: Path, data: bytes) -> None:
    tmp = path.with_name(path.name + ".tmp")
    try:
        tmp.write_bytes(data)
        os.replace(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)
