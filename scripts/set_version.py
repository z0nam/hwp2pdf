"""Compute the next ``yyyy.MM.dd.N`` build version and write ``version.py``.

Platform-neutral replacement for the version block that used to live inside
``scripts/build_windows.ps1``. Both ``build_windows.ps1`` and
``scripts/build_macos.sh`` call this so the two platforms never disagree about
what a given build is called.

The build number is the highest sequence number already present in ``dist/``
or ``release/`` for today's date, plus one.

Usage:
    python scripts/set_version.py            # compute, write, print
    python scripts/set_version.py 2026.08.28.3   # pin an explicit version
    python scripts/set_version.py --print-only   # compute and print, do not write
"""

import argparse
import datetime
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VERSION_FILE = ROOT / "src" / "hwp2pdf" / "version.py"
SCAN_DIRS = ("dist", "release")
EXPLICIT_VERSION_PATTERN = re.compile(r"^\d{4}\.\d{2}\.\d{2}\.\d+$")


def existing_numbers(date_part: str):
    pattern = re.compile(
        r"^hwp2pdf(?:-windows|-macos(?:-[a-z0-9_]+)?)?-"
        + re.escape(date_part)
        + r"\.(\d+)(?:\.exe|\.zip|\.app)?$"
    )
    numbers = []
    for name in SCAN_DIRS:
        directory = ROOT / name
        if not directory.is_dir():
            continue
        for entry in directory.iterdir():
            match = pattern.match(entry.name)
            if match:
                numbers.append(int(match.group(1)))
    return numbers


def next_version(today=None) -> str:
    date_part = (today or datetime.date.today()).strftime("%Y.%m.%d")
    numbers = existing_numbers(date_part)
    return f"{date_part}.{max(numbers) + 1 if numbers else 1}"


def write_version(version: str) -> None:
    VERSION_FILE.write_text(f'__version__ = "{version}"\n', encoding="utf-8")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("version", nargs="?", help="Explicit yyyy.MM.dd.N version to pin")
    parser.add_argument("--print-only", action="store_true", help="Do not write version.py")
    args = parser.parse_args(argv)

    if args.version:
        if not EXPLICIT_VERSION_PATTERN.match(args.version):
            parser.exit(1, f"ERROR: not a yyyy.MM.dd.N version: {args.version}\n")
        version = args.version
    else:
        version = next_version()

    if not args.print_only:
        write_version(version)
    print(version)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
