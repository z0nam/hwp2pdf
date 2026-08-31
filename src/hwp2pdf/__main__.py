"""Entry point shared by the GUI and the CLI.

macOS hands a Finder "Open with" over as plain ``sys.argv`` entries (the app
bundle enables argv emulation and declares HWP/HWPX document types). Those must
open the window with the files selected -- not silently start a headless
conversion, which is what treating any argument as a CLI invocation would do.
"""

import sys
from pathlib import Path

from hwp2pdf.app import main as gui_main
from hwp2pdf.cli import main as cli_main
from hwp2pdf.constants import enabled_extensions


def looks_like_documents(argv) -> bool:
    """True when every argument is an existing HWP/HWPX file."""
    if not argv:
        return False
    allowed = enabled_extensions()
    return all(
        not arg.startswith("-")
        and Path(arg).suffix.lower() in allowed
        and Path(arg).is_file()
        for arg in argv
    )


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv:
        gui_main()
        return 0
    if looks_like_documents(argv):
        gui_main(initial_paths=argv)
        return 0
    return cli_main(argv)


if __name__ == "__main__":
    raise SystemExit(main())
