"""Entry point shared by the GUI and the CLI.

macOS hands a Finder "Open with" over as plain ``sys.argv`` entries (the app
bundle enables argv emulation and declares HWP/HWPX document types). Those must
open the window with the files selected -- not silently start a headless
conversion, which is what treating any argument as a CLI invocation would do.
"""

import os
import stat
import sys
from pathlib import Path

from hwp2pdf.app import main as gui_main
from hwp2pdf.cli import main as cli_main
from hwp2pdf.constants import enabled_extensions


def _keep_tk_from_building_a_console() -> None:
    """Stop Tk from creating the Tcl console window it never needed.

    Tk builds one when standard input is a zero-block character device, which
    is exactly what launchd hands a bundled app: the same binary started from a
    terminal takes a different path and is unaffected. Building that window's
    menu bar aborts intermittently on macOS -- an NSMenuItem assertion inside
    tkSetMainMenu, before a single line of this app runs -- and the app dies
    before its window appears.

    A pipe with its write end already closed reads as empty like /dev/null, but
    is a FIFO rather than a character device, so the console is never built.
    Frozen macOS builds only: a terminal or a test runner has a stdin worth
    keeping.
    """
    if sys.platform != "darwin" or not getattr(sys, "frozen", False):
        return
    try:
        if os.isatty(0):
            return
        info = os.fstat(0)
        # getattr and "not": st_blocks is None on some platforms and missing
        # entirely on Windows, while Tk reads the raw struct where it is zero.
        replace = stat.S_ISCHR(info.st_mode) and not getattr(info, "st_blocks", 0)
    except OSError:
        replace = True          # no usable fd 0 at all, which Tk also consoles for
    if not replace:
        return
    try:
        read_fd, write_fd = os.pipe()
        os.close(write_fd)
        os.dup2(read_fd, 0)
        os.close(read_fd)
    except OSError:
        pass                    # worst case is the console Tk would have built


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
    _keep_tk_from_building_a_console()
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
