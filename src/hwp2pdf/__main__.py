import sys

from hwp2pdf.app import main as gui_main
from hwp2pdf.cli import main as cli_main


if __name__ == "__main__":
    if len(sys.argv) > 1:
        raise SystemExit(cli_main())
    gui_main()
