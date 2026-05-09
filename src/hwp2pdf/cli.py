import argparse
import sys
from pathlib import Path

from hwp2pdf.app import (
    APP_NAME,
    ConverterApp,
    enabled_extensions,
    get_hwp_processes,
    kill_hwp,
    translate,
)
from hwp2pdf.version import __version__


class CliEventSink:
    def __init__(self):
        self.exit_code = 0

    def put(self, item):
        kind, payload = item
        if kind == "log":
            if isinstance(payload, tuple):
                text, level = payload
                prefix = "ERROR: " if level == "error" else "WARN: " if level == "warning" else ""
                print(prefix + str(text), flush=True)
            else:
                print(payload, flush=True)
        elif kind == "progress":
            _current, _total, label = payload
            print(label, flush=True)
        elif kind == "done":
            success, failed, skipped, log_csv, all_success = payload
            if all_success:
                print(translate("ko", "success_message"), flush=True)
            else:
                print(
                    translate("ko", "done_status", success=success, failed=failed, skipped=skipped),
                    flush=True,
                )
                print(f"Log: {log_csv}", flush=True)
            self.exit_code = 0 if all_success else 2
        elif kind == "error":
            print(f"ERROR: {payload}", file=sys.stderr, flush=True)
            self.exit_code = 1


class CliConversionContext:
    collect_files = staticmethod(ConverterApp.collect_files)

    def __init__(self):
        self.log_queue = CliEventSink()
        self.stop_requested = False


def build_parser():
    parser = argparse.ArgumentParser(
        prog="hwp2pdf",
        description=f"{APP_NAME} v{__version__}",
    )
    parser.add_argument("target", help="HWP/HWPX file or folder to convert")
    parser.add_argument("--pdf", action="store_true", help="Export PDF")
    parser.add_argument("--docx", action="store_true", help="Export DOCX")
    parser.add_argument("-r", "--recursive", action="store_true", help="Include subfolders when target is a folder")
    parser.add_argument("--no-overwrite", action="store_true", help="Skip outputs that already exist")
    parser.add_argument("--no-safe-temp", action="store_true", help="Do not copy files through the safe local temp folder")
    parser.add_argument("--no-force-one-page", action="store_true", help="Do not reset one-page / N-up PDF print settings")
    parser.add_argument("--kill-hwp", action="store_true", help="Force close running HWP processes before conversion")
    parser.add_argument(
        "--allow-running-hwp",
        action="store_true",
        help="Continue even if HWP is already running",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def selected_formats(args):
    formats = []
    if args.pdf:
        formats.append("PDF")
    if args.docx:
        formats.append("DOCX")
    return tuple(formats or ("PDF",))


def validate_target(target: Path):
    if target.is_file() and target.suffix.lower() not in enabled_extensions():
        raise ValueError("Select an HWP or HWPX file.")
    if not target.exists():
        raise ValueError(f"Target not found: {target}")
    if not (target.is_file() or target.is_dir()):
        raise ValueError(f"Target must be a file or folder: {target}")


def prepare_hwp_processes(args):
    processes = get_hwp_processes()
    if not processes:
        return

    detail = ", ".join(f"PID {process['pid']}" for process in processes)
    if args.kill_hwp:
        kill_hwp()
        remaining = get_hwp_processes()
        if remaining:
            remaining_detail = ", ".join(f"PID {process['pid']}" for process in remaining)
            raise RuntimeError(f"Could not close running HWP processes: {remaining_detail}")
        return

    if args.allow_running_hwp:
        print(f"WARN: HWP is already running: {detail}", flush=True)
        return

    raise RuntimeError(
        "HWP is already running. Close it first, or run with --kill-hwp / --allow-running-hwp. "
        f"Detected: {detail}"
    )


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)

    target = Path(args.target).expanduser()
    try:
        validate_target(target)
        prepare_hwp_processes(args)
    except Exception as e:
        parser.exit(1, f"ERROR: {e}\n")

    context = CliConversionContext()
    ConverterApp._run_conversion(
        context,
        str(target),
        args.recursive,
        not args.no_overwrite,
        not args.no_safe_temp,
        not args.no_force_one_page,
        selected_formats(args),
        "ko",
    )
    return context.log_queue.exit_code


if __name__ == "__main__":
    raise SystemExit(main())
