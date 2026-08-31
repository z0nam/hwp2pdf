import argparse
import sys
from pathlib import Path

from hwp2pdf import config
from hwp2pdf.backends import BackendUnavailable, create_backend
from hwp2pdf.backends.windows_com import get_hwp_processes, kill_hwp
from hwp2pdf.constants import APP_NAME, enabled_extensions
from hwp2pdf.i18n import translate
from hwp2pdf.jobs import collect_files, run_batch
from hwp2pdf.version import __version__

# Importing hwp2pdf.app would pull in tkinter; the CLI and the conversion
# server deliberately stay GUI-free.


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
    collect_files = staticmethod(collect_files)

    def __init__(self, backend_settings=None):
        self.log_queue = CliEventSink()
        self.stop_requested = False
        self.backend_settings = backend_settings


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
    parser.add_argument(
        "--server",
        default="",
        help="Convert on a Windows conversion server, e.g. http://host:8765 "
             "(defaults to the saved setting or $HWP2PDF_SERVER_URL)",
    )
    parser.add_argument(
        "--token", default="", help="Bearer token for --server (or $HWP2PDF_TOKEN)"
    )
    parser.add_argument(
        "--rhwp-fallback",
        action="store_true",
        help="If the conversion server cannot be reached, render PDFs locally "
             "with rhwp. PDF only, and approximate -- see docs/remote-server.md.",
    )
    parser.add_argument(
        "--rhwp-path", default="", help="Path to the rhwp executable"
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=0,
        help="Force-close and restart Hangul if one conversion exceeds this many "
             "seconds (local conversion only). 0 waits forever.",
    )
    parser.add_argument(
        "--transport",
        choices=config.TRANSPORTS,
        default=None,
        help="How sources reach the server: auto, upload, or share",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def resolve_backend_settings(args):
    """Command line wins over saved settings, which win over defaults."""
    settings = config.server_settings()
    if args.server:
        settings["url"] = args.server
    if args.token:
        settings["token"] = args.token
    if args.transport:
        settings["transport"] = args.transport
    return settings


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
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv and argv[0] == "serve":
        from hwp2pdf.serve import main as serve_main

        return serve_main(argv[1:])

    parser = build_parser()
    args = parser.parse_args(argv)

    backend_settings = resolve_backend_settings(args)
    remote = bool(backend_settings.get("url"))

    target = Path(args.target).expanduser()
    try:
        validate_target(target)
        if not remote:
            prepare_hwp_processes(args)
    except Exception as e:
        parser.exit(1, f"ERROR: {e}\n")

    context = CliConversionContext(backend_settings)
    try:
        backend = create_backend(
            backend_settings, "ko",
            rhwp_fallback=args.rhwp_fallback,
            rhwp_path=args.rhwp_path,
        )
    except BackendUnavailable as e:
        parser.exit(1, f"ERROR: {e}\n")

    if args.timeout and hasattr(backend, "job_timeout"):
        backend.job_timeout = args.timeout

    run_batch(
        context.log_queue,
        backend,
        target=str(target),
        recursive=args.recursive,
        overwrite=not args.no_overwrite,
        use_safe_copy=not args.no_safe_temp,
        force_one_page=not args.no_force_one_page,
        output_formats=selected_formats(args),
        lang="ko",
        is_stopped=lambda: context.stop_requested,
        file_collector=context.collect_files,
    )
    return context.log_queue.exit_code


if __name__ == "__main__":
    raise SystemExit(main())
