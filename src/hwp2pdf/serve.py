"""``hwp2pdf serve`` -- run the Windows conversion server.

Must run inside an interactive desktop session. Registering this as a Windows
Service puts Hangul in Session 0, where it has no desktop and leaves zombie
``Hwp.exe`` processes behind (see docs/known-issues.md).
"""

import argparse
import os
import secrets
import subprocess
import sys
import time

from hwp2pdf import paths
from hwp2pdf.constants import APP_NAME
from hwp2pdf.server import protocol
from hwp2pdf.version import __version__

LOOPBACK = {"127.0.0.1", "::1", "localhost"}
TAILSCALE_BINARIES = (
    "tailscale",
    r"C:\Program Files\Tailscale\tailscale.exe",
    r"C:\Program Files (x86)\Tailscale\tailscale.exe",
    "/usr/local/bin/tailscale",
    "/opt/homebrew/bin/tailscale",
)


def tailscale_address():
    """First IPv4 address Tailscale has assigned to this machine."""
    for binary in TAILSCALE_BINARIES:
        try:
            result = subprocess.run(
                [binary, "ip", "-4"], capture_output=True, text=True, timeout=10
            )
        except (OSError, subprocess.SubprocessError):
            continue
        if result.returncode == 0:
            for line in result.stdout.splitlines():
                address = line.strip()
                if address:
                    return address
    return None


def resolve_bind(value: str) -> str:
    if value != "tailscale":
        return value
    address = tailscale_address()
    if not address:
        raise SystemExit(
            "ERROR: could not determine this machine's Tailscale IPv4 address.\n"
            "Is Tailscale running and logged in? Use --bind <address> instead."
        )
    return address


def parse_share_root(value: str):
    name, sep, path = value.partition("=")
    if not sep or not name.strip() or not path.strip():
        raise argparse.ArgumentTypeError("expected NAME=PATH, e.g. work=D:\\shared")
    return name.strip(), path.strip()


def load_or_create_token(explicit: str, create: bool) -> str:
    if explicit:
        return explicit
    env = os.environ.get("HWP2PDF_TOKEN")
    if env:
        return env

    token_path = paths.server_token_path()
    try:
        existing = token_path.read_text(encoding="utf-8").strip()
    except OSError:
        existing = ""
    if existing:
        return existing
    if not create:
        return ""

    token = secrets.token_urlsafe(32)
    token_path.parent.mkdir(parents=True, exist_ok=True)
    token_path.write_text(token + "\n", encoding="utf-8")
    if os.name != "nt":
        os.chmod(token_path, 0o600)
    return token


def build_parser():
    parser = argparse.ArgumentParser(
        prog="hwp2pdf serve",
        description=f"{APP_NAME} conversion server v{__version__}",
    )
    parser.add_argument(
        "--bind",
        default="127.0.0.1",
        help="Address to listen on. Use 'tailscale' to bind only this machine's "
             "Tailscale IP, an explicit address, or 0.0.0.0 for every interface.",
    )
    parser.add_argument("--port", type=int, default=protocol.DEFAULT_PORT)
    parser.add_argument("--token", default="", help="Bearer token clients must send")
    parser.add_argument(
        "--init", action="store_true",
        help="Generate and store a token if none exists yet, then print it",
    )
    parser.add_argument(
        "--no-auth", action="store_true",
        help="Disable authentication (only allowed when binding to loopback)",
    )
    parser.add_argument(
        "--share-root", action="append", default=[], type=parse_share_root,
        metavar="NAME=PATH", help="Expose a shared folder for path-passthrough conversion",
    )
    parser.add_argument("--max-upload-bytes", type=int, default=protocol.DEFAULT_MAX_UPLOAD_BYTES)
    parser.add_argument("--max-queue", type=int, default=protocol.DEFAULT_MAX_QUEUE)
    parser.add_argument("--job-ttl", type=int, default=protocol.DEFAULT_JOB_TTL_SECONDS)
    parser.add_argument("--tls-cert", default="")
    parser.add_argument("--tls-key", default="")
    parser.add_argument("--quiet", action="store_true", help="Do not log every request")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def banner(args, bind, token, share_roots, probe):
    print(f"{APP_NAME} conversion server v{__version__} (API {protocol.API_VERSION})")
    print(f"  listening   http://{bind}:{args.port}")
    if bind in LOOPBACK:
        print("              (loopback only -- reachable from this machine and its VMs)")
    print(f"  auth        {'token ' + token[:6] + '...' if token else 'DISABLED'}")
    print(f"  hangul      {'yes' if probe['installed'] else 'NO'} ({probe['detail']})")
    if probe["running"]:
        print(f"              Hwp.exe already running: {', '.join(probe['running'])}")
    print(f"  shares      {', '.join(sorted(share_roots)) if share_roots else '(none)'}")
    print(f"  max upload  {args.max_upload_bytes // (1024 * 1024)} MB")
    print("  note        keep this window open in a logged-in desktop session;")
    print("              Hangul automation does not work as a Windows Service.")
    print("Press Ctrl+C to stop.")


def main(argv=None) -> int:
    from hwp2pdf.backends.windows_com import WindowsComBackend, probe_hwp
    from hwp2pdf.server.http_server import create_server

    args = build_parser().parse_args(argv)
    bind = resolve_bind(args.bind)

    if args.no_auth:
        if bind not in LOOPBACK:
            raise SystemExit(
                f"ERROR: --no-auth is only allowed with a loopback --bind, not {bind}."
            )
        token = ""
    else:
        token = load_or_create_token(args.token, create=args.init or bind not in LOOPBACK)
        if not token and bind not in LOOPBACK:
            raise SystemExit(
                "ERROR: a token is required when binding to a non-loopback address.\n"
                "Run with --init to generate one, or pass --token."
            )

    share_roots = dict(args.share_root)
    probe = probe_hwp()

    httpd = create_server(
        bind,
        args.port,
        backend_factory=WindowsComBackend,
        hwp_probe=probe_hwp,
        token=token,
        share_roots=share_roots,
        max_upload_bytes=args.max_upload_bytes,
        max_queue=args.max_queue,
        job_ttl=args.job_ttl,
        tls_cert=args.tls_cert,
        tls_key=args.tls_key,
        quiet=args.quiet,
    )

    banner(args, bind, token, share_roots, probe)
    if args.init and token:
        print(f"\n  token: {token}\n  stored in {paths.server_token_path()}\n")

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nshutting down...")
    finally:
        httpd.shutdown()
        httpd.store.shutdown()
        httpd.server_close()
        time.sleep(0.1)
    return 0


if __name__ == "__main__":
    sys.exit(main())
