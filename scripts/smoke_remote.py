"""End-to-end smoke test against a running hwp2pdf conversion server.

Standard library only, so it runs from a clean macOS checkout:

    python scripts/smoke_remote.py http://namun-ji.<tailnet>.ts.net:17650 <token> sample.hwp

Exits non-zero on the first failure.
"""

import json
import shutil
import sys
import tempfile
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from hwp2pdf.backends.remote_http import RemoteHttpBackend  # noqa: E402
from hwp2pdf.jobs import LOG_CSV_NAME, run_batch  # noqa: E402
from hwp2pdf.server import protocol  # noqa: E402

PDF_MAGIC = b"%PDF-"
DOCX_MAGIC = b"PK\x03\x04"


class Sink:
    def __init__(self, verbose=True):
        self.verbose = verbose
        self.events = []

    def put(self, item):
        self.events.append(item)
        kind, payload = item
        if not self.verbose:
            return
        if kind == "log":
            text = payload[0] if isinstance(payload, tuple) else payload
            print(f"    {text}")
        elif kind == "error":
            print(f"    ERROR: {payload}")

    def done(self):
        for kind, payload in self.events:
            if kind == "done":
                return payload
        return None


FAILURES = []


def check(label, condition, detail=""):
    mark = "ok  " if condition else "FAIL"
    print(f"[{mark}] {label}" + (f" -- {detail}" if detail else ""))
    if not condition:
        FAILURES.append(label)
    return condition


def request(url, token, path, method="GET"):
    req = urllib.request.Request(url.rstrip("/") + path, method=method)
    if token:
        req.add_header(protocol.AUTH_HEADER, f"{protocol.AUTH_SCHEME} {token}")
    with urllib.request.urlopen(req, timeout=15) as response:
        raw = response.read()
        return response.status, (json.loads(raw.decode("utf-8")) if raw else {})


def convert(url, token, sample: Path, formats, workdir: Path):
    staged = workdir / sample.name
    shutil.copy2(sample, staged)
    sink = Sink()
    started = time.monotonic()
    run_batch(
        sink,
        RemoteHttpBackend({"url": url, "token": token, "transport": "upload", "shares": []}),
        target=str(workdir),
        recursive=False,
        overwrite=True,
        use_safe_copy=True,
        force_one_page=True,
        output_formats=formats,
        lang="ko",
    )
    return sink, time.monotonic() - started


def main(argv):
    if len(argv) < 3:
        print(__doc__)
        return 2
    url, token, sample_arg = argv[0], argv[1], argv[2]
    sample = Path(sample_arg).expanduser().resolve()

    if not check("sample file exists", sample.is_file(), str(sample)):
        return 1

    print(f"\n== {url} ==")

    try:
        status, health = request(url, "", protocol.PATH_HEALTH)
    except (urllib.error.URLError, OSError) as e:
        check("health reachable", False, str(e))
        return 1
    check("health returns 200", status == 200)
    check("api version matches", health.get("api") == protocol.API_VERSION,
          f"server api={health.get('api')} client api={protocol.API_VERSION}")

    _status, caps = request(url, token, protocol.PATH_CAPABILITIES)
    check("hangul installed on server", bool(caps.get("hwp_installed")), caps.get("hwp_detail", ""))
    check("PDF supported", "PDF" in (caps.get("formats") or []))
    check("DOCX supported", "DOCX" in (caps.get("formats") or []))
    print(f"       server v{caps.get('version')}, shares={caps.get('shares')}, "
          f"queue={caps.get('queue_depth')}")

    with tempfile.TemporaryDirectory(prefix="hwp2pdf-smoke-") as raw_dir:
        workdir = Path(raw_dir)

        print("\n-- PDF conversion")
        sink, elapsed = convert(url, token, sample, ("PDF",), workdir)
        pdf = workdir / (sample.stem + ".pdf")
        check("PDF produced", pdf.exists())
        if pdf.exists():
            data = pdf.read_bytes()
            check("PDF has a PDF header", data.startswith(PDF_MAGIC))
            check("PDF is larger than 1 KB", len(data) > 1024, f"{len(data)} bytes")
        check("PDF conversion under 120 s", elapsed < 120, f"{elapsed:.1f}s")
        check("batch reported success", (sink.done() or (0, 1, 0, "", False))[0] == 1)

        print("\n-- DOCX conversion")
        sink, _elapsed = convert(url, token, sample, ("DOCX",), workdir)
        docx = workdir / (sample.stem + ".docx")
        check("DOCX produced", docx.exists())
        if docx.exists():
            check("DOCX is a zip container", docx.read_bytes().startswith(DOCX_MAGIC))

        print("\n-- audit log")
        check("clean run leaves no CSV", not (workdir / LOG_CSV_NAME).exists())

    print("\n-- auth and error handling")
    try:
        request(url, token + "-wrong", protocol.PATH_CAPABILITIES)
        check("bad token rejected", False)
    except urllib.error.HTTPError as e:
        check("bad token rejected with 401", e.code == 401, f"got {e.code}")

    try:
        request(url, token, protocol.job_path("0" * 32) + "/events?cursor=0&wait=0")
        check("unknown job rejected", False)
    except urllib.error.HTTPError as e:
        check("unknown job rejected with 404", e.code == 404, f"got {e.code}")

    print()
    if FAILURES:
        print(f"FAILED: {len(FAILURES)} check(s): {', '.join(FAILURES)}")
        return 1
    print("All checks passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
