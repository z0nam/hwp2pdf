"""Stdlib HTTP front end for the conversion server.

``http.server`` is deliberate: the project ships PyInstaller onefile builds with
almost no runtime dependencies, the workload is a handful of endpoints on a
private network, and conversion is serialized anyway. ``ThreadingHTTPServer``
only exists so a long-polling event request cannot block an upload.
"""

import hmac
import json
import os
import re
import shutil
import ssl
import sys
import threading
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from hwp2pdf.constants import APP_NAME, OUTPUT_FORMATS, enabled_extensions
from hwp2pdf.server import protocol
from hwp2pdf.server.jobs import Item, JobStore, QueueFull
from hwp2pdf.version import __version__

_RE_INPUT = re.compile(r"^/v1/jobs/([0-9a-f]{32})/inputs/([A-Za-z0-9._-]{1,64})$")
_RE_RUN = re.compile(r"^/v1/jobs/([0-9a-f]{32})/items/([A-Za-z0-9._-]{1,64})/run$")
_RE_OUTPUT = re.compile(r"^/v1/jobs/([0-9a-f]{32})/outputs/([A-Za-z0-9._-]{1,64})$")
_RE_EVENTS = re.compile(r"^/v1/jobs/([0-9a-f]{32})/events$")
_RE_CANCEL = re.compile(r"^/v1/jobs/([0-9a-f]{32})/cancel$")
_RE_JOB = re.compile(r"^/v1/jobs/([0-9a-f]{32})$")

COPY_CHUNK = 1024 * 1024


class ConversionHTTPServer(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True

    def __init__(self, address, handler_class, *, store, token, max_upload_bytes,
                 hwp_probe, log_sink=None):
        self.store = store
        self.token = token
        self.max_upload_bytes = max_upload_bytes
        self.hwp_probe = hwp_probe
        # In a windowed build sys.stderr is None, so request logging has to go
        # somewhere the stdlib default would not reach.
        self.log_sink = log_sink
        super().__init__(address, handler_class)


class Handler(BaseHTTPRequestHandler):
    server_version = f"hwp2pdf/{__version__}"
    protocol_version = "HTTP/1.1"

    # -- plumbing --------------------------------------------------------
    def log_message(self, fmt, *args):  # quieter than the stdlib default
        if self.server.quiet:
            return
        line = f"{self.address_string()} {fmt % args}"
        sink = getattr(self.server, "log_sink", None)
        if sink is not None:
            sink(line)
        elif sys.stderr is not None:
            super().log_message(fmt, *args)

    def _send(self, status, payload=None, body=None, content_type="application/json", headers=None):
        if body is None:
            body = b"" if payload is None else json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _error(self, status, message):
        self._send(status, {"error": message})

    def _authorized(self) -> bool:
        token = self.server.token
        if not token:
            return True
        header = self.headers.get(protocol.AUTH_HEADER, "")
        prefix = protocol.AUTH_SCHEME + " "
        if not header.startswith(prefix):
            return False
        return hmac.compare_digest(header[len(prefix):].strip(), token)

    def _require_auth(self) -> bool:
        if self._authorized():
            return True
        self.send_response(HTTPStatus.UNAUTHORIZED)
        self.send_header("WWW-Authenticate", protocol.AUTH_SCHEME)
        self.send_header("Content-Length", "0")
        self.end_headers()
        return False

    def _read_json(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return {}
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except ValueError:
            return None

    def _job_or_404(self, job_id):
        job = self.server.store.get(job_id)
        if job is None:
            self._error(HTTPStatus.NOT_FOUND, "unknown job")
        return job

    def _query(self):
        if "?" not in self.path:
            return self.path, {}
        route, _, raw = self.path.partition("?")
        params = {}
        for pair in raw.split("&"):
            if not pair:
                continue
            key, _, value = pair.partition("=")
            params[key] = value
        return route, params

    # -- routes ----------------------------------------------------------
    def do_GET(self):
        route, params = self._query()

        if route == protocol.PATH_HEALTH:
            self._send(HTTPStatus.OK, {
                "app": APP_NAME,
                "version": __version__,
                "api": protocol.API_VERSION,
                "auth_required": bool(self.server.token),
            })
            return

        if not self._require_auth():
            return

        if route == protocol.PATH_CAPABILITIES:
            probe = self.server.hwp_probe()
            self._send(HTTPStatus.OK, {
                "app": APP_NAME,
                "version": __version__,
                "api": protocol.API_VERSION,
                "os": os.name,
                "hwp_installed": probe["installed"],
                "hwp_detail": probe["detail"],
                "hwp_running": probe["running"],
                "formats": sorted(OUTPUT_FORMATS),
                "extensions": list(enabled_extensions()),
                "shares": sorted(self.server.store.share_roots),
                "max_upload_bytes": self.server.max_upload_bytes,
                "queue_depth": self.server.store.queue_depth(),
            })
            return

        match = _RE_EVENTS.match(route)
        if match:
            job = self._job_or_404(match.group(1))
            if job is None:
                return
            try:
                cursor = max(0, int(params.get("cursor", "0")))
                wait = min(protocol.DEFAULT_EVENT_WAIT_SECONDS, max(0, int(params.get("wait", "0"))))
            except ValueError:
                self._error(HTTPStatus.BAD_REQUEST, "bad cursor or wait")
                return
            events, next_cursor = job.events_since(cursor, wait)
            self._send(HTTPStatus.OK, {
                "events": events,
                "cursor": next_cursor,
                "cancelled": job.cancelled,
                "queue_depth": self.server.store.queue_depth(),
            })
            return

        match = _RE_OUTPUT.match(route)
        if match:
            job = self._job_or_404(match.group(1))
            if job is None:
                return
            item = job.items.get(match.group(2))
            if item is None:
                self._error(HTTPStatus.NOT_FOUND, "unknown item")
                return
            path = item.target_path
            if path is None or not path.exists():
                self._error(HTTPStatus.NOT_FOUND, "no output")
                return
            self._send_file(path)
            return

        self._error(HTTPStatus.NOT_FOUND, "unknown route")

    def do_POST(self):
        route, _params = self._query()
        if not self._require_auth():
            return

        if route == protocol.PATH_JOBS:
            payload = self._read_json()
            if payload is None:
                self._error(HTTPStatus.BAD_REQUEST, "bad json")
                return
            lang = payload.get("lang") or "ko"
            job = self.server.store.create_job(lang=lang, safe_temp=bool(payload.get("safe_temp", True)))
            self._send(HTTPStatus.CREATED, {
                "job_id": job.job_id,
                "queue_depth": self.server.store.queue_depth(),
            })
            return

        match = _RE_RUN.match(route)
        if match:
            job = self._job_or_404(match.group(1))
            if job is None:
                return
            payload = self._read_json()
            if payload is None:
                self._error(HTTPStatus.BAD_REQUEST, "bad json")
                return
            output_format = payload.get("output_format")
            if output_format not in OUTPUT_FORMATS:
                self._error(HTTPStatus.BAD_REQUEST, "unsupported output format")
                return
            share = payload.get("share") or ""
            if share:
                try:
                    self.server.store.resolve_share(share, payload.get("rel") or "")
                    self.server.store.resolve_share(share, payload.get("out_rel") or "")
                except ValueError as e:
                    self._error(HTTPStatus.BAD_REQUEST, str(e))
                    return
            elif not job.input_path(match.group(2)).exists():
                self._error(HTTPStatus.BAD_REQUEST, "input not uploaded")
                return

            item = Item(
                item_id=match.group(2),
                name=str(payload.get("name") or match.group(2)),
                output_format=output_format,
                force_one_page=bool(payload.get("force_one_page", True)),
                share=share,
                rel=payload.get("rel") or "",
                out_rel=payload.get("out_rel") or "",
            )
            try:
                self.server.store.submit(job, item)
            except QueueFull:
                self._error(HTTPStatus.TOO_MANY_REQUESTS, "conversion queue is full")
                return
            self._send(HTTPStatus.ACCEPTED, {"queue_depth": self.server.store.queue_depth()})
            return

        match = _RE_CANCEL.match(route)
        if match:
            if not self.server.store.cancel_job(match.group(1)):
                self._error(HTTPStatus.NOT_FOUND, "unknown job")
                return
            self._send(HTTPStatus.OK, {"cancelled": True})
            return

        self._error(HTTPStatus.NOT_FOUND, "unknown route")

    def do_PUT(self):
        route, _params = self._query()
        if not self._require_auth():
            return

        match = _RE_INPUT.match(route)
        if not match:
            self._error(HTTPStatus.NOT_FOUND, "unknown route")
            return

        job = self._job_or_404(match.group(1))
        if job is None:
            return

        try:
            length = int(self.headers.get("Content-Length") or -1)
        except ValueError:
            length = -1
        if length < 0:
            self._error(HTTPStatus.LENGTH_REQUIRED, "Content-Length required")
            return
        if length > self.server.max_upload_bytes:
            self._error(HTTPStatus.REQUEST_ENTITY_TOO_LARGE, "file too large")
            return

        target = job.input_path(match.group(2))
        tmp = target.with_name(target.name + ".part")
        remaining = length
        try:
            with open(tmp, "wb") as f:
                while remaining > 0:
                    chunk = self.rfile.read(min(COPY_CHUNK, remaining))
                    if not chunk:
                        break
                    f.write(chunk)
                    remaining -= len(chunk)
            if remaining:
                raise OSError("short upload")
            os.replace(tmp, target)
        except OSError as e:
            tmp.unlink(missing_ok=True)
            self._error(HTTPStatus.INTERNAL_SERVER_ERROR, f"upload failed: {e}")
            return

        job.touch()
        self._send(HTTPStatus.NO_CONTENT)

    def do_DELETE(self):
        route, _params = self._query()
        if not self._require_auth():
            return
        match = _RE_JOB.match(route)
        if not match:
            self._error(HTTPStatus.NOT_FOUND, "unknown route")
            return
        if not self.server.store.delete_job(match.group(1)):
            self._error(HTTPStatus.NOT_FOUND, "unknown job")
            return
        self._send(HTTPStatus.OK, {"deleted": True})

    def do_HEAD(self):
        self.do_GET()

    def _send_file(self, path: Path):
        size = path.stat().st_size
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/octet-stream")
        self.send_header("Content-Length", str(size))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        if self.command == "HEAD":
            return
        with open(path, "rb") as f:
            shutil.copyfileobj(f, self.wfile, COPY_CHUNK)


def create_server(
    bind: str,
    port: int,
    *,
    backend_factory,
    hwp_probe,
    token: str = "",
    share_roots=None,
    max_upload_bytes: int = protocol.DEFAULT_MAX_UPLOAD_BYTES,
    max_queue: int = protocol.DEFAULT_MAX_QUEUE,
    job_ttl: float = protocol.DEFAULT_JOB_TTL_SECONDS,
    tls_cert: str = "",
    tls_key: str = "",
    quiet: bool = False,
    log_sink=None,
):
    store = JobStore(
        backend_factory=backend_factory,
        share_roots=share_roots,
        max_queue=max_queue,
        job_ttl=job_ttl,
    )
    httpd = ConversionHTTPServer(
        (bind, port),
        Handler,
        store=store,
        token=token,
        max_upload_bytes=max_upload_bytes,
        hwp_probe=hwp_probe,
        log_sink=log_sink,
    )
    httpd.quiet = quiet
    if tls_cert and tls_key:
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        context.load_cert_chain(tls_cert, tls_key)
        httpd.socket = context.wrap_socket(httpd.socket, server_side=True)
    store.start()
    return httpd


def serve_forever(httpd):
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    return thread
