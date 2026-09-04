"""Client backend that runs conversions on a Windows conversion server.

Mirrors :class:`~hwp2pdf.backends.windows_com.WindowsComBackend` one call at a
time, so ``jobs.run_batch`` cannot tell the difference: the destination
filesystem, the CSV log and the progress events all stay on this machine.
"""

import json
import os
import time
import urllib.error
import urllib.request
from pathlib import Path

from hwp2pdf import certs
from hwp2pdf.backends.base import BackendCapabilities, BackendUnavailable, JobResult
from hwp2pdf.constants import OUTPUT_FORMATS
from hwp2pdf.i18n import translate
from hwp2pdf.server import protocol

CONNECT_TIMEOUT = 10
JOB_TIMEOUT = 15
EVENT_WAIT = 25
EVENT_TIMEOUT = EVENT_WAIT + 10
DOWNLOAD_TIMEOUT = 120
UPLOAD_RATE_BYTES_PER_SECOND = 50_000
RETRY_DELAYS = (1, 2, 4)
SHARE_OUTPUT_GRACE_SECONDS = 5.0
COPY_CHUNK = 1024 * 1024


class RemoteError(Exception):
    """Transport or protocol failure that should abandon the whole session."""


class RemoteHttpBackend:
    capabilities = BackendCapabilities(
        name="remote_http",
        remote=True,
        local_staging=False,
        manages_hwp_process=False,
        local_preflight=False,
    )

    def __init__(self, server: dict):
        self.base_url = (server.get("url") or "").rstrip("/")
        self.token = server.get("token") or ""
        self.transport = server.get("transport") or "auto"
        self.shares = list(server.get("shares") or [])
        self.sink = None
        self.job_id = None
        self.cursor = 0
        self.capabilities_payload = {}
        self.server_shares = []
        self.cancelled = False

    # -- HTTP helpers ----------------------------------------------------
    def _url(self, path: str) -> str:
        return f"{self.base_url}{path}"

    def _open(self, method, path, *, data=None, timeout, headers=None, authorize=True):
        request = urllib.request.Request(self._url(path), data=data, method=method)
        if authorize and self.token:
            request.add_header(protocol.AUTH_HEADER, f"{protocol.AUTH_SCHEME} {self.token}")
        for key, value in (headers or {}).items():
            request.add_header(key, value)
        return certs.urlopen(request, timeout=timeout)

    def _json(self, method, path, *, payload=None, timeout, retries=0, authorize=True):
        body = None
        headers = {}
        if payload is not None:
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json"

        attempt = 0
        while True:
            try:
                with self._open(method, path, data=body, timeout=timeout,
                                headers=headers, authorize=authorize) as response:
                    raw = response.read()
                    if not raw:
                        return {}
                    return json.loads(raw.decode("utf-8"))
            except urllib.error.HTTPError as e:
                raise self._http_error(e) from None
            except (urllib.error.URLError, OSError, ValueError) as e:
                if attempt >= retries:
                    raise RemoteError(translate(self._lang(), "remote_unreachable",
                                                url=self.base_url, detail=e)) from None
                time.sleep(RETRY_DELAYS[min(attempt, len(RETRY_DELAYS) - 1)])
                attempt += 1

    def _http_error(self, e):
        lang = self._lang()
        if e.code in (401, 403):
            return BackendUnavailable(
                translate(lang, "remote_auth_failed"), fallback_allowed=False
            )
        if e.code == 429:
            return RemoteError(translate(lang, "remote_server_busy"))
        if e.code == 413:
            return RemoteError(translate(lang, "remote_upload_too_large"))
        detail = ""
        try:
            detail = json.loads(e.read().decode("utf-8")).get("error", "")
        except Exception:
            pass
        return RemoteError(translate(lang, "remote_http_error", status=e.code, detail=detail))

    def _lang(self):
        return getattr(self, "_session_lang", "ko")

    # -- backend contract -------------------------------------------------
    def preflight(self, lang: str) -> None:
        self._session_lang = lang
        if not self.base_url:
            raise BackendUnavailable(translate(lang, "server_not_configured"))

        try:
            health = self._json("GET", protocol.PATH_HEALTH, timeout=CONNECT_TIMEOUT,
                                retries=len(RETRY_DELAYS), authorize=False)
        except BackendUnavailable:
            raise
        except RemoteError as e:
            raise BackendUnavailable(str(e)) from None

        if health.get("api") != protocol.API_VERSION:
            raise BackendUnavailable(
                translate(lang, "remote_version_mismatch",
                          server=health.get("version", "?"), api=health.get("api", "?"),
                          client_api=protocol.API_VERSION),
                fallback_allowed=False,
            )

        if health.get("auth_required") and not self.token:
            raise BackendUnavailable(
                translate(lang, "remote_auth_failed"), fallback_allowed=False
            )

        try:
            caps = self._json("GET", protocol.PATH_CAPABILITIES, timeout=CONNECT_TIMEOUT,
                              retries=len(RETRY_DELAYS))
        except BackendUnavailable:
            raise
        except RemoteError as e:
            raise BackendUnavailable(str(e)) from None

        if not caps.get("hwp_installed"):
            raise BackendUnavailable(
                translate(lang, "remote_hwp_missing", detail=caps.get("hwp_detail", ""))
            )

        self.capabilities_payload = caps
        self.server_shares = list(caps.get("shares") or [])

    def open_session(self, sink, lang: str, options) -> None:
        self._session_lang = lang
        self.sink = sink
        self.cancelled = False
        self.cursor = 0

        try:
            created = self._json(
                "POST", protocol.PATH_JOBS,
                payload={"lang": lang, "safe_temp": bool(getattr(options, "safe_temp", True))},
                timeout=JOB_TIMEOUT,
            )
        except BackendUnavailable:
            raise
        except RemoteError as e:
            raise BackendUnavailable(str(e)) from None
        self.job_id = created.get("job_id")
        if not self.job_id:
            raise BackendUnavailable(
                translate(lang, "remote_unreachable", url=self.base_url, detail="no job id")
            )

        sink.put(("log", translate(
            lang, "remote_connected",
            url=self.base_url,
            version=self.capabilities_payload.get("version", "?"),
        )))

    def session_notes(self, lang: str) -> list:
        mode = translate(lang, "transport_share") if self._share_usable() else translate(lang, "transport_upload")
        return [("log", translate(lang, "remote_transport", mode=mode))]

    def blocked_reason(self, src_path, output_format, lang):
        # The server owns the HWP FileHeader preflight; it reports a blocked
        # file as a normal failed item with the same localized message.
        return None

    def convert(self, job) -> JobResult:
        lang = job.lang
        if job.output_format not in OUTPUT_FORMATS:
            return JobResult(ok=False, message=translate(lang, "remote_format_unsupported",
                                                         format=job.output_format))

        item_id = f"{job.index:05d}-{job.output_format}"
        share_ref = self._share_ref(job)

        try:
            if share_ref is None:
                self._upload(job, item_id)

            payload = {
                "name": job.src_path.name,
                "output_format": job.output_format,
                "force_one_page": job.force_one_page,
            }
            if share_ref is not None:
                payload.update(share_ref)

            self._json("POST", protocol.run_path(self.job_id, item_id),
                       payload=payload, timeout=JOB_TIMEOUT)

            outcome = self._await_item(item_id, lang)
            notices = list(outcome.get("notices") or [])

            if outcome.get("status") != protocol.ITEM_OK:
                return JobResult(ok=False, message=outcome.get("message") or "", notices=notices)

            if share_ref is None:
                self._download(item_id, job.save_path, lang)
            elif not self._wait_for_share_output(job.save_path):
                return JobResult(ok=False, notices=notices,
                                 message=translate(lang, "remote_output_missing",
                                                   path=job.save_path))

            return JobResult(ok=True, actual_format=outcome.get("actual") or job.output_format,
                             notices=notices)

        except BackendUnavailable:
            raise
        except RemoteError as e:
            return JobResult(ok=False, message=str(e))

    def cancel(self) -> None:
        self.cancelled = True
        if not self.job_id:
            return
        try:
            self._json("POST", protocol.cancel_path(self.job_id), payload={}, timeout=JOB_TIMEOUT)
        except Exception:
            pass

    def close_session(self) -> None:
        if not self.job_id:
            return
        try:
            self._json("DELETE", protocol.job_path(self.job_id), timeout=JOB_TIMEOUT)
        except Exception:
            pass
        self.job_id = None
        self.sink = None

    # -- transfer ---------------------------------------------------------
    def _upload(self, job, item_id):
        size = job.open_path.stat().st_size
        timeout = max(60, size / UPLOAD_RATE_BYTES_PER_SECOND)
        attempt = 0
        while True:
            try:
                with open(job.open_path, "rb") as f:
                    with self._open(
                        "PUT",
                        protocol.input_path(self.job_id, item_id),
                        data=f,
                        timeout=timeout,
                        headers={
                            "Content-Type": "application/octet-stream",
                            "Content-Length": str(size),
                        },
                    ):
                        return
            except urllib.error.HTTPError as e:
                raise self._http_error(e) from None
            except (urllib.error.URLError, OSError) as e:
                # The PUT is idempotent -- the server writes to a .part file and
                # renames -- so retrying cannot leave a half-written input.
                if attempt >= len(RETRY_DELAYS) - 1:
                    raise RemoteError(translate(job.lang, "remote_upload_failed",
                                                name=job.src_path.name, detail=e)) from None
                time.sleep(RETRY_DELAYS[attempt])
                attempt += 1

    def _download(self, item_id, save_path: Path, lang):
        tmp = save_path.with_name(save_path.name + ".part")
        attempt = 0
        while True:
            try:
                with self._open("GET", protocol.output_path(self.job_id, item_id),
                                timeout=DOWNLOAD_TIMEOUT) as response:
                    save_path.parent.mkdir(parents=True, exist_ok=True)
                    with open(tmp, "wb") as f:
                        while True:
                            chunk = response.read(COPY_CHUNK)
                            if not chunk:
                                break
                            f.write(chunk)
                os.replace(tmp, save_path)
                return
            except urllib.error.HTTPError as e:
                tmp.unlink(missing_ok=True)
                raise self._http_error(e) from None
            except (urllib.error.URLError, OSError) as e:
                tmp.unlink(missing_ok=True)
                if attempt >= len(RETRY_DELAYS) - 1:
                    raise RemoteError(translate(lang, "remote_download_failed", detail=e)) from None
                time.sleep(RETRY_DELAYS[attempt])
                attempt += 1

    def _await_item(self, item_id, lang):
        """Drain the server event log until this item reports a result."""
        while True:
            page = self._json(
                "GET", protocol.events_path(self.job_id, self.cursor, EVENT_WAIT),
                timeout=EVENT_TIMEOUT, retries=len(RETRY_DELAYS),
            )
            self.cursor = page.get("cursor", self.cursor)
            for event in page.get("events") or []:
                kind = event.get("kind")
                if kind == protocol.EVENT_LOG and self.sink is not None:
                    self.sink.put((
                        "log",
                        (translate(lang, "server_prefix", text=event.get("text", "")),
                         event.get("level", "info")),
                    ))
                elif kind == protocol.EVENT_ITEM and event.get("item") == item_id:
                    return event
            if page.get("cancelled") and self.cancelled:
                return {"status": protocol.ITEM_FAILED, "message": translate(lang, "stopped")}

    # -- share transport --------------------------------------------------
    def _share_usable(self) -> bool:
        if self.transport == protocol.TRANSPORT_UPLOAD:
            return False
        return bool(self.shares) and bool(self.server_shares)

    def _share_ref(self, job):
        """Map a local path onto a server share, or None to upload instead."""
        if not self._share_usable():
            return None
        for share in self.shares:
            name = share.get("name")
            mount = share.get("local_mount")
            if not name or not mount or name not in self.server_shares:
                continue
            mount_path = Path(mount)
            try:
                rel = job.open_path.resolve().relative_to(mount_path.resolve())
                out_rel = job.save_path.resolve().relative_to(mount_path.resolve())
            except (ValueError, OSError):
                continue
            return {"share": name, "rel": rel.as_posix(), "out_rel": out_rel.as_posix()}
        return None

    def _wait_for_share_output(self, save_path: Path) -> bool:
        """SMB clients can lag behind the server's write by a moment."""
        deadline = time.monotonic() + SHARE_OUTPUT_GRACE_SECONDS
        while time.monotonic() < deadline:
            if save_path.exists() and save_path.stat().st_size > 0:
                return True
            time.sleep(0.2)
        return save_path.exists()
