"""Server-side job state and the single conversion worker.

The Hancom COM engine is a per-process, per-thread singleton, so exactly one
worker thread owns it. Everything else -- uploads, event polling, downloads --
runs on the HTTP handler threads and only touches the job's lock.
"""

import shutil
import threading
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from queue import Empty, Queue

from hwp2pdf import paths
from hwp2pdf.backends.base import JobSpec, SessionOptions
from hwp2pdf.constants import output_extension
from hwp2pdf.i18n import translate
from hwp2pdf.server import protocol


class QueueFull(Exception):
    """Raised when more work is queued than ``--max-queue`` allows."""


def _safe_name(name: str) -> str:
    """Reduce a client-supplied file name to a leaf with no path separators."""
    leaf = Path(str(name).replace("\\", "/")).name
    return "".join(ch for ch in leaf if ch not in '<>:"|?*' and ord(ch) >= 32).strip()


@dataclass
class Item:
    item_id: str
    name: str
    output_format: str
    force_one_page: bool
    #: Set for share transport; the server resolves these against a share root.
    share: str = ""
    rel: str = ""
    out_rel: str = ""
    state: str = "pending"
    #: Resolved once the worker picks the item up; the download route reads it.
    target_path: Path | None = None


@dataclass
class Job:
    job_id: str
    lang: str
    safe_temp: bool
    workdir: Path
    created_at: float = field(default_factory=time.monotonic)
    touched_at: float = field(default_factory=time.monotonic)
    cancelled: bool = False
    session_open: bool = False
    items: dict = field(default_factory=dict)
    events: list = field(default_factory=list)
    lock: threading.Lock = field(default_factory=threading.Lock)
    changed: threading.Condition = None
    backend: object = None

    def __post_init__(self):
        if self.changed is None:
            self.changed = threading.Condition(self.lock)

    def touch(self):
        self.touched_at = time.monotonic()

    def emit(self, event: dict):
        with self.lock:
            event = dict(event)
            event["seq"] = len(self.events) + 1
            self.events.append(event)
            self.changed.notify_all()

    def events_since(self, cursor: int, wait: float):
        """Return events after ``cursor``, long-polling up to ``wait`` seconds."""
        deadline = time.monotonic() + max(0.0, wait)
        with self.lock:
            while True:
                pending = self.events[cursor:]
                if pending or self.cancelled:
                    return list(pending), len(self.events)
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    return [], len(self.events)
                self.changed.wait(remaining)

    def input_path(self, item_id: str) -> Path:
        """Where an upload lands before its original name is known."""
        return self.workdir / f"{item_id}.in"

    def item_dir(self, item_id: str) -> Path:
        return self.workdir / item_id

    def staged_paths(self, item: "Item"):
        """Source and target inside the item's own folder.

        The upload is re-homed under its original file name so Hangul, the
        FileHeader preflight and any error message all see a sane path.
        """
        name = _safe_name(item.name) or item.item_id
        directory = self.item_dir(item.item_id)
        source = directory / name
        target = source.with_suffix(output_extension(item.output_format))
        return source, target


class JobStore:
    """Job registry plus the single COM worker thread."""

    def __init__(
        self,
        backend_factory,
        root: Path | None = None,
        share_roots: dict | None = None,
        max_queue: int = protocol.DEFAULT_MAX_QUEUE,
        job_ttl: float = protocol.DEFAULT_JOB_TTL_SECONDS,
    ):
        self.backend_factory = backend_factory
        self.root = Path(root) if root else paths.server_state_dir() / "jobs"
        self.share_roots = {k: Path(v).resolve() for k, v in (share_roots or {}).items()}
        self.max_queue = max_queue
        self.job_ttl = job_ttl
        self.jobs = {}
        #: The one job whose COM session is currently open, if any. Hangul is a
        #: process/thread singleton, so two live sessions can never coexist.
        self.active_job = None
        self.lock = threading.Lock()
        self.queue = Queue()
        self.stop_event = threading.Event()
        self.worker = None
        self._reap_stale_workdirs()

    # -- lifecycle -------------------------------------------------------
    def start(self):
        if self.worker is None:
            self.worker = threading.Thread(target=self._run, name="hwp2pdf-worker", daemon=True)
            self.worker.start()

    def shutdown(self, timeout: float = 10.0):
        """Stop the worker, closing the live COM session on its own thread.

        Order matters: the session must be closed by the worker (it owns the COM
        apartment) before the worker exits, or Hwp.exe is left running.
        """
        self.stop_event.set()
        self.queue.put(None)
        if self.worker is not None:
            self.worker.join(timeout)
            if self.worker.is_alive():
                # The worker is wedged inside Hangul; fall back to closing here
                # so a stuck conversion cannot leak the process indefinitely.
                self._close_active_session()
            self.worker = None
        else:
            self._close_active_session()
        for job_id in list(self.jobs):
            self.delete_job(job_id)

    # -- job API ---------------------------------------------------------
    def create_job(self, lang: str, safe_temp: bool) -> Job:
        job_id = uuid.uuid4().hex
        workdir = self.root / job_id
        workdir.mkdir(parents=True, exist_ok=True)
        job = Job(job_id=job_id, lang=lang, safe_temp=safe_temp, workdir=workdir)
        with self.lock:
            self.jobs[job_id] = job
        self.reap_expired()
        return job

    def get(self, job_id: str):
        with self.lock:
            job = self.jobs.get(job_id)
        if job is not None:
            job.touch()
        return job

    def delete_job(self, job_id: str):
        with self.lock:
            job = self.jobs.pop(job_id, None)
        if job is None:
            return False
        job.cancelled = True
        with job.lock:
            job.changed.notify_all()
        worker_alive = self.worker is not None and self.worker.is_alive()
        if job.session_open and worker_alive:
            # Closing must happen on the worker thread that owns the COM apartment.
            self.queue.put(("close", job))
        else:
            # No worker to run it: close here rather than leak the engine.
            if job.session_open:
                self._close_session(job)
            shutil.rmtree(job.workdir, ignore_errors=True)
        return True

    def cancel_job(self, job_id: str):
        job = self.get(job_id)
        if job is None:
            return False
        job.cancelled = True
        with job.lock:
            job.changed.notify_all()
        return True

    def queue_depth(self) -> int:
        return self.queue.qsize()

    def submit(self, job: Job, item: Item):
        if self.queue.qsize() >= self.max_queue:
            raise QueueFull()
        with job.lock:
            job.items[item.item_id] = item
        job.touch()
        self.queue.put(("convert", job, item))

    def reap_expired(self):
        now = time.monotonic()
        with self.lock:
            expired = [j for j, job in self.jobs.items() if now - job.touched_at > self.job_ttl]
        for job_id in expired:
            self.delete_job(job_id)

    def _reap_stale_workdirs(self):
        """Job workdirs never outlive the process that created them."""
        try:
            if self.root.is_dir():
                shutil.rmtree(self.root, ignore_errors=True)
            self.root.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass

    # -- share resolution -------------------------------------------------
    def resolve_share(self, share: str, rel: str) -> Path:
        root = self.share_roots.get(share)
        if root is None:
            raise ValueError(f"unknown share: {share}")
        if not rel:
            raise ValueError("empty share path")
        candidate = (root / rel).resolve()
        if candidate != root and not candidate.is_relative_to(root):
            raise ValueError("share path escapes its root")
        return candidate

    # -- worker ------------------------------------------------------------
    def _run(self):
        try:
            while not self.stop_event.is_set():
                try:
                    task = self.queue.get(timeout=0.5)
                except Empty:
                    self.reap_expired()
                    continue
                if task is None:
                    break
                try:
                    if task[0] == "close":
                        self._close_session(task[1])
                    else:
                        self._convert(task[1], task[2])
                except Exception as e:  # never let the worker die
                    _job = task[1]
                    _job.emit({"kind": protocol.EVENT_LOG, "text": f"worker error: {e}", "level": "error"})
        finally:
            # Whatever happens, the engine goes down with the thread that owns it.
            self._close_active_session()

    def _sink_for(self, job: Job):
        class _Sink:
            @staticmethod
            def put(item):
                kind, payload = item
                if kind != "log":
                    return
                if isinstance(payload, tuple):
                    text, level = payload
                else:
                    text, level = payload, "info"
                job.emit({"kind": protocol.EVENT_LOG, "text": str(text), "level": level})

        return _Sink()

    def _close_active_session(self):
        active = self.active_job
        if active is not None:
            self._close_session(active)

    def _ensure_session(self, job: Job):
        if job.session_open:
            return
        # Only one engine session may exist; retire the previous job's first.
        if self.active_job is not None and self.active_job is not job:
            self._close_session(self.active_job)
        job.backend = self.backend_factory()
        job.backend.preflight(job.lang)
        job.backend.open_session(
            self._sink_for(job),
            job.lang,
            SessionOptions(
                lang=job.lang,
                output_formats=(),
                force_one_page=False,
                safe_temp=job.safe_temp,
                total_files=0,
            ),
        )
        for note in job.backend.session_notes(job.lang):
            self._sink_for(job).put(note)
        job.session_open = True
        self.active_job = job
        job.emit({"kind": protocol.EVENT_SESSION, "state": "started"})

    def _close_session(self, job: Job):
        if job.session_open and job.backend is not None:
            try:
                job.backend.close_session()
            except Exception:
                pass
            job.emit({"kind": protocol.EVENT_SESSION, "state": "closed"})
        job.session_open = False
        job.backend = None
        if self.active_job is job:
            self.active_job = None
        shutil.rmtree(job.workdir, ignore_errors=True)

    def _convert(self, job: Job, item: Item):
        if job.cancelled:
            item.state = protocol.ITEM_FAILED
            job.emit({
                "kind": protocol.EVENT_ITEM, "item": item.item_id,
                "status": protocol.ITEM_FAILED,
                "message": translate(job.lang, "stopped"), "actual": "", "notices": [],
            })
            return

        try:
            self._ensure_session(job)
        except Exception as e:
            item.state = protocol.ITEM_FAILED
            job.emit({
                "kind": protocol.EVENT_ITEM, "item": item.item_id,
                "status": protocol.ITEM_FAILED, "message": str(e), "actual": "", "notices": [],
            })
            return

        if item.share:
            source = self.resolve_share(item.share, item.rel)
            target = self.resolve_share(item.share, item.out_rel)
        else:
            upload = job.input_path(item.item_id)
            source, target = job.staged_paths(item)
            source.parent.mkdir(parents=True, exist_ok=True)
            if upload.exists():
                upload.replace(source)

        blocked = job.backend.blocked_reason(source, item.output_format, job.lang)
        if blocked:
            item.state = protocol.ITEM_BLOCKED
            job.emit({
                "kind": protocol.EVENT_ITEM, "item": item.item_id,
                "status": protocol.ITEM_BLOCKED, "message": blocked, "actual": "", "notices": [],
            })
            return

        result = job.backend.convert(
            JobSpec(
                index=1,
                # ``src_path`` only supplies the source extension for messages.
                src_path=Path(item.name),
                open_path=source,
                save_path=target,
                output_format=item.output_format,
                force_one_page=item.force_one_page,
                safe_temp=job.safe_temp,
                lang=job.lang,
            )
        )

        if result.ok and not item.share and not target.exists():
            result.ok = False
            result.message = translate(job.lang, "temp_missing", format=item.output_format)

        item.target_path = target
        item.state = protocol.ITEM_OK if result.ok else protocol.ITEM_FAILED
        job.emit({
            "kind": protocol.EVENT_ITEM,
            "item": item.item_id,
            "status": item.state,
            "actual": result.actual_format,
            "message": result.message,
            "notices": list(result.notices),
        })
