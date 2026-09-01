"""Test doubles for the conversion backend and the event sink."""

from hwp2pdf.backends.base import BackendCapabilities, BackendUnavailable, JobResult

PDF_STUB = b"%PDF-1.4 fake\n"
DOCX_STUB = b"PK\x03\x04 fake\n"


class RecordingSink:
    """Stands in for the Tk log queue / ``CliEventSink``."""

    def __init__(self):
        self.events = []

    def put(self, item):
        self.events.append(item)

    def kinds(self):
        return [kind for kind, _payload in self.events]

    def of_kind(self, kind):
        return [payload for k, payload in self.events if k == kind]

    def logs(self):
        texts = []
        for payload in self.of_kind("log"):
            texts.append(payload[0] if isinstance(payload, tuple) else payload)
        return texts

    def done(self):
        payloads = self.of_kind("done")
        return payloads[-1] if payloads else None


class FakeBackend:
    """Writes stub output instead of driving Hancom Office."""

    capabilities = BackendCapabilities(
        name="fake",
        remote=False,
        local_staging=True,
        manages_hwp_process=False,
        local_preflight=True,
    )

    def __init__(self, fail_on=(), blocked=None, unavailable=None, open_unavailable=None,
                 on_convert=None):
        self.fail_on = set(fail_on)
        self.blocked = dict(blocked or {})
        self.unavailable = unavailable
        self.open_unavailable = open_unavailable
        self.on_convert = on_convert
        self.converted = []
        self.sessions_opened = 0
        self.sessions_closed = 0
        self.cancels = 0

    def preflight(self, lang):
        if self.unavailable:
            if isinstance(self.unavailable, BackendUnavailable):
                raise self.unavailable
            raise BackendUnavailable(self.unavailable)

    def open_session(self, sink, lang, options):
        self.sessions_opened += 1
        if self.open_unavailable:
            if isinstance(self.open_unavailable, BackendUnavailable):
                raise self.open_unavailable
            raise BackendUnavailable(self.open_unavailable)
        sink.put(("log", "fake session started"))

    def session_notes(self, lang):
        return [("log", "fake session note")]

    def blocked_reason(self, src_path, output_format, lang):
        return self.blocked.get(src_path.name)

    def convert(self, job):
        self.converted.append((job.src_path.name, job.output_format))
        if self.on_convert is not None:
            self.on_convert(job)
        if job.src_path.name in self.fail_on:
            return JobResult(ok=False, message=f"fake failure: {job.src_path.name}")
        stub = PDF_STUB if job.output_format == "PDF" else DOCX_STUB
        job.save_path.parent.mkdir(parents=True, exist_ok=True)
        job.save_path.write_bytes(stub)
        return JobResult(ok=True, actual_format=job.output_format)

    def cancel(self):
        self.cancels += 1

    def close_session(self):
        self.sessions_closed += 1
