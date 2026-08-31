"""Try one backend, fall back to another when it cannot start.

Used for "convert locally with rhwp when the conversion server is unreachable".
The fallback only engages on ``BackendUnavailable`` -- a backend that starts and
then fails a file is not a reason to switch engines mid-batch.
"""

from hwp2pdf.backends.base import BackendUnavailable
from hwp2pdf.i18n import translate


class FallbackBackend:
    """Delegates to ``primary``, or to ``fallback`` if the primary cannot start."""

    def __init__(self, primary, fallback):
        self.primary = primary
        self.fallback = fallback
        self.active = primary
        self.primary_error = ""

    @property
    def capabilities(self):
        return self.active.capabilities

    def preflight(self, lang: str) -> None:
        try:
            self.primary.preflight(lang)
            self.active = self.primary
            return
        except BackendUnavailable as e:
            self.primary_error = str(e)

        try:
            self.fallback.preflight(lang)
        except BackendUnavailable:
            # Report the primary's problem: that is the one the user meant to use.
            raise BackendUnavailable(self.primary_error) from None
        self.active = self.fallback

    def open_session(self, sink, lang: str, options) -> None:
        if self.active is self.fallback and self.primary_error:
            first_line = self.primary_error.splitlines()[0]
            sink.put(("log", (translate(lang, "fallback_engaged", detail=first_line), "warning")))
        self.active.open_session(sink, lang, options)

    def session_notes(self, lang: str) -> list:
        return self.active.session_notes(lang)

    def blocked_reason(self, src_path, output_format, lang):
        return self.active.blocked_reason(src_path, output_format, lang)

    def convert(self, job):
        return self.active.convert(job)

    def cancel(self) -> None:
        self.active.cancel()

    def close_session(self) -> None:
        self.active.close_session()
