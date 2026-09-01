"""Try the preferred Hancom backend, then an explicitly enabled fallback.

The switch may happen during preflight or while opening the conversion session,
but never after file conversion begins.  That keeps one batch from silently
mixing exact Hancom output with approximate rhwp output.
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
            if not e.fallback_allowed:
                raise
            self._activate_fallback(e, lang)

    def open_session(self, sink, lang: str, options) -> None:
        if self.active is self.primary:
            try:
                self.primary.open_session(sink, lang, options)
                return
            except BackendUnavailable as e:
                if not e.fallback_allowed:
                    raise
                try:
                    self.primary.close_session()
                except Exception:
                    pass
                self._activate_fallback(e, lang)

        if self.active is self.fallback and self.primary_error:
            first_line = self.primary_error.splitlines()[0]
            sink.put(("log", (translate(lang, "fallback_engaged", detail=first_line), "warning")))
        self.active.open_session(sink, lang, options)

    def _activate_fallback(self, primary_error: BackendUnavailable, lang: str) -> None:
        self.primary_error = str(primary_error)
        try:
            self.fallback.preflight(lang)
        except BackendUnavailable:
            # Report the preferred engine's problem: that is the one the user
            # expected to use, and the rhwp installation status is shown in UI.
            raise BackendUnavailable(self.primary_error) from None
        self.active = self.fallback

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
