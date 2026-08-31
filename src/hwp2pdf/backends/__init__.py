"""Backend selection.

Windows keeps using the local COM engine by default. Anywhere else -- and on
Windows when a server address is configured -- conversion is delegated to a
Windows conversion server over HTTP.
"""

import os

from hwp2pdf import config
from hwp2pdf.backends.base import (
    BackendCapabilities,
    BackendUnavailable,
    ConversionBackend,
    JobResult,
    JobSpec,
    SessionOptions,
)
from hwp2pdf.i18n import translate

__all__ = [
    "BackendCapabilities",
    "BackendUnavailable",
    "ConversionBackend",
    "JobResult",
    "JobSpec",
    "SessionOptions",
    "create_backend",
]


def create_backend(server=None, lang: str = "ko", rhwp_fallback: bool = False,
                   rhwp_path: str = ""):
    """Return the backend to use for a batch.

    ``server`` is a mapping like ``config.server_settings()``. When it is None
    the saved settings and the ``HWP2PDF_SERVER_URL`` / ``HWP2PDF_TOKEN``
    environment variables are consulted.

    ``rhwp_fallback`` wraps the result so that an unreachable conversion server
    falls back to local rhwp rendering, which is PDF-only and approximate.
    """
    resolved = server if server is not None else config.server_settings()
    url = (resolved or {}).get("url", "").strip()

    primary = None
    if url:
        from hwp2pdf.backends.remote_http import RemoteHttpBackend

        primary = RemoteHttpBackend(resolved)
    elif os.name == "nt":
        from hwp2pdf.backends.windows_com import WindowsComBackend

        primary = WindowsComBackend()

    if primary is None and not rhwp_fallback:
        raise BackendUnavailable(translate(lang, "no_backend"))

    if not rhwp_fallback:
        return primary

    from hwp2pdf.backends.local_rhwp import RhwpBackend

    rhwp = RhwpBackend(rhwp_path)
    if primary is None:
        return rhwp

    from hwp2pdf.backends.fallback import FallbackBackend

    return FallbackBackend(primary, rhwp)
