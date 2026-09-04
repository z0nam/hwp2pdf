"""One TLS context for everything this app fetches over HTTPS.

Python's ``ssl`` module reads its trusted certificates from a path OpenSSL
recorded when that interpreter was *built*. That is fine while the interpreter
stays on the machine that built it, and wrong the moment PyInstaller freezes it
and ships it somewhere else: the macOS release inherited the CI runner's
Homebrew path, which exists on no user's Mac, so every HTTPS request failed
with ``CERTIFICATE_VERIFY_FAILED`` -- update checks, and the update download
itself.

A local build never showed it, because the developer's Mac really does have
that path. Only an installed CI artifact reproduces it.

``certifi`` carries its own copy of the CA bundle and PyInstaller collects it,
so it travels with the app. It is a fallback rather than the default: Windows
reads the system store and Linux has ``/etc/ssl/certs``, and a platform that
manages its own trust should keep managing it -- an administrator who adds a
corporate CA expects the app to honour it.
"""

import ssl
import urllib.request

_context = None


def ssl_context() -> ssl.SSLContext:
    """A verifying context that still works inside a frozen bundle."""
    global _context
    if _context is None:
        _context = ssl.create_default_context()
        if not _context.cert_store_stats().get("x509_ca"):
            # Nothing was loaded: the compiled-in path is absent here.
            try:
                import certifi

                _context.load_verify_locations(certifi.where())
            except Exception:
                # Leave the empty context rather than disabling verification;
                # a clear TLS failure beats a silently unauthenticated one.
                pass
    return _context


def urlopen(request, timeout):
    """``urllib.request.urlopen`` with this app's trust settings.

    Harmless for plain HTTP, where the context is simply unused, so callers do
    not have to know which scheme they were handed.
    """
    return urllib.request.urlopen(request, timeout=timeout, context=ssl_context())
