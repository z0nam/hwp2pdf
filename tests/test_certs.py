"""Trusting the right certificates from inside a frozen app.

The bug these cover shipped: the macOS release inherited the CI runner's
OpenSSL cert path, which exists on no user's Mac, so every HTTPS request failed
and the update button could neither check nor download. A local build never
reproduces it -- the developer's machine really does have that path -- so the
fallback is exercised here by emptying the store instead.
"""

import ssl

import pytest

from hwp2pdf import certs


@pytest.fixture(autouse=True)
def fresh_context():
    """The context is cached for the process; each test gets its own."""
    certs._context = None
    yield
    certs._context = None


def test_a_platform_that_has_its_own_trust_keeps_using_it(monkeypatch):
    # Windows reads the system store and Linux has /etc/ssl/certs. An admin who
    # added a corporate CA there expects the app to honour it.
    loaded = []
    real = ssl.create_default_context

    def context_with_cas():
        ctx = real()
        monkeypatch.setattr(ctx, "load_verify_locations", lambda *a, **k: loaded.append(a))
        monkeypatch.setattr(ctx, "cert_store_stats", lambda: {"x509_ca": 193})
        return ctx

    monkeypatch.setattr(certs.ssl, "create_default_context", context_with_cas)
    certs.ssl_context()
    assert loaded == []


def test_an_empty_store_falls_back_to_the_bundled_certificates(monkeypatch):
    loaded = []
    real = ssl.create_default_context

    def empty_context():
        ctx = real()
        monkeypatch.setattr(ctx, "load_verify_locations", lambda p: loaded.append(p))
        monkeypatch.setattr(ctx, "cert_store_stats", lambda: {"x509_ca": 0})
        return ctx

    monkeypatch.setattr(certs.ssl, "create_default_context", empty_context)
    certs.ssl_context()

    assert len(loaded) == 1
    assert loaded[0].endswith(".pem")


def test_the_bundle_ships_with_the_app():
    # If this import fails the frozen build has nothing to fall back to.
    import certifi
    from pathlib import Path

    assert Path(certifi.where()).is_file()


def test_verification_is_never_silently_switched_off(monkeypatch):
    """A failed fallback must still refuse a bad certificate."""
    real = ssl.create_default_context

    def empty_context():
        ctx = real()
        monkeypatch.setattr(ctx, "cert_store_stats", lambda: {"x509_ca": 0})
        monkeypatch.setattr(
            ctx, "load_verify_locations",
            lambda *a, **k: (_ for _ in ()).throw(OSError("no bundle")),
        )
        return ctx

    monkeypatch.setattr(certs.ssl, "create_default_context", empty_context)
    context = certs.ssl_context()

    assert context.verify_mode == ssl.CERT_REQUIRED
    assert context.check_hostname is True


def test_the_context_is_built_once(monkeypatch):
    calls = []
    real = ssl.create_default_context
    monkeypatch.setattr(certs.ssl, "create_default_context",
                        lambda: calls.append(1) or real())

    first = certs.ssl_context()
    second = certs.ssl_context()
    assert first is second
    assert len(calls) == 1
