"""A busy port must fail loudly -- the windowless build has no other symptom."""

import socket
import threading

import pytest

from fakes import FakeBackend

from hwp2pdf.server import protocol
from hwp2pdf.server.http_server import PortUnavailable, create_server

PROBE = {"installed": True, "detail": "fake", "running": []}


def test_default_port_avoids_crowded_and_ephemeral_ranges():
    # 8765/8766 were already taken by unrelated local services.
    assert protocol.DEFAULT_PORT not in (8000, 8080, 8765, 8766, 8888, 9000)
    # Windows hands out 49152+ as ephemeral ports; a default there binds randomly.
    assert 1024 < protocol.DEFAULT_PORT < 49152


def test_busy_port_raises_a_useful_error(tmp_path):
    blocker = socket.socket()
    blocker.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    blocker.bind(("127.0.0.1", 0))
    blocker.listen(1)
    port = blocker.getsockname()[1]
    try:
        with pytest.raises(PortUnavailable) as excinfo:
            create_server(
                "127.0.0.1", port,
                backend_factory=FakeBackend, hwp_probe=lambda: PROBE,
                token="", quiet=True, bind_retries=1, bind_retry_delay=0,
            )
        message = str(excinfo.value)
        assert str(port) in message
        assert "--port" in message  # tells the user how to fix it
    finally:
        blocker.close()


def test_a_transient_conflict_is_retried(tmp_path):
    """A previous instance shutting down should not be a fatal conflict."""
    blocker = socket.socket()
    blocker.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    blocker.bind(("127.0.0.1", 0))
    blocker.listen(1)
    port = blocker.getsockname()[1]

    # Release the port shortly after the first bind attempt fails.
    threading.Timer(0.15, blocker.close).start()

    notes = []
    httpd = create_server(
        "127.0.0.1", port,
        backend_factory=FakeBackend, hwp_probe=lambda: PROBE,
        token="", quiet=True, log_sink=notes.append,
        bind_retries=10, bind_retry_delay=0.1,
    )
    try:
        assert httpd.server_address[1] == port
        assert any("busy" in n for n in notes)
    finally:
        # serve_forever() was never started here, and BaseServer.shutdown()
        # blocks forever waiting for a loop that does not exist.
        httpd.store.shutdown()
        httpd.server_close()


def test_the_job_store_is_not_left_running_after_a_failed_bind():
    blocker = socket.socket()
    blocker.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    blocker.bind(("127.0.0.1", 0))
    blocker.listen(1)
    port = blocker.getsockname()[1]
    before = threading.active_count()
    try:
        with pytest.raises(PortUnavailable):
            create_server(
                "127.0.0.1", port,
                backend_factory=FakeBackend, hwp_probe=lambda: PROBE,
                token="", quiet=True, bind_retries=0,
            )
    finally:
        blocker.close()
    # The worker thread must not outlive the failed startup.
    for _ in range(50):
        if threading.active_count() <= before:
            break
        threading.Event().wait(0.05)
    assert threading.active_count() <= before + 1


def test_windows_does_not_allow_two_servers_on_one_port():
    """SO_REUSEADDR means something different on Windows.

    There it permits binding a port that is already listening, so a second
    instance would silently share the port instead of failing to start.
    """
    import os

    from hwp2pdf.server.http_server import ConversionHTTPServer

    assert ConversionHTTPServer.allow_reuse_address is (os.name != "nt")
