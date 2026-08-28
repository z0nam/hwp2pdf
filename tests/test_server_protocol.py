"""End-to-end protocol test: the real client against the real server.

Runs anywhere -- the server is wired to ``FakeBackend`` instead of the Hancom
COM engine, so this exercises routing, auth, the event log, transfers, queueing
and cancellation on a machine with no Hangul installed.
"""

import json
import threading
import urllib.error
import urllib.request

import pytest

from fakes import DOCX_STUB, PDF_STUB, FakeBackend, RecordingSink

from hwp2pdf import jobs
from hwp2pdf.backends.base import BackendUnavailable
from hwp2pdf.backends.remote_http import RemoteHttpBackend
from hwp2pdf.server import protocol
from hwp2pdf.server.http_server import create_server

TOKEN = "test-token"


class ServerHandle:
    def __init__(self, httpd, backends):
        self.httpd = httpd
        self.backends = backends

    @property
    def url(self):
        host, port = self.httpd.server_address[:2]
        return f"http://{host}:{port}"


@pytest.fixture
def server(request):
    created = []

    def factory():
        backend = FakeBackend(**getattr(request, "param", {}))
        created.append(backend)
        return backend

    httpd = create_server(
        "127.0.0.1", 0,
        backend_factory=factory,
        hwp_probe=lambda: {"installed": True, "detail": "fake", "running": []},
        token=TOKEN,
        quiet=True,
    )
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    try:
        yield ServerHandle(httpd, created)
    finally:
        httpd.shutdown()
        httpd.store.shutdown()
        httpd.server_close()


def client(server, **overrides):
    settings = {"url": server.url, "token": TOKEN, "transport": "auto", "shares": []}
    settings.update(overrides)
    return RemoteHttpBackend(settings)


def get(server, path, token=TOKEN):
    request = urllib.request.Request(server.url + path)
    if token:
        request.add_header(protocol.AUTH_HEADER, f"{protocol.AUTH_SCHEME} {token}")
    with urllib.request.urlopen(request, timeout=10) as response:
        return response.status, json.loads(response.read().decode("utf-8"))


def run_remote(tmp_path, server, **overrides):
    sink = RecordingSink()
    options = dict(
        target=str(tmp_path),
        recursive=True,
        overwrite=True,
        use_safe_copy=True,
        force_one_page=True,
        output_formats=("PDF",),
        lang="ko",
    )
    options.update(overrides)
    jobs.run_batch(sink, client(server), **options)
    return sink


def make_files(tmp_path, *names):
    for name in names:
        (tmp_path / name).write_bytes(b"fake hwp bytes")
    return tmp_path


# -- basics --------------------------------------------------------------
def test_health_needs_no_token(server):
    status, payload = get(server, protocol.PATH_HEALTH, token=None)
    assert status == 200
    assert payload["api"] == protocol.API_VERSION
    assert payload["auth_required"] is True


def test_capabilities_requires_a_token(server):
    with pytest.raises(urllib.error.HTTPError) as excinfo:
        get(server, protocol.PATH_CAPABILITIES, token=None)
    assert excinfo.value.code == 401

    with pytest.raises(urllib.error.HTTPError) as excinfo:
        get(server, protocol.PATH_CAPABILITIES, token="wrong")
    assert excinfo.value.code == 401

    status, payload = get(server, protocol.PATH_CAPABILITIES)
    assert status == 200
    assert payload["hwp_installed"] is True
    assert "PDF" in payload["formats"] and "DOCX" in payload["formats"]


def test_unknown_job_is_404(server):
    with pytest.raises(urllib.error.HTTPError) as excinfo:
        get(server, protocol.job_path("0" * 32) + "/events?cursor=0&wait=0")
    assert excinfo.value.code == 404


# -- full round trip -----------------------------------------------------
def test_upload_convert_download_round_trip(tmp_path, server):
    make_files(tmp_path, "a.hwp", "b.hwpx")

    sink = run_remote(tmp_path, server)

    assert (tmp_path / "a.pdf").read_bytes() == PDF_STUB
    assert (tmp_path / "b.pdf").read_bytes() == PDF_STUB
    assert sink.done()[:3] == (2, 0, 0)
    assert not (tmp_path / jobs.LOG_CSV_NAME).exists()
    # The source keeps its extension on the server so the HWP preflight works.
    assert sorted(name for name, _fmt in server.backends[0].converted) == ["a.hwp", "b.hwpx"]


def test_both_formats_round_trip(tmp_path, server):
    make_files(tmp_path, "a.hwp")

    run_remote(tmp_path, server, output_formats=("PDF", "DOCX"))

    assert (tmp_path / "a.pdf").read_bytes() == PDF_STUB
    assert (tmp_path / "a.docx").read_bytes() == DOCX_STUB


def test_server_log_lines_reach_the_client(tmp_path, server):
    make_files(tmp_path, "a.hwp")
    sink = run_remote(tmp_path, server)
    assert any("fake session started" in text for text in sink.logs())
    assert any(text.startswith("서버: ") for text in sink.logs())


@pytest.mark.parametrize("server", [{"fail_on": {"a.hwp"}}], indirect=True)
def test_remote_failure_is_reported_per_file(tmp_path, server):
    make_files(tmp_path, "a.hwp", "b.hwp")

    sink = run_remote(tmp_path, server)

    assert not (tmp_path / "a.pdf").exists()
    assert (tmp_path / "b.pdf").exists()
    assert sink.done()[:3] == (1, 1, 0)
    assert any("fake failure: a.hwp" in text for text in sink.logs())


@pytest.mark.parametrize("server", [{"blocked": {"a.hwp": "암호 문서"}}], indirect=True)
def test_server_side_preflight_message_survives_the_wire(tmp_path, server):
    make_files(tmp_path, "a.hwp")

    sink = run_remote(tmp_path, server)

    assert any("암호 문서" in text for text in sink.logs())
    assert sink.done()[:3] == (0, 1, 0)


def test_session_is_closed_when_the_batch_ends(tmp_path, server):
    make_files(tmp_path, "a.hwp")
    run_remote(tmp_path, server)

    backend = server.backends[0]
    deadline = threading.Event()
    for _ in range(50):
        if backend.sessions_closed == 1:
            break
        deadline.wait(0.05)
    assert backend.sessions_opened == 1
    assert backend.sessions_closed == 1


# -- failure modes -------------------------------------------------------
def test_bad_token_surfaces_as_backend_unavailable(tmp_path, server):
    backend = client(server, token="nope")
    with pytest.raises(BackendUnavailable):
        backend.preflight("ko")


def test_missing_url_surfaces_as_backend_unavailable():
    with pytest.raises(BackendUnavailable):
        RemoteHttpBackend({"url": "", "token": ""}).preflight("ko")


def test_unreachable_server_surfaces_as_backend_unavailable(monkeypatch):
    monkeypatch.setattr("hwp2pdf.backends.remote_http.RETRY_DELAYS", (0, 0, 0))
    backend = RemoteHttpBackend({"url": "http://127.0.0.1:9", "token": ""})
    with pytest.raises(BackendUnavailable) as excinfo:
        backend.preflight("ko")
    assert "127.0.0.1:9" in str(excinfo.value)


def test_server_without_hangul_is_refused(tmp_path):
    httpd = create_server(
        "127.0.0.1", 0,
        backend_factory=FakeBackend,
        hwp_probe=lambda: {"installed": False, "detail": "no hangul", "running": []},
        token="",
        quiet=True,
    )
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    try:
        host, port = httpd.server_address[:2]
        backend = RemoteHttpBackend({"url": f"http://{host}:{port}", "token": ""})
        with pytest.raises(BackendUnavailable) as excinfo:
            backend.preflight("ko")
        assert "no hangul" in str(excinfo.value)
    finally:
        httpd.shutdown()
        httpd.store.shutdown()
        httpd.server_close()


def test_upload_over_the_limit_is_rejected(tmp_path):
    httpd = create_server(
        "127.0.0.1", 0,
        backend_factory=FakeBackend,
        hwp_probe=lambda: {"installed": True, "detail": "fake", "running": []},
        token="",
        max_upload_bytes=4,
        quiet=True,
    )
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    try:
        host, port = httpd.server_address[:2]
        make_files(tmp_path, "a.hwp")
        sink = RecordingSink()
        jobs.run_batch(
            sink,
            RemoteHttpBackend({"url": f"http://{host}:{port}", "token": ""}),
            target=str(tmp_path), recursive=False, overwrite=True, use_safe_copy=False,
            force_one_page=True, output_formats=("PDF",), lang="ko",
        )
        assert sink.done()[:3] == (0, 1, 0)
        assert any("업로드 상한" in text for text in sink.logs())
    finally:
        httpd.shutdown()
        httpd.store.shutdown()
        httpd.server_close()


def test_run_without_upload_is_rejected(server):
    backend = client(server)
    backend.preflight("ko")
    backend.open_session(RecordingSink(), "ko", None)
    try:
        request = urllib.request.Request(
            server.url + protocol.run_path(backend.job_id, "00001-PDF"),
            data=json.dumps({"name": "a.hwp", "output_format": "PDF"}).encode("utf-8"),
            method="POST",
        )
        request.add_header(protocol.AUTH_HEADER, f"{protocol.AUTH_SCHEME} {TOKEN}")
        request.add_header("Content-Type", "application/json")
        with pytest.raises(urllib.error.HTTPError) as excinfo:
            urllib.request.urlopen(request, timeout=10)
        assert excinfo.value.code == 400
    finally:
        backend.close_session()


def test_unsupported_output_format_is_rejected(server):
    backend = client(server)
    backend.preflight("ko")
    backend.open_session(RecordingSink(), "ko", None)
    try:
        request = urllib.request.Request(
            server.url + protocol.run_path(backend.job_id, "00001-XLS"),
            data=json.dumps({"name": "a.hwp", "output_format": "XLS"}).encode("utf-8"),
            method="POST",
        )
        request.add_header(protocol.AUTH_HEADER, f"{protocol.AUTH_SCHEME} {TOKEN}")
        request.add_header("Content-Type", "application/json")
        with pytest.raises(urllib.error.HTTPError) as excinfo:
            urllib.request.urlopen(request, timeout=10)
        assert excinfo.value.code == 400
    finally:
        backend.close_session()


def test_deleting_a_job_makes_it_unknown(server):
    backend = client(server)
    backend.preflight("ko")
    backend.open_session(RecordingSink(), "ko", None)
    job_id = backend.job_id
    backend.close_session()

    with pytest.raises(urllib.error.HTTPError) as excinfo:
        get(server, protocol.job_path(job_id) + "/events?cursor=0&wait=0")
    assert excinfo.value.code == 404


# -- share transport -----------------------------------------------------
def test_share_transport_skips_upload_and_download(tmp_path):
    share_root = tmp_path / "share"
    (share_root / "docs").mkdir(parents=True)
    (share_root / "docs" / "a.hwp").write_bytes(b"fake hwp bytes")

    httpd = create_server(
        "127.0.0.1", 0,
        backend_factory=FakeBackend,
        hwp_probe=lambda: {"installed": True, "detail": "fake", "running": []},
        token="",
        share_roots={"work": str(share_root)},
        quiet=True,
    )
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    try:
        host, port = httpd.server_address[:2]
        sink = RecordingSink()
        jobs.run_batch(
            sink,
            RemoteHttpBackend({
                "url": f"http://{host}:{port}", "token": "", "transport": "share",
                "shares": [{"name": "work", "local_mount": str(share_root)}],
            }),
            target=str(share_root / "docs"), recursive=False, overwrite=True,
            use_safe_copy=True, force_one_page=True, output_formats=("PDF",), lang="ko",
        )
        assert (share_root / "docs" / "a.pdf").read_bytes() == PDF_STUB
        assert sink.done()[:3] == (1, 0, 0)
        assert any("공유 폴더" in text for text in sink.logs())
    finally:
        httpd.shutdown()
        httpd.store.shutdown()
        httpd.server_close()


def test_share_path_traversal_is_rejected(tmp_path):
    from hwp2pdf.server.jobs import JobStore

    share_root = tmp_path / "share"
    share_root.mkdir()
    store = JobStore(backend_factory=FakeBackend, root=tmp_path / "jobs",
                     share_roots={"work": str(share_root)})
    try:
        assert store.resolve_share("work", "a/b.hwp") == (share_root / "a" / "b.hwp").resolve()
        with pytest.raises(ValueError):
            store.resolve_share("work", "../outside.hwp")
        with pytest.raises(ValueError):
            store.resolve_share("nope", "a.hwp")
        with pytest.raises(ValueError):
            store.resolve_share("work", "")
    finally:
        store.shutdown()


# -- queueing and cancellation -------------------------------------------
def test_queue_full_returns_429(tmp_path):
    """The COM engine is serialized, so the queue has a hard ceiling."""
    from hwp2pdf.server.jobs import Item, JobStore, QueueFull

    store = JobStore(backend_factory=FakeBackend, root=tmp_path / "jobs", max_queue=1)
    try:
        job = store.create_job("ko", True)
        item = Item(item_id="1", name="a.hwp", output_format="PDF", force_one_page=True)
        # The worker is not started, so nothing drains the queue.
        store.queue.put(("noop", job, item))
        with pytest.raises(QueueFull):
            store.submit(job, item)
    finally:
        store.shutdown()


def test_cancel_marks_the_job_and_wakes_pollers(tmp_path):
    from hwp2pdf.server.jobs import JobStore

    store = JobStore(backend_factory=FakeBackend, root=tmp_path / "jobs")
    try:
        job = store.create_job("ko", True)
        assert store.cancel_job(job.job_id) is True
        assert job.cancelled is True
        # A cancelled job returns from a long poll immediately.
        events, cursor = job.events_since(0, 30)
        assert events == [] and cursor == 0
        assert store.cancel_job("0" * 32) is False
    finally:
        store.shutdown()


def test_client_cancel_stops_the_remote_batch(tmp_path, server):
    make_files(tmp_path, "a.hwp", "b.hwp", "c.hwp")
    state = {"stop": False}
    sink = RecordingSink()

    backend = client(server)
    original_convert = backend.convert

    def convert_then_stop(job):
        result = original_convert(job)
        state["stop"] = True
        return result

    backend.convert = convert_then_stop

    jobs.run_batch(
        sink, backend,
        target=str(tmp_path), recursive=False, overwrite=True, use_safe_copy=True,
        force_one_page=True, output_formats=("PDF",), lang="ko",
        is_stopped=lambda: state["stop"],
    )

    converted = [name for name, _fmt in server.backends[0].converted]
    assert len(converted) == 1
    assert sink.done()[4] is False
