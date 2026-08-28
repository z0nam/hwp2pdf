import csv

import pytest

from fakes import DOCX_STUB, PDF_STUB, FakeBackend, RecordingSink

from hwp2pdf import jobs
from hwp2pdf.i18n import translate


def make_files(tmp_path, *names):
    for name in names:
        (tmp_path / name).write_bytes(b"fake hwp")
    return tmp_path


def run(tmp_path, backend=None, sink=None, **overrides):
    options = dict(
        target=str(tmp_path),
        recursive=True,
        overwrite=True,
        use_safe_copy=False,
        force_one_page=True,
        output_formats=("PDF",),
        lang="ko",
    )
    options.update(overrides)
    sink = sink or RecordingSink()
    backend = backend or FakeBackend()
    jobs.run_batch(sink, backend, **options)
    return sink, backend


def read_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.reader(f))


def test_collect_files_filters_by_extension(tmp_path):
    make_files(tmp_path, "a.hwp", "b.hwpx", "c.txt")
    (tmp_path / "sub").mkdir()
    (tmp_path / "sub" / "d.hwp").write_bytes(b"x")

    flat = {p.name for p in jobs.collect_files(str(tmp_path), False)}
    deep = {p.name for p in jobs.collect_files(str(tmp_path), True)}

    assert flat == {"a.hwp", "b.hwpx"}
    assert deep == {"a.hwp", "b.hwpx", "d.hwp"}


def test_single_file_target(tmp_path):
    make_files(tmp_path, "a.hwp")
    assert jobs.collect_files(str(tmp_path / "a.hwp"), False) == [tmp_path / "a.hwp"]
    assert jobs.collect_files(str(tmp_path / "a.txt"), False) == []


def test_happy_path_writes_output_and_removes_csv(tmp_path):
    make_files(tmp_path, "a.hwp", "b.hwpx")
    sink, backend = run(tmp_path)

    assert (tmp_path / "a.pdf").read_bytes() == PDF_STUB
    assert (tmp_path / "b.pdf").read_bytes() == PDF_STUB
    assert backend.sessions_opened == 1
    assert backend.sessions_closed == 1
    # A fully clean run deletes its own audit log.
    assert not (tmp_path / jobs.LOG_CSV_NAME).exists()
    assert sink.done() == (2, 0, 0, str(tmp_path / jobs.LOG_CSV_NAME), True)


def test_total_jobs_is_files_times_formats(tmp_path):
    make_files(tmp_path, "a.hwp", "b.hwp")
    sink, backend = run(tmp_path, output_formats=("PDF", "DOCX"))

    assert len(backend.converted) == 4
    assert (tmp_path / "a.docx").read_bytes() == DOCX_STUB
    totals = {total for _current, total, _label in sink.of_kind("progress")}
    assert totals == {4}


def test_progress_events_run_from_zero_to_total(tmp_path):
    make_files(tmp_path, "a.hwp", "b.hwp")
    sink, _ = run(tmp_path)

    progress = [(current, total) for current, total, _label in sink.of_kind("progress")]
    assert progress == [(0, 2), (1, 2), (1, 2), (2, 2)]


def test_no_files_emits_error(tmp_path):
    sink, backend = run(tmp_path)

    assert sink.of_kind("error") == [translate("ko", "no_files", extensions=".HWP, .HWPX")]
    assert sink.done() is None
    assert backend.sessions_opened == 0


def test_existing_output_is_skipped_when_not_overwriting(tmp_path):
    make_files(tmp_path, "a.hwp")
    (tmp_path / "a.pdf").write_bytes(b"existing pdf")

    sink, backend = run(tmp_path, overwrite=False)

    assert backend.converted == []
    assert (tmp_path / "a.pdf").read_bytes() == b"existing pdf"
    rows = read_csv(tmp_path / jobs.LOG_CSV_NAME)
    assert rows[1][0] == "SKIPPED"
    assert sink.done()[:3] == (0, 0, 1)


def test_zero_byte_output_is_regenerated_even_without_overwrite(tmp_path):
    make_files(tmp_path, "a.hwp")
    (tmp_path / "a.pdf").write_bytes(b"")

    _sink, backend = run(tmp_path, overwrite=False)

    assert backend.converted == [("a.hwp", "PDF")]
    assert (tmp_path / "a.pdf").read_bytes() == PDF_STUB


def test_overwrite_replaces_existing_output(tmp_path):
    make_files(tmp_path, "a.hwp")
    (tmp_path / "a.pdf").write_bytes(b"stale")

    _sink, backend = run(tmp_path, overwrite=True)

    assert backend.converted == [("a.hwp", "PDF")]
    assert (tmp_path / "a.pdf").read_bytes() == PDF_STUB


def test_blocked_file_fails_without_converting(tmp_path):
    make_files(tmp_path, "a.hwp")
    backend = FakeBackend(blocked={"a.hwp": "배포용 문서"})

    sink, backend = run(tmp_path, backend=backend)

    assert backend.converted == []
    rows = read_csv(tmp_path / jobs.LOG_CSV_NAME)
    assert rows[1][0] == "FAILED"
    assert rows[1][3] == "배포용 문서"
    assert sink.done()[:3] == (0, 1, 0)


def test_failed_conversion_is_logged_and_batch_continues(tmp_path):
    make_files(tmp_path, "a.hwp", "b.hwp")
    backend = FakeBackend(fail_on={"a.hwp"})

    sink, backend = run(tmp_path, backend=backend)

    assert len(backend.converted) == 2
    statuses = [row[0] for row in read_csv(tmp_path / jobs.LOG_CSV_NAME)[1:]]
    assert sorted(statuses) == ["FAILED", "OK"]
    assert sink.done()[:3] == (1, 1, 0)
    assert any("fake failure: a.hwp" in text for text in sink.logs())


def _run_until_stopped(tmp_path, backend, output_formats):
    state = {"stop": False}
    backend.on_convert = lambda job: state.update(stop=True)
    sink = RecordingSink()
    jobs.run_batch(
        sink,
        backend,
        target=str(tmp_path),
        recursive=True,
        overwrite=True,
        use_safe_copy=False,
        force_one_page=True,
        output_formats=output_formats,
        lang="ko",
        is_stopped=lambda: state["stop"],
    )
    return sink


def test_stop_after_a_file_cancels_and_leaves_the_rest_unconverted(tmp_path):
    make_files(tmp_path, "a.hwp", "b.hwp", "c.hwp")
    backend = FakeBackend()

    sink = _run_until_stopped(tmp_path, backend, ("PDF",))

    assert len(backend.converted) == 1
    assert backend.cancels == 1
    assert backend.sessions_closed == 1
    assert sink.done()[4] is False


def test_stop_mid_format_list_writes_a_stopped_row(tmp_path):
    make_files(tmp_path, "a.hwp", "b.hwp")
    backend = FakeBackend()

    sink = _run_until_stopped(tmp_path, backend, ("PDF", "DOCX"))

    assert backend.converted == [("a.hwp", "PDF")]
    assert backend.cancels == 1
    rows = read_csv(tmp_path / jobs.LOG_CSV_NAME)
    assert rows[-1][0] == "STOPPED"
    assert any(translate("ko", "stopped") in text for text in sink.logs())
    assert sink.done()[4] is False


def test_safe_temp_stages_through_the_temp_workdir(tmp_path, monkeypatch):
    staging = tmp_path / "staging"
    monkeypatch.setattr(jobs.paths, "temp_workdir", lambda: staging)
    source = tmp_path / "docs"
    source.mkdir()
    make_files(source, "a.hwp")

    seen = {}
    backend = FakeBackend(on_convert=lambda job: seen.update(
        open_parent=job.open_path.parent, save_parent=job.save_path.parent
    ))
    run(source, backend=backend, use_safe_copy=True)

    assert seen["open_parent"] == staging
    assert seen["save_parent"] == staging
    assert (source / "a.pdf").read_bytes() == PDF_STUB
    # Staging files are cleaned up after each job.
    assert list(staging.iterdir()) == []


def test_backend_without_local_staging_ignores_safe_temp(tmp_path, monkeypatch):
    staging = tmp_path / "staging"
    monkeypatch.setattr(jobs.paths, "temp_workdir", lambda: staging)
    make_files(tmp_path, "a.hwp")

    backend = FakeBackend()
    backend.capabilities = type(backend.capabilities)(
        name="fake-remote", remote=True, local_staging=False,
        manages_hwp_process=False, local_preflight=False,
    )
    seen = {}
    backend.on_convert = lambda job: seen.update(open_path=job.open_path, save_path=job.save_path)

    run(tmp_path, backend=backend, use_safe_copy=True)

    assert seen["open_path"] == tmp_path / "a.hwp"
    assert seen["save_path"] == tmp_path / "a.pdf"
    assert not staging.exists()


def test_unavailable_backend_reports_error(tmp_path):
    make_files(tmp_path, "a.hwp")
    backend = FakeBackend(unavailable="백엔드 없음")

    sink, backend = run(tmp_path, backend=backend)

    assert sink.of_kind("error") == ["백엔드 없음"]
    assert backend.sessions_opened == 0
    assert sink.done() is None


@pytest.mark.parametrize("lang", ["ko", "en"])
def test_csv_header_is_localized(tmp_path, lang):
    make_files(tmp_path, "a.hwp")
    backend = FakeBackend(fail_on={"a.hwp"})
    run(tmp_path, backend=backend, lang=lang)

    header = read_csv(tmp_path / jobs.LOG_CSV_NAME)[0]
    assert header == [
        translate(lang, "status_header"),
        translate(lang, "source_header"),
        translate(lang, "output_header"),
        translate(lang, "message_header"),
    ]
