"""Batch orchestration shared by the GUI, the CLI and the conversion server.

This is ``ConverterApp._run_conversion`` with the COM-specific parts lifted into
a :class:`~hwp2pdf.backends.base.ConversionBackend`. Everything that depends on
the destination filesystem stays here so both the local Windows path and the
remote macOS path produce identical logs, CSV rows and progress events.
"""

import csv
import shutil
from pathlib import Path

from hwp2pdf import paths
from hwp2pdf.backends.base import BackendUnavailable, JobSpec, SessionOptions
from hwp2pdf.constants import enabled_extensions, output_extension
from hwp2pdf.i18n import translate

LOG_CSV_NAME = "hwp2pdf_log.csv"


def collect_files(target, recursive: bool):
    """Resolve a conversion target to a list of files.

    ``target`` is either a single folder/file path, or an explicit sequence of
    file paths (the GUI's multi-file selection). Handling both here keeps the
    behaviour identical for the local and the remote backend.
    """
    allowed_extensions = enabled_extensions()
    if isinstance(target, (tuple, list)):
        return [
            Path(path)
            for path in target
            if Path(path).is_file() and Path(path).suffix.lower() in allowed_extensions
        ]

    root = Path(target)
    if root.is_file():
        return [root] if root.suffix.lower() in allowed_extensions else []

    iterator = root.rglob("*") if recursive else root.glob("*")
    files = []
    for path in iterator:
        if path.is_file() and path.suffix.lower() in allowed_extensions:
            files.append(path)
    return files


def run_batch(
    sink,
    backend,
    *,
    target,
    recursive: bool,
    overwrite: bool,
    use_safe_copy: bool,
    force_one_page: bool,
    output_formats,
    lang: str,
    is_stopped=None,
    file_collector=None,
):
    """Convert every matching file under ``target``.

    ``sink`` receives ``(kind, payload)`` tuples exactly as before: ``log``,
    ``progress``, ``done`` and ``error``. ``is_stopped`` is polled between jobs
    so the GUI stop button and the CLI behave the same way.
    """
    if is_stopped is None:
        def is_stopped():
            return False
    if file_collector is None:
        file_collector = collect_files

    try:
        sink.put(("log", translate(lang, "scanning")))
        files = file_collector(target, recursive)
        total_files = len(files)
        total_jobs = total_files * len(output_formats)
        extension_label = ", ".join(ext.upper() for ext in enabled_extensions())
        if total_files == 0:
            sink.put(("error", translate(lang, "no_files", extensions=extension_label)))
            return

        if isinstance(target, (tuple, list)):
            # An explicit file selection can span folders; log next to the first.
            log_root = files[0].parent if files else Path.cwd()
        else:
            target_path = Path(target)
            log_root = target_path.parent if target_path.is_file() else target_path
        log_csv = str(log_root / LOG_CSV_NAME)
        success = 0
        failed = 0
        skipped = 0
        stopped = False
        cancel_state = {"sent": False}

        def signal_cancel():
            if cancel_state["sent"]:
                return
            cancel_state["sent"] = True
            try:
                backend.cancel()
            except Exception:
                pass

        temp_workdir = paths.temp_workdir()
        staging_enabled = use_safe_copy and backend.capabilities.local_staging

        session_options = SessionOptions(
            lang=lang,
            output_formats=tuple(output_formats),
            force_one_page=force_one_page,
            safe_temp=use_safe_copy,
            total_files=total_files,
        )

        try:
            backend.preflight(lang)
        except BackendUnavailable as e:
            sink.put(("error", str(e)))
            return

        backend.open_session(sink, lang, session_options)

        try:
            if staging_enabled:
                temp_workdir.mkdir(parents=True, exist_ok=True)

            on_label = translate(lang, "on")
            off_label = translate(lang, "off")
            sink.put(("log", translate(lang, "found_files", count=total_files, extensions=extension_label)))
            sink.put(("log", translate(lang, "csv_log", path=log_csv)))
            sink.put(("log", translate(lang, "safe_temp_mode", state=on_label if use_safe_copy else off_label)))
            sink.put(
                ("log", translate(lang, "force_one_page_mode", state=on_label if force_one_page else off_label))
            )
            sink.put(("log", translate(lang, "output_formats", formats=", ".join(output_formats))))
            sink.put(("log", translate(lang, "auto_confirm_docx")))
            for note in backend.session_notes(lang):
                sink.put(note)

            with open(log_csv, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(
                    [
                        translate(lang, "status_header"),
                        translate(lang, "source_header"),
                        translate(lang, "output_header"),
                        translate(lang, "message_header"),
                    ]
                )
                f.flush()

                job_index = 0
                for file_index, src_path in enumerate(files, start=1):
                    src_path = Path(src_path)

                    sink.put(("log", translate(lang, "processing", path=src_path)))

                    for output_format in output_formats:
                        if is_stopped():
                            signal_cancel()
                            writer.writerow(["STOPPED", "", "", translate(lang, "stopped_csv")])
                            f.flush()
                            sink.put(("log", (translate(lang, "stopped"), "warning")))
                            stopped = True
                            break

                        job_index += 1
                        output_ext = output_extension(output_format)
                        output_path = src_path.with_suffix(output_ext)
                        sink.put(
                            (
                                "progress",
                                (
                                    job_index - 1,
                                    total_jobs,
                                    translate(
                                        lang,
                                        "progress_convert",
                                        current=job_index,
                                        total=total_jobs,
                                        name=src_path.name,
                                        format=output_format,
                                    ),
                                ),
                            )
                        )

                        temp_input = None
                        temp_output = None

                        try:
                            if output_path.exists() and not overwrite:
                                try:
                                    output_size = output_path.stat().st_size
                                except OSError:
                                    output_size = 1

                                if output_size > 0:
                                    skipped += 1
                                    msg = translate(lang, "skipped_exists", format=output_format)
                                    writer.writerow(["SKIPPED", str(src_path), str(output_path), msg])
                                    f.flush()
                                    sink.put(
                                        (
                                            "log",
                                            (
                                                translate(
                                                    lang, "skipped_log", format=output_format, path=output_path
                                                ),
                                                "warning",
                                            ),
                                        )
                                    )
                                    sink.put(
                                        (
                                            "progress",
                                            (
                                                job_index,
                                                total_jobs,
                                                translate(
                                                    lang, "progress_skipped", current=job_index, total=total_jobs
                                                ),
                                            ),
                                        )
                                    )
                                    continue

                                output_path.unlink()

                            blocked_reason = backend.blocked_reason(src_path, output_format, lang)
                            if blocked_reason:
                                failed += 1
                                writer.writerow(["FAILED", str(src_path), str(output_path), blocked_reason])
                                f.flush()
                                sink.put(
                                    (
                                        "log",
                                        (
                                            translate(
                                                lang,
                                                "failed_log",
                                                format=output_format,
                                                path=src_path,
                                                message=blocked_reason,
                                            ),
                                            "error",
                                        ),
                                    )
                                )
                                sink.put(
                                    (
                                        "progress",
                                        (
                                            job_index,
                                            total_jobs,
                                            translate(lang, "progress_failed", current=job_index, total=total_jobs),
                                        ),
                                    )
                                )
                                continue

                            if staging_enabled:
                                temp_input = temp_workdir / f"{file_index:05d}_{output_format}_{src_path.name}"
                                temp_output = (
                                    temp_workdir
                                    / f"{file_index:05d}_{output_format}_{src_path.stem}{output_ext}"
                                )

                                if temp_input.exists():
                                    temp_input.unlink()
                                if temp_output.exists():
                                    temp_output.unlink()

                                shutil.copy2(src_path, temp_input)
                                open_target = temp_input
                                save_target = temp_output
                            else:
                                open_target = src_path
                                save_target = output_path

                            if output_path.exists() and overwrite:
                                output_path.unlink()

                            result = backend.convert(
                                JobSpec(
                                    index=file_index,
                                    src_path=src_path,
                                    open_path=open_target,
                                    save_path=save_target,
                                    output_format=output_format,
                                    force_one_page=force_one_page,
                                    safe_temp=use_safe_copy,
                                    lang=lang,
                                )
                            )

                            for notice in result.notices:
                                sink.put(("log", notice))

                            if not result.ok:
                                raise RuntimeError(result.message)

                            if staging_enabled:
                                if not temp_output.exists():
                                    raise RuntimeError(translate(lang, "temp_missing", format=output_format))
                                shutil.move(str(temp_output), str(output_path))

                            success += 1
                            writer.writerow(["OK", str(src_path), str(output_path), ""])
                            f.flush()
                            sink.put(
                                (
                                    "log",
                                    translate(
                                        lang,
                                        "ok_log",
                                        format=output_format,
                                        actual=result.actual_format,
                                        path=output_path,
                                    ),
                                )
                            )

                        except Exception as e:
                            failed += 1
                            failure_message = str(e)
                            writer.writerow(["FAILED", str(src_path), str(output_path), failure_message])
                            f.flush()
                            sink.put(
                                (
                                    "log",
                                    (
                                        translate(
                                            lang,
                                            "failed_log",
                                            format=output_format,
                                            path=src_path,
                                            message=failure_message,
                                        ),
                                        "error",
                                    ),
                                )
                            )

                        finally:
                            for tmp in (temp_input, temp_output):
                                try:
                                    if tmp and Path(tmp).exists():
                                        Path(tmp).unlink()
                                except Exception:
                                    pass

                        sink.put(
                            (
                                "progress",
                                (
                                    job_index,
                                    total_jobs,
                                    translate(lang, "progress_done", current=job_index, total=total_jobs),
                                ),
                            )
                        )

                    if is_stopped():
                        signal_cancel()
                        stopped = True
                        break

        finally:
            try:
                backend.close_session()
            except Exception:
                pass

        all_success = success == total_jobs and failed == 0 and skipped == 0 and not stopped
        if all_success:
            try:
                Path(log_csv).unlink()
            except FileNotFoundError:
                pass
            except Exception as e:
                all_success = False
                sink.put(("log", (translate(lang, "remove_log_failed", message=e), "warning")))

        sink.put(("done", (success, failed, skipped, log_csv, all_success)))

    except Exception as e:
        sink.put(("error", translate(lang, "unexpected_error", message=e)))
