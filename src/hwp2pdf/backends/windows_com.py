"""Hancom Office Hangul COM engine (Windows only).

Every function here is moved verbatim from ``app.py``. The module must stay
importable on macOS and Linux -- all ``win32``/``winreg``/``pythoncom`` imports
are function-local, and ``IS_WINDOWS`` gates the entry points.
"""

import csv
import os
import shutil
import struct
import subprocess
import threading
from pathlib import Path

from hwp2pdf import paths
from hwp2pdf.backends.base import (
    BackendCapabilities,
    BackendUnavailable,
    JobResult,
)
from hwp2pdf.constants import (
    DEFAULT_OPEN_OPTION,
    HANCOM_BLOCKING_DIALOG_MESSAGES,
    HANCOM_DIALOG_CONFIRM_BUTTONS,
    HWP_FILEHEADER_STREAM,
    HWP_FILE_SIGNATURE,
    HWP_FLAG_DISTRIBUTION_DOCUMENT,
    HWP_FLAG_PASSWORD_PROTECTED,
    HWP_SECURITY_DLL_NAME,
    HWP_SECURITY_INSTALL_DIR,
    HWP_SECURITY_MODULE,
    HWP_SECURITY_REG_KEY,
    HWP_SECURITY_REG_VALUE,
    MESSAGE_BOX_AUTO_CONFIRM,
    SAVE_FORMAT_ALIASES,
)
from hwp2pdf.i18n import translate

IS_WINDOWS = os.name == "nt"

# ``_resource_root`` moved to ``paths`` when the update/security directories
# became platform-aware; keep the original private name for the moved code.
_resource_root = paths.resource_root

def ensure_pywin32():
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401

        return True, ""
    except Exception as e:
        return False, str(e)


def _console_text_kwargs():
    """Decode console tool output with the OEM code page, not UTF-8.

    ``tasklist`` and ``taskkill`` emit localized messages in the console OEM code
    page. Under UTF-8 mode (``PYTHONUTF8=1``) a bare ``text=True`` decodes Korean
    Windows' "정보: 실행 중인 작업 중 ..." as UTF-8, which kills subprocess's reader
    thread and leaves ``stdout`` as ``None`` -- so process detection silently
    stopped working and a traceback was printed on every check.
    """
    if IS_WINDOWS:
        return {"text": True, "encoding": "oem", "errors": "replace"}
    return {"text": True}


def get_hwp_processes():
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Hwp.exe", "/FO", "CSV", "/NH"],
            capture_output=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            **_console_text_kwargs(),
        )
        if result.returncode != 0 or not result.stdout:
            return []

        processes = []
        for row in csv.reader(result.stdout.splitlines()):
            if len(row) >= 2 and row[0].lower() == "hwp.exe":
                processes.append({"name": row[0], "pid": row[1]})
        return processes
    except Exception:
        return []


def is_hwp_running():
    return bool(get_hwp_processes())


def kill_hwp():
    try:
        result = subprocess.run(
            ["taskkill", "/IM", "Hwp.exe", "/F"],
            capture_output=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            **_console_text_kwargs(),
        )
        return result.returncode == 0 and not is_hwp_running()
    except Exception:
        return False


def read_hwp_file_flags(path: Path):
    if path.suffix.lower() != ".hwp":
        return None

    try:
        import pythoncom

        stgm_read = 0
        stgm_share_exclusive = 0x10
        storage = pythoncom.StgOpenStorage(str(path), None, stgm_read | stgm_share_exclusive)
        stream = storage.OpenStream(HWP_FILEHEADER_STREAM, None, stgm_read | stgm_share_exclusive)
        data = stream.Read(256)
    except Exception:
        return None

    if len(data) < 40 or not data.startswith(HWP_FILE_SIGNATURE):
        return None

    return struct.unpack("<I", data[36:40])[0]


def blocked_conversion_reason(src_path: Path, output_format: str, lang: str = "ko"):
    flags = read_hwp_file_flags(src_path)
    if flags is None:
        return None

    if flags & HWP_FLAG_PASSWORD_PROTECTED:
        return translate(lang, "password_blocked")

    if output_format == "PDF" and flags & HWP_FLAG_DISTRIBUTION_DOCUMENT:
        return translate(lang, "distribution_blocked")

    return None


def build_save_failure_message(output_format: str, errors, lang: str = "ko"):
    detail = translate(lang, "save_failed", format=output_format, errors="; ".join(errors))
    if output_format == "PDF":
        return f"{translate(lang, 'pdf_blocked')} {detail}"
    return detail


def set_hwp_parameter(pset, name: str, value):
    try:
        setattr(pset, name, value)
    except Exception:
        pass

    for target in (pset, getattr(pset, "HSet", None)):
        if target is None:
            continue
        try:
            target.SetItem(name, value)
        except Exception:
            pass


def is_nup_print_method(print_method):
    return print_method in {3, 4, 5, 6, 7, 8, 9, 10}


def save_document_as(hwp, save_target: Path, output_format: str, lang: str = "ko"):
    errors = []
    for save_format in SAVE_FORMAT_ALIASES[output_format]:
        previous_message_box_mode = None
        try:
            if output_format == "DOCX":
                _, previous_message_box_mode, _ = enable_auto_confirm_message_boxes(hwp)
            saved = hwp.SaveAs(str(save_target), save_format, "")
            if saved is not False and save_target.exists():
                return save_format
            errors.append(f"SaveAs {save_format} returned {saved}")
        except Exception as e:
            errors.append(f"SaveAs {save_format}: {e}")
        finally:
            if output_format == "DOCX":
                restore_message_box_mode(hwp, previous_message_box_mode)

    for save_format in SAVE_FORMAT_ALIASES[output_format]:
        previous_message_box_mode = None
        try:
            pset = hwp.HParameterSet.HFileOpenSave
            hwp.HAction.GetDefault("FileSaveAs_S", pset.HSet)

            # pywin32 can expose this property with different casing depending on generated wrappers.
            for attr in ("filename", "FileName"):
                try:
                    setattr(pset, attr, str(save_target))
                except Exception:
                    pass

            pset.Format = save_format
            try:
                pset.Attributes = 0
            except Exception:
                pass

            if output_format == "DOCX":
                _, previous_message_box_mode, _ = enable_auto_confirm_message_boxes(hwp)
            executed = hwp.HAction.Execute("FileSaveAs_S", pset.HSet)
            if executed is not False and save_target.exists():
                return save_format
            errors.append(f"FileSaveAs_S {save_format} returned {executed}")
        except Exception as e:
            errors.append(f"FileSaveAs_S {save_format}: {e}")
        finally:
            if output_format == "DOCX":
                restore_message_box_mode(hwp, previous_message_box_mode)

    raise RuntimeError(build_save_failure_message(output_format, errors, lang))


def force_one_page_view(hwp, lang: str = "ko"):
    ps = hwp.HParameterSet.HViewProperties
    try:
        hwp.HAction.GetDefault("ViewZoom", ps.HSet)
    except Exception:
        pass

    # ZoomCustomDlg + ZoomCntX/ZoomCntY is the Hancom action pattern for explicit multi-page view.
    # 1 x 1 forces the document back to a single-page view before PDF export.
    set_hwp_parameter(ps, "ZoomCustomDlg", 1)
    set_hwp_parameter(ps, "ZoomCntX", 1)
    set_hwp_parameter(ps, "ZoomCntY", 1)
    set_hwp_parameter(ps, "ZoomType", 1)
    set_hwp_parameter(ps, "PageDir", 0)
    executed = hwp.HAction.Execute("ViewZoom", ps.HSet)
    if executed is False:
        raise RuntimeError(translate(lang, "view_failed"))


def configure_pdf_print(hwp, save_target: Path | None = None):
    ps = hwp.HParameterSet.HPrint
    try:
        hwp.HAction.GetDefault("PrintToPDFEx", ps.HSet)
    except Exception:
        try:
            hwp.HAction.GetDefault("FilePrint", ps.HSet)
        except Exception:
            pass

    original_print_method = None
    try:
        original_print_method = int(ps.PrintMethod)
    except Exception:
        pass

    if save_target is not None:
        set_hwp_parameter(ps, "FileName", str(save_target))
        set_hwp_parameter(ps, "filename", str(save_target))

    values = {
        "Collate": 1,
        "UserOrder": 0,
        "PrintToFile": 0,
        "NumCopy": 1,
        "PrinterName": "Hancom PDF",
        "UsingPagenum": 1,
        "ReverseOrder": 0,
        "Pause": 0,
        "PrintImage": 1,
        "PrintDrawObj": 1,
        "PrintClickHere": 0,
        "PrintAutoFootnoteLtext": "^f",
        "PrintAutoFootnoteCtext": "^t",
        "PrintAutoFootnoteRtext": "^P쪽 중 ^p쪽",
        "PrintAutoHeadnoteLtext": "^c",
        "PrintAutoHeadnoteCtext": "^n",
        "PrintAutoHeadnoteRtext": "^p",
        "PrintFormObj": 1,
        "PrintMarkPen": 0,
        "PrintMemo": 0,
        "PrintMemoContents": 0,
        "PrintRevision": 1,
        "PrintBarcode": 1,
        "PrintPronounce": 0,
        # 0 = automatic/basic print. This clears saved N-up / multiple-pages print mode before SaveAs PDF.
        "PrintMethod": 0,
    }
    for name, value in values.items():
        set_hwp_parameter(ps, name, value)

    return ps, original_print_method


def reset_pdf_print_method(hwp, lang: str = "ko"):
    ps, _original_print_method = configure_pdf_print(hwp)

    try:
        executed = hwp.HAction.Execute("PrintToPDFEx", ps.HSet)
    except Exception as e:
        raise RuntimeError(f"{translate(lang, 'pdf_print_method_failed')}: {e}") from e

    if executed is False:
        raise RuntimeError(translate(lang, "pdf_print_method_failed"))


def save_pdf_with_print_to_pdf(hwp, save_target: Path, lang: str = "ko"):
    ps, original_print_method = configure_pdf_print(hwp, save_target)

    try:
        executed = hwp.HAction.Execute("PrintToPDFEx", ps.HSet)
    except Exception as e:
        raise RuntimeError(f"{translate(lang, 'pdf_print_save_failed')}: {e}") from e

    if executed is False or not save_target.exists():
        raise RuntimeError(translate(lang, "pdf_print_save_failed"))

    return "PrintToPDFEx", original_print_method


def enable_auto_confirm_message_boxes(hwp):
    previous_mode = None
    try:
        previous_mode = hwp.GetMessageBoxMode()
    except Exception:
        pass

    try:
        mode = MESSAGE_BOX_AUTO_CONFIRM
        if isinstance(previous_mode, int):
            mode = previous_mode | MESSAGE_BOX_AUTO_CONFIRM
        hwp.SetMessageBoxMode(mode)
        return True, previous_mode, ""
    except Exception as e:
        return False, previous_mode, str(e)


def restore_message_box_mode(hwp, previous_mode):
    if previous_mode is None:
        return
    try:
        hwp.SetMessageBoxMode(previous_mode)
    except Exception:
        pass


def hwp_process_id(hwp):
    try:
        hwnd = int(hwp.XHwpWindows.Item(0).Handle)
    except Exception:
        return None

    try:
        import win32process

        _thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid or None
    except Exception:
        return None


class HancomDialogWatcher:
    def __init__(self, process_id):
        self.process_id = process_id
        self.stop_event = threading.Event()
        self.thread = None
        self.lock = threading.Lock()
        self.closed_messages = []

    def start(self):
        if not self.process_id:
            return
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def stop(self):
        self.stop_event.set()
        if self.thread:
            self.thread.join(timeout=1)

    def mark(self):
        with self.lock:
            return len(self.closed_messages)

    def blocking_message_since(self, marker):
        with self.lock:
            messages = self.closed_messages[marker:]
        for message in messages:
            if any(text in message for text in HANCOM_BLOCKING_DIALOG_MESSAGES):
                return message
        return ""

    def _record(self, message):
        with self.lock:
            if message and message not in self.closed_messages:
                self.closed_messages.append(message)

    def _run(self):
        try:
            import win32con
            import win32gui
            import win32process
        except Exception:
            return

        def child_texts(hwnd):
            values = []

            def enum_child(child_hwnd, _param):
                try:
                    text = win32gui.GetWindowText(child_hwnd).strip()
                    if text:
                        values.append((child_hwnd, text))
                except Exception:
                    pass

            try:
                win32gui.EnumChildWindows(hwnd, enum_child, None)
            except Exception:
                pass
            return values

        def click_confirm_button(hwnd, children):
            for child_hwnd, text in children:
                if text.replace("&", "") in HANCOM_DIALOG_CONFIRM_BUTTONS:
                    try:
                        win32gui.SendMessage(child_hwnd, win32con.BM_CLICK, 0, 0)
                        return True
                    except Exception:
                        pass
            try:
                win32gui.PostMessage(hwnd, win32con.WM_COMMAND, win32con.IDOK, 0)
                return True
            except Exception:
                return False

        def enum_window(hwnd, _param):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                _thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid != self.process_id:
                    return True
                if win32gui.GetClassName(hwnd) != "#32770":
                    return True

                title = win32gui.GetWindowText(hwnd).strip()
                children = child_texts(hwnd)
                message = " | ".join([text for text in [title, *(value for _hwnd, value in children)] if text])
                if not message:
                    return True

                if click_confirm_button(hwnd, children):
                    self._record(message)
            except Exception:
                pass
            return True

        while not self.stop_event.is_set():
            try:
                win32gui.EnumWindows(enum_window, None)
            except Exception:
                pass
            self.stop_event.wait(0.25)


def _bundled_security_dll(arch: str) -> Path:
    return _resource_root() / "vendor" / arch / HWP_SECURITY_DLL_NAME


def _hwp_install_path() -> Path | None:
    import winreg

    candidates = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\HNC\HwpRun"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Hancom\HwpRun"),
    ]
    for hive, subkey in candidates:
        try:
            with winreg.OpenKey(hive, subkey) as key:
                index = 0
                while True:
                    try:
                        version_name = winreg.EnumKey(key, index)
                    except OSError:
                        break
                    index += 1
                    try:
                        with winreg.OpenKey(key, version_name) as version_key:
                            for name in ("Path", "BinPath", ""):
                                try:
                                    value, _ = winreg.QueryValueEx(version_key, name)
                                    if isinstance(value, str) and value.strip():
                                        return Path(value).expanduser()
                                except OSError:
                                    continue
                    except OSError:
                        continue
        except OSError:
            continue

    try:
        import win32com.client

        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        try:
            install_path = hwp.GetHwpInfo("InstallPath")
            if isinstance(install_path, str) and install_path.strip():
                return Path(install_path).expanduser()
        finally:
            try:
                hwp.Quit()
            except Exception:
                pass
    except Exception:
        return None

    return None


def _pe_machine(path: Path) -> int | None:
    try:
        with path.open("rb") as f:
            dos = f.read(64)
            if len(dos) < 64 or dos[:2] != b"MZ":
                return None
            (e_lfanew,) = struct.unpack("<I", dos[60:64])
            f.seek(e_lfanew)
            sig = f.read(4)
            if sig != b"PE\0\0":
                return None
            machine_bytes = f.read(2)
            if len(machine_bytes) != 2:
                return None
            return struct.unpack("<H", machine_bytes)[0]
    except OSError:
        return None


def detect_hwp_arch() -> str:
    install_path = _hwp_install_path()
    if install_path:
        hwp_exe = install_path / "Hwp.exe" if install_path.is_dir() else install_path
        if hwp_exe.exists():
            machine = _pe_machine(hwp_exe)
            if machine == 0x8664:
                return "x64"
            if machine == 0x014C:
                return "x86"
        parts = {part.lower() for part in install_path.parts}
        if "program files (x86)" in parts:
            return "x86"
        if "program files" in parts:
            return "x64"

    return "x86"


def _registered_security_dll() -> Path | None:
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, HWP_SECURITY_REG_KEY) as key:
            value, _ = winreg.QueryValueEx(key, HWP_SECURITY_REG_VALUE)
            if isinstance(value, str) and value.strip():
                return Path(value)
    except OSError:
        return None
    return None


def ensure_hwp_security_module_registered():
    """Make sure HKCU\\Software\\HNC\\HwpAutomation\\Modules\\FilePathCheckerModule
    points to a usable DLL. Copies the bundled stub for the matching HWP bitness
    into %LOCALAPPDATA%\\hwp2pdf\\security\\ and writes the registry value when needed.

    Returns (state, detail) where state is one of:
      "already": registry already had a valid DLL
      "registered": we copied the DLL and wrote the registry
      "bundled-missing": vendor DLL is not bundled with this build
      "error: <reason>": something else went wrong
    """
    if os.name != "nt":
        return "error: non-windows", ""

    arch = detect_hwp_arch()
    expected_machine = 0x8664 if arch == "x64" else 0x014C

    existing = _registered_security_dll()
    if existing and existing.exists() and existing.stat().st_size > 0:
        existing_machine = _pe_machine(existing)
        if existing_machine is None or existing_machine == expected_machine:
            return "already", str(existing)

    source = _bundled_security_dll(arch)
    if not source.exists():
        return "bundled-missing", str(source)

    try:
        HWP_SECURITY_INSTALL_DIR.mkdir(parents=True, exist_ok=True)
        target = HWP_SECURITY_INSTALL_DIR / HWP_SECURITY_DLL_NAME
        needs_copy = True
        if target.exists():
            try:
                needs_copy = target.stat().st_size != source.stat().st_size
            except OSError:
                needs_copy = True
        if needs_copy:
            shutil.copy2(source, target)
    except Exception as e:
        return f"error: copy: {e}", str(source)

    try:
        import winreg

        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, HWP_SECURITY_REG_KEY) as key:
            winreg.SetValueEx(key, HWP_SECURITY_REG_VALUE, 0, winreg.REG_SZ, str(target))
    except Exception as e:
        return f"error: registry: {e}", str(target)

    return "registered", f"{arch}:{target}"


def register_hwp_security_module(hwp):
    try:
        module_name, module_class = HWP_SECURITY_MODULE
        return bool(hwp.RegisterModule(module_name, module_class)), ""
    except Exception as e:
        return False, str(e)


class WindowsComBackend:
    """Drives the locally installed Hancom Office Hangul through COM.

    The per-job body is the ``hwp.Open`` -> save -> ``hwp.Clear`` sequence that
    used to sit inline in ``ConverterApp._run_conversion``; session setup and
    teardown are that method's ``try``/``finally`` blocks.
    """

    capabilities = BackendCapabilities(
        name="windows_com",
        remote=False,
        local_staging=True,
        manages_hwp_process=True,
        local_preflight=True,
    )

    def __init__(self, job_timeout: float | None = None):
        #: Seconds a single conversion may take before Hangul is force-closed
        #: and restarted. ``None`` keeps the original behaviour of waiting
        #: forever, which is what a stuck Hancom dialog does today.
        self.job_timeout = job_timeout
        self.hwp = None
        self.dialog_watcher = None
        self.global_message_box_mode = None
        self.security_ok = False
        self.security_detail = ""
        self._com_initialized = False
        self._sink = None
        self._session_lang = "ko"
        self._engine_broken = False
        self._timed_out = False

    def preflight(self, lang: str) -> None:
        if not IS_WINDOWS:
            raise BackendUnavailable(translate(lang, "backend_requires_windows"))
        ok, detail = ensure_pywin32()
        if not ok:
            raise BackendUnavailable(translate(lang, "pywin32_missing", detail=detail))

        hwp_status = probe_hwp()
        if not hwp_status.get("installed"):
            raise BackendUnavailable(
                translate(lang, "local_hwp_missing", detail=hwp_status.get("detail", ""))
            )

    def open_session(self, sink, lang: str, options) -> None:
        import pythoncom

        self._sink = sink
        self._session_lang = lang

        try:
            sink.put(("log", translate(lang, "init_com")))
            pythoncom.CoInitialize()
            self._com_initialized = True

            paths.temp_workdir().mkdir(parents=True, exist_ok=True)
            self._start_engine(sink, lang)
        except Exception as e:
            self.close_session()
            raise BackendUnavailable(
                translate(lang, "local_hwp_start_failed", detail=e)
            ) from None

    def _start_engine(self, sink, lang: str) -> None:
        """Bring up an HwpObject and everything that hangs off it."""
        import win32com.client

        sink.put(("log", translate(lang, "start_hwp")))
        self.hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        sink.put(("log", translate(lang, "hwp_started")))

        try:
            self.hwp.XHwpWindows.Item(0).Visible = False
        except Exception:
            pass
        try:
            _enabled, self.global_message_box_mode, _detail = enable_auto_confirm_message_boxes(self.hwp)
        except Exception:
            self.global_message_box_mode = None
        try:
            self.dialog_watcher = HancomDialogWatcher(hwp_process_id(self.hwp))
            self.dialog_watcher.start()
        except Exception:
            self.dialog_watcher = None

        sink.put(("log", translate(lang, "register_security")))
        state, detail = ensure_hwp_security_module_registered()
        if state == "registered":
            sink.put(("log", translate(lang, "security_self_registered", detail=detail)))
        elif state == "bundled-missing":
            sink.put(
                ("log", (translate(lang, "security_bundle_missing", detail=detail), "warning"))
            )
        elif state.startswith("error"):
            sink.put(
                (
                    "log",
                    (
                        translate(lang, "security_self_register_failed", state=state, detail=detail),
                        "warning",
                    ),
                )
            )
        self.security_ok, self.security_detail = register_hwp_security_module(self.hwp)
        self._engine_broken = False

    def _abandon_engine(self, job) -> None:
        """Watchdog: the conversion overran, so take the engine down.

        Killing Hwp.exe makes the blocked COM call fail, which unblocks the
        worker; the next convert() brings a fresh engine up.
        """
        self._timed_out = True
        self._engine_broken = True
        if self._sink is not None:
            self._sink.put((
                "log",
                (
                    translate(
                        job.lang, "engine_timeout_kill",
                        seconds=int(self.job_timeout), name=job.src_path.name,
                    ),
                    "warning",
                ),
            ))
        try:
            kill_hwp()
        except Exception:
            pass

    def _restart_engine(self, lang: str) -> bool:
        """Replace a dead engine so one bad file cannot end the batch."""
        try:
            if self.dialog_watcher is not None:
                self.dialog_watcher.stop()
        except Exception:
            pass
        self.dialog_watcher = None
        self.hwp = None
        try:
            self._start_engine(self._sink, lang)
        except Exception as e:
            if self._sink is not None:
                self._sink.put(
                    ("log", (translate(lang, "engine_restart_failed", detail=e), "error"))
                )
            return False
        if self._sink is not None:
            self._sink.put(("log", translate(lang, "engine_restarted")))
        return True

    def session_notes(self, lang: str) -> list:
        on_label = translate(lang, "on")
        off_label = translate(lang, "off")
        if self.security_ok:
            state = on_label
        else:
            state = f"{off_label} ({self.security_detail or translate(lang, 'module_unavailable')})"
        return [("log", translate(lang, "security_module", state=state))]

    def blocked_reason(self, src_path: Path, output_format: str, lang: str):
        return blocked_conversion_reason(src_path, output_format, lang)

    def convert(self, job) -> JobResult:
        from hwp2pdf.i18n import print_method_label

        if self._engine_broken and not self._restart_engine(job.lang):
            return JobResult(ok=False, message=translate(job.lang, "engine_restart_failed", detail=""))

        self._timed_out = False
        watchdog = None
        if self.job_timeout:
            watchdog = threading.Timer(self.job_timeout, self._abandon_engine, args=(job,))
            watchdog.daemon = True
            watchdog.start()

        marker = self.dialog_watcher.mark() if self.dialog_watcher else 0
        source_format = job.src_path.suffix.replace(".", "").upper()
        notices = []

        try:
            opened = self.hwp.Open(str(job.open_path), "", DEFAULT_OPEN_OPTION)
            if opened is False:
                raise RuntimeError(translate(job.lang, "open_failed", format=source_format))

            if job.force_one_page:
                force_one_page_view(self.hwp, job.lang)

            if job.force_one_page and job.output_format == "PDF":
                actual_format, original_print_method = save_pdf_with_print_to_pdf(
                    self.hwp, job.save_path, job.lang
                )
                if is_nup_print_method(original_print_method):
                    notices.append(
                        translate(
                            job.lang,
                            "nup_print_reset",
                            method=print_method_label(original_print_method, job.lang),
                        )
                    )
            else:
                actual_format = save_document_as(
                    self.hwp, job.save_path, job.output_format, job.lang
                )

            if self.dialog_watcher:
                blocking_message = self.dialog_watcher.blocking_message_since(marker)
                if blocking_message:
                    raise RuntimeError(
                        translate(job.lang, "hancom_dialog_blocked", message=blocking_message)
                    )

            self._clear_document()
            return JobResult(ok=True, actual_format=actual_format, notices=notices)

        except Exception as e:
            if self._timed_out:
                # The watchdog killed Hangul out from under this call; report the
                # timeout rather than the RPC error it produced.
                return JobResult(
                    ok=False,
                    message=translate(job.lang, "job_timeout", seconds=int(self.job_timeout)),
                    notices=notices,
                )
            failure_message = str(e)
            if self.dialog_watcher:
                blocking_message = self.dialog_watcher.blocking_message_since(marker)
                if blocking_message:
                    failure_message = translate(
                        job.lang, "hancom_dialog_blocked", message=blocking_message
                    )
            self._clear_document()
            return JobResult(ok=False, message=failure_message, notices=notices)

        finally:
            if watchdog is not None:
                watchdog.cancel()

    def cancel(self) -> None:
        # COM conversion is synchronous; the batch loop stops between jobs.
        pass

    def close_session(self) -> None:
        try:
            if self.dialog_watcher is not None:
                self.dialog_watcher.stop()
        except Exception:
            pass
        try:
            if self.hwp is not None:
                restore_message_box_mode(self.hwp, self.global_message_box_mode)
        except Exception:
            pass
        try:
            if self.hwp is not None:
                self.hwp.Quit()
        except Exception:
            pass
        self.hwp = None
        self.dialog_watcher = None
        if self._com_initialized:
            try:
                import pythoncom

                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_initialized = False

    def _clear_document(self) -> None:
        try:
            self.hwp.Clear(1)
        except Exception:
            pass


def probe_hwp() -> dict:
    """Cheap availability report for the conversion server's capabilities call.

    Deliberately avoids ``Dispatch``: creating an HwpObject just to answer a
    health check would launch Hangul on every poll.
    """
    if not IS_WINDOWS:
        return {"installed": False, "detail": "not windows", "running": []}

    ok, detail = ensure_pywin32()
    if not ok:
        return {"installed": False, "detail": f"pywin32: {detail}", "running": []}

    # The ProgID registration is the marker that actually matters and is the
    # one thing COM needs. HwpRun/install-path keys are a nice-to-have: a 32-bit
    # Hangul on 64-bit Windows files them under WOW6432Node, and Office 2022
    # does not create HwpRun at all.
    marker = None
    registered = False
    try:
        import winreg

        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "HWPFrame.HwpObject"):
                registered = True
        except OSError:
            registered = False

        for subkey in (
            r"SOFTWARE\HNC\HwpRun",
            r"SOFTWARE\Hancom\HwpRun",
            r"SOFTWARE\WOW6432Node\HNC\Hwp",
            r"SOFTWARE\HNC\Hwp",
        ):
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subkey):
                    marker = subkey
                    break
            except OSError:
                continue
    except Exception:
        pass

    running = [str(process["pid"]) for process in get_hwp_processes()]
    if not registered:
        return {
            "installed": False,
            "detail": "HWPFrame.HwpObject is not registered; install Hancom Office "
                      "or run 'Hwp.exe -regserver' from an elevated shell",
            "running": running,
        }
    return {
        "installed": True,
        "detail": marker or "HWPFrame.HwpObject registered",
        "running": running,
    }
