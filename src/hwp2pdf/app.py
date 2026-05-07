import csv
import os
import queue
import shutil
import struct
import subprocess
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from hwp2pdf.version import __version__

APP_NAME = "HWP/HWPX -> PDF/DOCX Converter"
APP_TITLE = f"{APP_NAME} v{__version__}"
DEFAULT_OPEN_OPTION = "forceopen:true;versionwarning:false;"
TEMP_WORKDIR = Path(r"C:\temp\hwp_convert_safe")
BASE_EXTENSIONS = (".hwp", ".hwpx")
OUTPUT_FORMATS = {
    "PDF": ".pdf",
    "DOCX": ".docx",
}
SAVE_FORMAT_ALIASES = {
    "PDF": ("PDF",),
    "DOCX": ("OOXML", "DOCX", "MSWORD"),
}
HWP_SECURITY_MODULE = ("FilePathCheckDLL", "FilePathCheckerModule")
MESSAGE_BOX_AUTO_CONFIRM = 0x10
HWP_FILEHEADER_STREAM = "FileHeader"
HWP_FILE_SIGNATURE = b"HWP Document File"
HWP_FLAG_COMPRESSED = 1 << 0
HWP_FLAG_PASSWORD_PROTECTED = 1 << 1
HWP_FLAG_DISTRIBUTION_DOCUMENT = 1 << 2
LANGUAGE_LABELS = {
    "ko": "한국어",
    "en": "English",
}

LANGUAGE_CODES = {label: code for code, label in LANGUAGE_LABELS.items()}

TEXT = {
    "ko": {
        "target_label": "대상 폴더 또는 파일",
        "browse_folder": "폴더 선택...",
        "pick_file": "파일 선택...",
        "options": "옵션",
        "include_subfolders": "하위 폴더 포함",
        "overwrite": "기존 출력 파일 덮어쓰기",
        "output": "출력",
        "safe_temp": "안전한 로컬 임시 폴더 변환 사용(구글 드라이브/네트워크 드라이브 권장)",
        "force_one_page": "저장 전 한쪽 보기 강제 적용",
        "start": "변환 시작",
        "stop": "중지",
        "open_selected": "선택 위치 열기",
        "ready": "준비",
        "log": "로그",
        "notes_title": "참고",
        "notes": (
            "- 안정성을 위해 시작 전에 아래한글을 닫아 주세요.\n"
            "- 안전한 임시 폴더 모드는 각 파일을 짧은 로컬 경로로 복사한 뒤 변환합니다.\n"
            "- PDF가 2쪽 보기로 저장되는 문제를 피하려고 기본적으로 한쪽 보기를 강제 적용합니다.\n"
            "- DOCX 출력 품질은 한컴오피스의 DOCX 내보내기 지원에 따라 달라집니다.\n"
            "- 실패, 건너뜀, 중단이 있으면 선택 위치에 CSV 로그가 남습니다."
        ),
        "language": "언어",
        "select_folder_title": "변환 대상 폴더 선택",
        "select_file_title": "변환할 HWP/HWPX 파일 선택",
        "all_files": "모든 파일",
        "invalid_target": "올바른 폴더 또는 HWP/HWPX 파일을 선택하세요.",
        "invalid_file": "HWP 또는 HWPX 파일을 선택하세요.",
        "invalid_open_target": "올바른 폴더 또는 파일을 먼저 선택하세요.",
        "already_running": "이미 변환 작업이 실행 중입니다.",
        "select_output": "출력 형식을 하나 이상 선택하세요: PDF 또는 DOCX.",
        "pywin32_missing": "pywin32를 사용할 수 없습니다.\n\n설치 명령:\npython -m pip install pywin32\n\n상세:\n{detail}",
        "hwp_running_prompt": (
            "아래한글 프로세스가 이미 백그라운드에서 실행 중입니다.\n\n"
            "감지됨: {process_detail}\n\n"
            "예: HWP를 강제 종료하고 계속\n"
            "아니오: 그대로 계속\n"
            "취소: 중단"
        ),
        "hwp_kill_failed": "HWP를 자동으로 종료하지 못했습니다.\n\n작업 관리자에서 Hwp.exe를 닫은 뒤 다시 시작하세요.",
        "starting_conversion": "변환 시작...",
        "scanning": "파일 검색 중...",
        "no_files": "{extensions} 파일을 찾지 못했습니다.",
        "init_com": "한컴 자동화 초기화 중...",
        "start_hwp": "HWPFrame.HwpObject 시작 중...",
        "hwp_started": "HWPFrame.HwpObject가 시작되었습니다.",
        "register_security": "한컴 파일 접근 보안 모듈 등록 중...",
        "found_files": "대상 파일 {count}개 발견: {extensions}",
        "csv_log": "CSV 로그: {path}",
        "safe_temp_mode": "안전 임시 폴더 모드: {state}",
        "force_one_page_mode": "한쪽 보기 강제 적용: {state}",
        "output_formats": "출력 형식: {formats}",
        "auto_confirm_docx": "DOCX 호환 문서 확인창 자동 확인: 켜짐",
        "security_module": "HWP 파일 접근 보안 모듈: {state}",
        "on": "켜짐",
        "off": "꺼짐",
        "module_unavailable": "모듈 사용 불가",
        "processing": "처리 중: {path}",
        "stopped": "사용자가 중지했습니다.",
        "stop_requested": "중지를 요청했습니다. 현재 파일 처리 후 멈춥니다.",
        "skipped_exists": "{format} 파일이 이미 있어 건너뜀",
        "skipped_log": "건너뜀 {format} -> {path}",
        "failed_log": "실패 {format} -> {path} | {message}",
        "ok_log": "성공 {format} ({actual}) -> {path}",
        "progress_skipped": "[{current}/{total}] 건너뜀",
        "progress_failed": "[{current}/{total}] 실패",
        "progress_done": "[{current}/{total}] 완료",
        "progress_convert": "[{current}/{total}] {name} -> {format}",
        "open_failed": "{format} 열기 실패",
        "temp_missing": "임시 {format} 파일이 생성되지 않았습니다.",
        "remove_log_failed": "성공 로그를 삭제하지 못했습니다: {message}",
        "unexpected_error": "예상치 못한 오류:\n{message}",
        "success_status": "성공",
        "success_message": "변환이 완료되었습니다.",
        "done_status": "완료. 성공: {success}, 실패: {failed}, 건너뜀: {skipped}",
        "done_message": (
            "변환이 끝났습니다.\n\n"
            "성공: {success}\n실패: {failed}\n건너뜀: {skipped}\n\n"
            "로그 파일:\n{log_csv}"
        ),
        "error_status": "오류",
        "status_header": "status",
        "source_header": "source",
        "output_header": "output",
        "message_header": "message",
        "stopped_csv": "사용자가 중지 요청",
        "distribution_blocked": (
            "HWP FileHeader에서 배포용 문서 보안이 감지되었습니다. "
            "한컴의 인쇄/PDF 제한으로 PDF 변환이 비활성화되었을 수 있어 이 파일은 열지 않고 실패 처리했습니다."
        ),
        "password_blocked": (
            "HWP FileHeader에서 암호 보호 문서가 감지되었습니다. "
            "문서 암호 없이는 한컴에서 자동으로 열거나 내보낼 수 없습니다."
        ),
        "pdf_blocked": (
            "PDF 내보내기를 사용할 수 없거나 차단되었습니다. 일반적으로 한컴 문서 보안, "
            "배포용 문서 설정, 인쇄/PDF 제한 때문에 발생합니다."
        ),
        "save_failed": "SaveAs {format} 실패. 시도: {errors}",
        "view_failed": "한쪽 보기 설정에 실패했습니다.",
    },
    "en": {
        "target_label": "Root folder or file",
        "browse_folder": "Browse folder...",
        "pick_file": "Pick file...",
        "options": "Options",
        "include_subfolders": "Include subfolders",
        "overwrite": "Overwrite existing output",
        "output": "Output",
        "safe_temp": "Use safe local temp conversion (recommended for Google Drive / network drives)",
        "force_one_page": "Force one-page view before export",
        "start": "Start conversion",
        "stop": "Stop",
        "open_selected": "Open selected folder",
        "ready": "Ready",
        "log": "Log",
        "notes_title": "Notes",
        "notes": (
            "- Close Hancom HWP before starting for best stability.\n"
            "- Safe temp mode copies each file to a short local path before conversion.\n"
            "- One-page view is forced before export by default to avoid two-page PDF output.\n"
            "- DOCX output uses Hancom Office export, so layout fidelity depends on Hancom's DOCX support.\n"
            "- A CSV log is kept in the selected location when there are failures, skips, or stops."
        ),
        "language": "Language",
        "select_folder_title": "Select target folder",
        "select_file_title": "Select one HWP/HWPX file to convert",
        "all_files": "All files",
        "invalid_target": "Select a valid root folder or HWP/HWPX file.",
        "invalid_file": "Select an HWP or HWPX file.",
        "invalid_open_target": "Select a valid folder or file first.",
        "already_running": "A conversion job is already running.",
        "select_output": "Select at least one output format: PDF or DOCX.",
        "pywin32_missing": "pywin32 is not available.\n\nInstall it with:\npython -m pip install pywin32\n\nDetails:\n{detail}",
        "hwp_running_prompt": (
            "Hancom HWP process is already running in the background.\n\n"
            "Detected: {process_detail}\n\n"
            "Yes: force close HWP and continue\n"
            "No: continue anyway\n"
            "Cancel: stop"
        ),
        "hwp_kill_failed": "Could not close HWP automatically.\n\nClose Hwp.exe from Task Manager, then start conversion again.",
        "starting_conversion": "Starting conversion...",
        "scanning": "Scanning files...",
        "no_files": "No {extensions} files found.",
        "init_com": "Initializing Hancom COM automation...",
        "start_hwp": "Starting HWPFrame.HwpObject...",
        "hwp_started": "HWPFrame.HwpObject started.",
        "register_security": "Registering HWP file access security module...",
        "found_files": "Found {count} file(s): {extensions}",
        "csv_log": "CSV log: {path}",
        "safe_temp_mode": "Safe temp mode: {state}",
        "force_one_page_mode": "Force one-page view: {state}",
        "output_formats": "Output formats: {formats}",
        "auto_confirm_docx": "Auto-confirm DOCX compatibility dialogs: ON",
        "security_module": "HWP file access security module: {state}",
        "on": "ON",
        "off": "OFF",
        "module_unavailable": "module unavailable",
        "processing": "Processing: {path}",
        "stopped": "Stopped by user.",
        "stop_requested": "Stop requested. Current file will finish first.",
        "skipped_exists": "Skipped because {format} already exists",
        "skipped_log": "SKIPPED {format} -> {path}",
        "failed_log": "FAILED {format} -> {path} | {message}",
        "ok_log": "OK {format} ({actual}) -> {path}",
        "progress_skipped": "[{current}/{total}] Skipped",
        "progress_failed": "[{current}/{total}] Failed",
        "progress_done": "[{current}/{total}] Done",
        "progress_convert": "[{current}/{total}] {name} -> {format}",
        "open_failed": "Open failed for {format}",
        "temp_missing": "Temporary {format} was not created",
        "remove_log_failed": "Could not remove success log: {message}",
        "unexpected_error": "Unexpected error:\n{message}",
        "success_status": "Success",
        "success_message": "Conversion succeeded.",
        "done_status": "Done. Success: {success}, Failed: {failed}, Skipped: {skipped}",
        "done_message": (
            "Conversion finished.\n\n"
            "Success: {success}\nFailed: {failed}\nSkipped: {skipped}\n\n"
            "Log file:\n{log_csv}"
        ),
        "error_status": "Error",
        "status_header": "status",
        "source_header": "source",
        "output_header": "output",
        "message_header": "message",
        "stopped_csv": "User requested stop",
        "distribution_blocked": (
            "Distribution-document security detected from HWP FileHeader. "
            "PDF export may be disabled by Hancom print/PDF restrictions, so this file was not opened for conversion."
        ),
        "password_blocked": (
            "Password-protected HWP document detected from FileHeader. "
            "Hancom cannot open or export it automatically without the document password."
        ),
        "pdf_blocked": (
            "PDF export is unavailable or blocked. This is commonly caused by Hancom document security "
            "or distribution-document settings such as disabled print/PDF export."
        ),
        "save_failed": "SaveAs {format} failed. Tried: {errors}",
        "view_failed": "ViewZoom one-page setting failed",
    },
}


def translate(lang: str, key: str, **kwargs):
    text = TEXT.get(lang, TEXT["ko"]).get(key, TEXT["ko"].get(key, key))
    return text.format(**kwargs) if kwargs else text


def ensure_pywin32():
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401

        return True, ""
    except Exception as e:
        return False, str(e)


def get_hwp_processes():
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Hwp.exe", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if result.returncode != 0:
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
            text=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return result.returncode == 0 and not is_hwp_running()
    except Exception:
        return False


def enabled_extensions():
    return BASE_EXTENSIONS


def output_extension(output_format: str):
    return OUTPUT_FORMATS[output_format]


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
    hwp.HAction.GetDefault("ViewZoom", ps.HSet)
    ps.ZoomCntX = 1
    ps.ZoomCntY = 1
    ps.PageDir = 0
    executed = hwp.HAction.Execute("ViewZoom", ps.HSet)
    if executed is False:
        raise RuntimeError(translate(lang, "view_failed"))


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


def register_hwp_security_module(hwp):
    try:
        module_name, module_class = HWP_SECURITY_MODULE
        return bool(hwp.RegisterModule(module_name, module_class)), ""
    except Exception as e:
        return False, str(e)


class ConverterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("920x710")
        self.root.minsize(820, 580)

        self.folder_var = tk.StringVar()
        self.overwrite_var = tk.BooleanVar(value=True)
        self.recursive_var = tk.BooleanVar(value=True)
        self.use_safe_copy_var = tk.BooleanVar(value=True)
        self.force_one_page_var = tk.BooleanVar(value=True)
        self.output_pdf_var = tk.BooleanVar(value=True)
        self.output_docx_var = tk.BooleanVar(value=False)
        self.language_var = tk.StringVar(value=LANGUAGE_LABELS["ko"])

        self.log_queue = queue.Queue()
        self.worker = None
        self.stop_requested = False
        self.is_running = False
        self.recursive_check = None
        self.ui = {}

        self._build_ui()
        self.folder_var.trace_add("write", self._on_target_path_changed)
        self._poll_log_queue()

    def lang(self):
        return LANGUAGE_CODES.get(self.language_var.get(), "ko")

    def tr(self, key: str, **kwargs):
        return translate(self.lang(), key, **kwargs)

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=12)
        top.pack(fill="x")

        header_row = ttk.Frame(top)
        header_row.grid(row=0, column=0, sticky="ew")
        header_row.columnconfigure(0, weight=1)
        self.ui["target_label"] = ttk.Label(header_row)
        self.ui["target_label"].grid(row=0, column=0, sticky="w")
        language_frame = ttk.Frame(header_row)
        language_frame.grid(row=0, column=1, sticky="e")
        self.ui["language_label"] = ttk.Label(language_frame)
        self.ui["language_label"].pack(side="left", padx=(0, 6))
        language_combo = ttk.Combobox(
            language_frame,
            textvariable=self.language_var,
            values=[LANGUAGE_LABELS["ko"], LANGUAGE_LABELS["en"]],
            width=10,
            state="readonly",
        )
        language_combo.pack(side="left")
        language_combo.bind("<<ComboboxSelected>>", self._on_language_changed)
        top.columnconfigure(0, weight=1)

        path_row = ttk.Frame(top)
        path_row.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        path_row.columnconfigure(0, weight=1)

        ttk.Entry(path_row, textvariable=self.folder_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.ui["browse_btn"] = ttk.Button(path_row, command=self.browse_folder)
        self.ui["browse_btn"].grid(
            row=0, column=1, sticky="e", padx=(0, 8)
        )
        self.ui["pick_btn"] = ttk.Button(path_row, command=self.pick_file_folder)
        self.ui["pick_btn"].grid(row=0, column=2, sticky="e")

        opts = ttk.LabelFrame(self.root, padding=12)
        self.ui["opts"] = opts
        opts.pack(fill="x", padx=12, pady=(0, 12))

        self.recursive_check = ttk.Checkbutton(opts, variable=self.recursive_var)
        self.recursive_check.grid(row=0, column=0, sticky="w", padx=(0, 16))
        self.ui["overwrite_check"] = ttk.Checkbutton(opts, variable=self.overwrite_var)
        self.ui["overwrite_check"].grid(
            row=0, column=1, sticky="w", padx=(0, 16)
        )

        output_frame = ttk.Frame(opts)
        output_frame.grid(row=0, column=2, sticky="w")
        self.ui["output_label"] = ttk.Label(output_frame)
        self.ui["output_label"].pack(side="left", padx=(0, 8))
        ttk.Checkbutton(output_frame, text="PDF", variable=self.output_pdf_var).pack(side="left")
        ttk.Checkbutton(output_frame, text="DOCX", variable=self.output_docx_var).pack(side="left", padx=(8, 0))
        self.ui["safe_temp_check"] = ttk.Checkbutton(
            opts,
            variable=self.use_safe_copy_var,
        )
        self.ui["safe_temp_check"].grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.ui["force_one_page_check"] = ttk.Checkbutton(
            opts,
            variable=self.force_one_page_var,
        )
        self.ui["force_one_page_check"].grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        actions = ttk.Frame(self.root, padding=(12, 0, 12, 12))
        actions.pack(fill="x")

        self.start_btn = ttk.Button(actions, command=self.start_conversion)
        self.start_btn.pack(side="left")

        self.stop_btn = ttk.Button(actions, command=self.request_stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))

        self.ui["open_btn"] = ttk.Button(actions, command=self.open_selected_folder)
        self.ui["open_btn"].pack(side="left", padx=(8, 0))

        progress_frame = ttk.Frame(self.root, padding=(12, 0, 12, 0))
        progress_frame.pack(fill="x")

        self.progress_label_var = tk.StringVar()
        ttk.Label(progress_frame, textvariable=self.progress_label_var).pack(anchor="w")

        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=(6, 0))

        log_frame = ttk.Frame(self.root, padding=12)
        log_frame.pack(fill="both", expand=True)

        self.ui["log_label"] = ttk.Label(log_frame)
        self.ui["log_label"].pack(anchor="w")
        self.log_text = tk.Text(log_frame, wrap="word")
        self.log_text.pack(side="left", fill="both", expand=True)

        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=log_scroll.set)

        note_frame = ttk.LabelFrame(self.root, padding=12)
        self.ui["note_frame"] = note_frame
        note_frame.pack(fill="x", padx=12, pady=(0, 12))
        self.ui["notes_label"] = ttk.Label(note_frame, justify="left")
        self.ui["notes_label"].pack(anchor="w")
        self._apply_language()

    def _on_language_changed(self, _event=None):
        self._apply_language()

    def _apply_language(self):
        self.root.title(APP_TITLE)
        self.ui["target_label"].configure(text=self.tr("target_label"))
        self.ui["language_label"].configure(text=self.tr("language"))
        self.ui["browse_btn"].configure(text=self.tr("browse_folder"))
        self.ui["pick_btn"].configure(text=self.tr("pick_file"))
        self.ui["opts"].configure(text=self.tr("options"))
        self.recursive_check.configure(text=self.tr("include_subfolders"))
        self.ui["overwrite_check"].configure(text=self.tr("overwrite"))
        self.ui["output_label"].configure(text=self.tr("output"))
        self.ui["safe_temp_check"].configure(text=self.tr("safe_temp"))
        self.ui["force_one_page_check"].configure(text=self.tr("force_one_page"))
        self.start_btn.configure(text=self.tr("start"))
        self.stop_btn.configure(text=self.tr("stop"))
        self.ui["open_btn"].configure(text=self.tr("open_selected"))
        self.ui["log_label"].configure(text=self.tr("log"))
        self.ui["note_frame"].configure(text=self.tr("notes_title"))
        self.ui["notes_label"].configure(text=self.tr("notes"))
        if not self.is_running:
            self.progress_label_var.set(self.tr("ready"))

    def browse_folder(self):
        initial_dir = self.folder_var.get().strip()
        if initial_dir and os.path.isfile(initial_dir):
            initial_dir = str(Path(initial_dir).parent)
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = str(Path.home())

        folder = filedialog.askdirectory(parent=self.root, title=self.tr("select_folder_title"), initialdir=initial_dir)
        if folder:
            self.folder_var.set(folder)

    def pick_file_folder(self):
        initial_dir = self.folder_var.get().strip()
        if initial_dir and os.path.isfile(initial_dir):
            initial_dir = str(Path(initial_dir).parent)
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = str(Path.home())

        selected_file = filedialog.askopenfilename(
            parent=self.root,
            title=self.tr("select_file_title"),
            initialdir=initial_dir,
            filetypes=[
                ("HWP/HWPX", "*.hwp *.hwpx"),
                (self.tr("all_files"), "*.*"),
            ],
        )
        if selected_file:
            self.folder_var.set(selected_file)

    def _on_target_path_changed(self, *_args):
        if self.recursive_check is None:
            return

        target = self.folder_var.get().strip()
        if target and os.path.isfile(target):
            self.recursive_var.set(False)
            self.recursive_check.state(["disabled"])
        else:
            self.recursive_check.state(["!disabled"])

    def open_selected_folder(self):
        target = self.folder_var.get().strip()
        if target and os.path.isfile(target):
            os.startfile(str(Path(target).parent))
        elif target and os.path.isdir(target):
            os.startfile(target)
        else:
            messagebox.showwarning(APP_TITLE, self.tr("invalid_open_target"))

    def append_log(self, text: str):
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")

    def request_stop(self):
        self.stop_requested = True
        self.append_log(self.tr("stop_requested"))

    def _poll_log_queue(self):
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()

                if kind == "log":
                    self.append_log(payload)

                elif kind == "progress":
                    current, total, label = payload
                    self.progress["maximum"] = max(total, 1)
                    self.progress["value"] = current
                    self.progress_label_var.set(label)

                elif kind == "done":
                    success, failed, skipped, log_csv, all_success = payload
                    self.is_running = False
                    self.start_btn.config(state="normal")
                    self.stop_btn.config(state="disabled")
                    if all_success:
                        self.progress_label_var.set(self.tr("success_status"))
                        messagebox.showinfo(APP_TITLE, self.tr("success_message"))
                    else:
                        self.progress_label_var.set(
                            self.tr("done_status", success=success, failed=failed, skipped=skipped)
                        )
                        messagebox.showinfo(
                            APP_TITLE,
                            self.tr("done_message", success=success, failed=failed, skipped=skipped, log_csv=log_csv),
                        )

                elif kind == "error":
                    self.is_running = False
                    self.start_btn.config(state="normal")
                    self.stop_btn.config(state="disabled")
                    self.progress_label_var.set(self.tr("error_status"))
                    messagebox.showerror(APP_TITLE, payload)

        except queue.Empty:
            pass

        self.root.after(150, self._poll_log_queue)

    def start_conversion(self):
        if self.is_running:
            messagebox.showwarning(APP_TITLE, self.tr("already_running"))
            return

        target = self.folder_var.get().strip()
        if not target or not (os.path.isdir(target) or os.path.isfile(target)):
            messagebox.showerror(APP_TITLE, self.tr("invalid_target"))
            return

        if os.path.isfile(target) and Path(target).suffix.lower() not in enabled_extensions():
            messagebox.showerror(APP_TITLE, self.tr("invalid_file"))
            return

        output_formats = self.selected_output_formats()
        if not output_formats:
            messagebox.showerror(APP_TITLE, self.tr("select_output"))
            return

        ok, detail = ensure_pywin32()
        if not ok:
            messagebox.showerror(
                APP_TITLE,
                self.tr("pywin32_missing", detail=detail),
            )
            return

        hwp_processes = get_hwp_processes()
        if hwp_processes:
            process_detail = ", ".join(f"PID {process['pid']}" for process in hwp_processes)
            answer = messagebox.askyesnocancel(
                APP_TITLE,
                self.tr("hwp_running_prompt", process_detail=process_detail),
            )
            if answer is None:
                return
            if answer is True:
                killed = kill_hwp()
                if not killed:
                    messagebox.showerror(
                        APP_TITLE,
                        self.tr("hwp_kill_failed"),
                    )
                    return

        self.stop_requested = False
        self.is_running = True
        self.log_text.delete("1.0", "end")
        lang = self.lang()
        self.append_log(translate(lang, "starting_conversion"))
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.progress["value"] = 0
        self.progress_label_var.set(translate(lang, "scanning"))

        self.worker = threading.Thread(
            target=self._run_conversion,
            args=(
                target,
                self.recursive_var.get(),
                self.overwrite_var.get(),
                self.use_safe_copy_var.get(),
                self.force_one_page_var.get(),
                output_formats,
                lang,
            ),
            daemon=True,
        )
        self.worker.start()

    def selected_output_formats(self):
        formats = []
        if self.output_pdf_var.get():
            formats.append("PDF")
        if self.output_docx_var.get():
            formats.append("DOCX")
        return tuple(formats)

    @staticmethod
    def collect_files(target: str, recursive: bool):
        root = Path(target)
        allowed_extensions = enabled_extensions()
        if root.is_file():
            return [root] if root.suffix.lower() in allowed_extensions else []

        iterator = root.rglob("*") if recursive else root.glob("*")
        files = []
        for path in iterator:
            if path.is_file() and path.suffix.lower() in allowed_extensions:
                files.append(path)
        return files

    def _run_conversion(
        self,
        target: str,
        recursive: bool,
        overwrite: bool,
        use_safe_copy: bool,
        force_one_page: bool,
        output_formats,
        lang: str,
    ):
        try:
            self.log_queue.put(("log", translate(lang, "scanning")))
            target_path = Path(target)
            files = self.collect_files(target, recursive)
            total_files = len(files)
            total_jobs = total_files * len(output_formats)
            extension_label = ", ".join(ext.upper() for ext in enabled_extensions())
            if total_files == 0:
                self.log_queue.put(("error", translate(lang, "no_files", extensions=extension_label)))
                return

            log_root = target_path.parent if target_path.is_file() else target_path
            log_csv = str(log_root / "hwp2pdf_log.csv")
            success = 0
            failed = 0
            skipped = 0
            stopped = False

            import pythoncom
            import win32com.client

            self.log_queue.put(("log", translate(lang, "init_com")))
            pythoncom.CoInitialize()
            hwp = None

            try:
                TEMP_WORKDIR.mkdir(parents=True, exist_ok=True)

                self.log_queue.put(("log", translate(lang, "start_hwp")))
                hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                self.log_queue.put(("log", translate(lang, "hwp_started")))
                try:
                    hwp.XHwpWindows.Item(0).Visible = False
                except Exception:
                    pass

                self.log_queue.put(("log", translate(lang, "register_security")))
                security_ok, security_detail = register_hwp_security_module(hwp)

                on_label = translate(lang, "on")
                off_label = translate(lang, "off")
                self.log_queue.put(("log", translate(lang, "found_files", count=total_files, extensions=extension_label)))
                self.log_queue.put(("log", translate(lang, "csv_log", path=log_csv)))
                self.log_queue.put(("log", translate(lang, "safe_temp_mode", state=on_label if use_safe_copy else off_label)))
                self.log_queue.put(
                    ("log", translate(lang, "force_one_page_mode", state=on_label if force_one_page else off_label))
                )
                self.log_queue.put(("log", translate(lang, "output_formats", formats=", ".join(output_formats))))
                self.log_queue.put(("log", translate(lang, "auto_confirm_docx")))
                security_msg = on_label if security_ok else f"{off_label} ({security_detail or translate(lang, 'module_unavailable')})"
                self.log_queue.put(("log", translate(lang, "security_module", state=security_msg)))

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

                    job_index = 0
                    for file_index, src_path in enumerate(files, start=1):
                        src_path = Path(src_path)
                        fmt = src_path.suffix.replace(".", "").upper()

                        self.log_queue.put(("log", translate(lang, "processing", path=src_path)))

                        for output_format in output_formats:
                            if self.stop_requested:
                                writer.writerow(["STOPPED", "", "", translate(lang, "stopped_csv")])
                                self.log_queue.put(("log", translate(lang, "stopped")))
                                stopped = True
                                break

                            job_index += 1
                            output_ext = output_extension(output_format)
                            output_path = src_path.with_suffix(output_ext)
                            self.log_queue.put(
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
                                    skipped += 1
                                    msg = translate(lang, "skipped_exists", format=output_format)
                                    writer.writerow(["SKIPPED", str(src_path), str(output_path), msg])
                                    self.log_queue.put(
                                        ("log", translate(lang, "skipped_log", format=output_format, path=output_path))
                                    )
                                    self.log_queue.put(
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

                                blocked_reason = blocked_conversion_reason(src_path, output_format, lang)
                                if blocked_reason:
                                    failed += 1
                                    writer.writerow(["FAILED", str(src_path), str(output_path), blocked_reason])
                                    self.log_queue.put(
                                        (
                                            "log",
                                            translate(
                                                lang,
                                                "failed_log",
                                                format=output_format,
                                                path=src_path,
                                                message=blocked_reason,
                                            ),
                                        )
                                    )
                                    self.log_queue.put(
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

                                if use_safe_copy:
                                    temp_input = TEMP_WORKDIR / f"{file_index:05d}_{output_format}_{src_path.name}"
                                    temp_output = TEMP_WORKDIR / f"{file_index:05d}_{output_format}_{src_path.stem}{output_ext}"

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

                                opened = hwp.Open(str(open_target), "", DEFAULT_OPEN_OPTION)
                                if opened is False:
                                    raise RuntimeError(translate(lang, "open_failed", format=fmt))

                                if force_one_page:
                                    force_one_page_view(hwp, lang)

                                actual_save_format = save_document_as(hwp, save_target, output_format, lang)

                                try:
                                    hwp.Clear(1)
                                except Exception:
                                    pass

                                if use_safe_copy:
                                    if not temp_output.exists():
                                        raise RuntimeError(translate(lang, "temp_missing", format=output_format))
                                    shutil.move(str(temp_output), str(output_path))

                                success += 1
                                writer.writerow(["OK", str(src_path), str(output_path), ""])
                                self.log_queue.put(
                                    (
                                        "log",
                                        translate(
                                            lang,
                                            "ok_log",
                                            format=output_format,
                                            actual=actual_save_format,
                                            path=output_path,
                                        ),
                                    )
                                )

                            except Exception as e:
                                failed += 1
                                writer.writerow(["FAILED", str(src_path), str(output_path), str(e)])
                                self.log_queue.put(
                                    (
                                        "log",
                                        translate(
                                            lang, "failed_log", format=output_format, path=src_path, message=e
                                        ),
                                    )
                                )
                                try:
                                    hwp.Clear(1)
                                except Exception:
                                    pass

                            finally:
                                for tmp in (temp_input, temp_output):
                                    try:
                                        if tmp and Path(tmp).exists():
                                            Path(tmp).unlink()
                                    except Exception:
                                        pass

                            self.log_queue.put(
                                (
                                    "progress",
                                    (
                                        job_index,
                                        total_jobs,
                                        translate(lang, "progress_done", current=job_index, total=total_jobs),
                                    ),
                                )
                            )

                        if self.stop_requested:
                            stopped = True
                            break

            finally:
                try:
                    if hwp is not None:
                        hwp.Quit()
                except Exception:
                    pass
                pythoncom.CoUninitialize()

            all_success = success == total_jobs and failed == 0 and skipped == 0 and not stopped
            if all_success:
                try:
                    Path(log_csv).unlink()
                except FileNotFoundError:
                    pass
                except Exception as e:
                    all_success = False
                    self.log_queue.put(("log", translate(lang, "remove_log_failed", message=e)))

            self.log_queue.put(("done", (success, failed, skipped, log_csv, all_success)))

        except Exception as e:
            self.log_queue.put(("error", translate(lang, "unexpected_error", message=e)))


def main():
    root = tk.Tk()
    style = ttk.Style(root)
    try:
        style.theme_use("vista")
    except Exception:
        pass
    ConverterApp(root)
    root.mainloop()
