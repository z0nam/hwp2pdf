import csv
import json
import os
import queue
import shutil
import struct
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from hwp2pdf.version import __version__

APP_NAME = "HWP/HWPX -> PDF/DOCX Converter"
APP_TITLE = f"{APP_NAME} v{__version__}"
GITHUB_RELEASES_API_URL = "https://api.github.com/repos/z0nam/hwp2pdf/releases/latest"
GITHUB_RELEASES_PAGE_URL = "https://github.com/z0nam/hwp2pdf/releases/latest"
UPDATE_CHECK_INTERVAL_SECONDS = 24 * 60 * 60
UPDATE_STATE_PATH = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "hwp2pdf" / "update_state.json"
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
HWP_SECURITY_REG_KEY = r"Software\HNC\HwpAutomation\Modules"
HWP_SECURITY_REG_VALUE = "FilePathCheckerModule"
HWP_SECURITY_DLL_NAME = "FilePathCheckerModule.dll"
HWP_SECURITY_INSTALL_DIR = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "hwp2pdf" / "security"
MESSAGE_BOX_AUTO_CONFIRM = 0x10
HANCOM_BLOCKING_DIALOG_MESSAGES = (
    "복합 파일을 현재 구현하기에 너무 큽니다.",
    "알 수 없는 형식의 파일입니다.",
)
HANCOM_DIALOG_CONFIRM_BUTTONS = ("확인", "OK", "예", "Yes", "계속", "Continue", "닫기", "Close")
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
        "file_count_estimate": "변환 예상 파일: 총 {count}개",
        "file_count_unavailable": "변환 예상 파일 수를 확인할 수 없습니다.",
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
        "upgrade": "최신 버전 다운로드",
        "update_status_checking": "업데이트 확인 중...",
        "update_status_current": "최신 버전입니다. 현재: {current}",
        "update_status_available": "새 버전이 있습니다. 현재: {current} / 최신: {latest}",
        "update_status_no_release": "최신 버전입니다. 현재: {current}",
        "update_status_failed": "현재 버전: {current}. 업데이트 확인 불가",
        "ready": "준비",
        "log": "로그",
        "notes_title": "참고",
        "notes": (
            "- 안정성을 위해 시작 전에 아래한글을 닫아 주세요.\n"
            "- 안전한 임시 폴더 모드는 각 파일을 짧은 로컬 경로로 복사한 뒤 변환합니다.\n"
            "- PDF가 2쪽 보기/모아찍기로 저장되는 문제를 피하려고 기본적으로 한쪽 보기를 강제 적용합니다.\n"
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
        "hwp_closed_already": "HWP 프로세스가 이미 종료되어 계속 진행합니다.",
        "starting_conversion": "변환 시작...",
        "scanning": "파일 검색 중...",
        "no_files": "{extensions} 파일을 찾지 못했습니다.",
        "init_com": "한컴 자동화 초기화 중...",
        "start_hwp": "HWPFrame.HwpObject 시작 중...",
        "hwp_started": "HWPFrame.HwpObject가 시작되었습니다.",
        "register_security": "한컴 파일 접근 보안 모듈 등록 중...",
        "security_self_registered": "보안 모듈 자가등록 완료: {detail}",
        "security_bundle_missing": "번들 보안 모듈 DLL을 찾지 못해 자가등록을 건너뜁니다: {detail}",
        "security_self_register_failed": "보안 모듈 자가등록 실패({state}): {detail}",
        "found_files": "대상 파일 {count}개 발견: {extensions}",
        "csv_log": "CSV 로그: {path}",
        "safe_temp_mode": "안전 임시 폴더 모드: {state}",
        "force_one_page_mode": "한쪽 보기/모아찍기 해제 강제 적용: {state}",
        "nup_print_reset": "기존 인쇄 방식이 '{method}'로 설정되어 있어 PDF 저장 전 '자동 인쇄(1페이지)'로 강제 적용했습니다.",
        "output_formats": "출력 형식: {formats}",
        "auto_confirm_docx": "한컴 확인/오류 대화상자 자동 확인: 켜짐",
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
        "error_log": "오류: {message}",
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
        "pdf_print_method_failed": "PDF 인쇄 방식 초기화에 실패했습니다.",
        "pdf_print_save_failed": "한컴 PDF 인쇄 방식으로 PDF 저장에 실패했습니다.",
        "hancom_dialog_blocked": "한컴 오류 대화상자가 표시되어 해당 파일을 실패 처리했습니다: {message}",
    },
    "en": {
        "target_label": "Root folder or file",
        "file_count_estimate": "Estimated files to convert: {count}",
        "file_count_unavailable": "Could not estimate the number of files to convert.",
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
        "upgrade": "Download latest",
        "update_status_checking": "Checking for updates...",
        "update_status_current": "Up to date. Current: {current}",
        "update_status_available": "New version available. Current: {current} / Latest: {latest}",
        "update_status_no_release": "Up to date. Current: {current}",
        "update_status_failed": "Current: {current}. Update check unavailable",
        "ready": "Ready",
        "log": "Log",
        "notes_title": "Notes",
        "notes": (
            "- Close Hancom HWP before starting for best stability.\n"
            "- Safe temp mode copies each file to a short local path before conversion.\n"
            "- One-page view and PDF print method reset are forced before export by default to avoid two-page PDF output.\n"
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
        "hwp_closed_already": "HWP process is already closed. Continuing.",
        "starting_conversion": "Starting conversion...",
        "scanning": "Scanning files...",
        "no_files": "No {extensions} files found.",
        "init_com": "Initializing Hancom COM automation...",
        "start_hwp": "Starting HWPFrame.HwpObject...",
        "hwp_started": "HWPFrame.HwpObject started.",
        "register_security": "Registering HWP file access security module...",
        "security_self_registered": "Security module self-registered: {detail}",
        "security_bundle_missing": "Bundled security module DLL not found; skipping self-registration: {detail}",
        "security_self_register_failed": "Security module self-registration failed ({state}): {detail}",
        "found_files": "Found {count} file(s): {extensions}",
        "csv_log": "CSV log: {path}",
        "safe_temp_mode": "Safe temp mode: {state}",
        "force_one_page_mode": "Force one-page view / reset N-up printing: {state}",
        "nup_print_reset": "Existing print method was '{method}', so it was forced to 'Automatic print (one page)' before PDF export.",
        "output_formats": "Output formats: {formats}",
        "auto_confirm_docx": "Auto-confirm Hancom confirmation/error dialogs: ON",
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
        "error_log": "ERROR: {message}",
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
        "pdf_print_method_failed": "PDF print method reset failed",
        "pdf_print_save_failed": "PDF export through Hancom PDF printing failed",
        "hancom_dialog_blocked": "A Hancom error dialog appeared, so this file was marked as failed: {message}",
    },
}


def translate(lang: str, key: str, **kwargs):
    text = TEXT.get(lang, TEXT["ko"]).get(key, TEXT["ko"].get(key, key))
    return text.format(**kwargs) if kwargs else text


def parse_version(value: str):
    parts = []
    for part in value.strip().lstrip("vV").split("."):
        try:
            parts.append(int(part))
        except ValueError:
            break
    return tuple(parts)


def latest_release_version(release: dict):
    for key in ("tag_name", "name"):
        value = str(release.get(key) or "").strip()
        parsed = parse_version(value)
        if parsed:
            return value.lstrip("vV")
    return ""


def latest_release_download_url(release: dict):
    assets = release.get("assets") or []
    if not isinstance(assets, list):
        return ""

    candidates = []
    for asset in assets:
        if not isinstance(asset, dict):
            continue
        name = str(asset.get("name") or "").lower()
        url = str(asset.get("browser_download_url") or "").strip()
        if name and url:
            candidates.append((name, url))

    preferred_patterns = (
        ("setup", ".exe"),
        ("windows", ".zip"),
        (".exe",),
        (".zip",),
    )
    for pattern in preferred_patterns:
        for name, url in candidates:
            if all(part in name for part in pattern):
                return url
    return candidates[0][1] if candidates else ""


def fetch_latest_release():
    request = urllib.request.Request(
        GITHUB_RELEASES_API_URL,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": f"hwp2pdf/{__version__}",
        },
    )
    with urllib.request.urlopen(request, timeout=10) as response:
        return json.loads(response.read().decode("utf-8"))


def load_update_state():
    try:
        with UPDATE_STATE_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_update_state(state: dict):
    try:
        UPDATE_STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
        with UPDATE_STATE_PATH.open("w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def should_check_updates(state: dict):
    try:
        checked_at = float(state.get("checked_at", 0))
    except (TypeError, ValueError):
        checked_at = 0
    return time.time() - checked_at >= UPDATE_CHECK_INTERVAL_SECONDS


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


PRINT_METHOD_LABELS = {
    "ko": {
        0: "자동 인쇄",
        1: "공급 용지에 맞추어",
        2: "나눠 찍기",
        3: "자동으로 모아 찍기",
        4: "2쪽씩 모아 찍기",
        5: "3쪽씩 모아 찍기",
        6: "4쪽씩 모아 찍기",
        7: "6쪽씩 모아 찍기",
        8: "8쪽씩 모아 찍기",
        9: "9쪽씩 모아 찍기",
        10: "16쪽씩 모아 찍기",
    },
    "en": {
        0: "Automatic print",
        1: "Fit to paper",
        2: "Tile pages",
        3: "Automatic N-up printing",
        4: "2 pages per sheet",
        5: "3 pages per sheet",
        6: "4 pages per sheet",
        7: "6 pages per sheet",
        8: "8 pages per sheet",
        9: "9 pages per sheet",
        10: "16 pages per sheet",
    },
}


def print_method_label(print_method, lang: str):
    labels = PRINT_METHOD_LABELS.get(lang, PRINT_METHOD_LABELS["ko"])
    return labels.get(print_method, f"PrintMethod={print_method}")


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


def _resource_root() -> Path:
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base)
    return Path(__file__).resolve().parent.parent.parent


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
        self.file_count_var = tk.StringVar()
        self.update_status_var = tk.StringVar()
        self.upgrade_btn = None
        self.latest_release_url = GITHUB_RELEASES_PAGE_URL
        self.latest_download_url = GITHUB_RELEASES_PAGE_URL
        self.update_check_running = False
        self.ui = {}

        self._build_ui()
        self.folder_var.trace_add("write", self._on_target_path_changed)
        self.recursive_var.trace_add("write", self._on_target_path_changed)
        self._apply_cached_update_state()
        self._poll_log_queue()
        self.root.after(1000, self.check_for_updates_if_due)

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
        self.ui["browse_btn"].grid(row=0, column=1, sticky="e", padx=(0, 8))
        self.ui["pick_btn"] = ttk.Button(path_row, command=self.pick_file_folder)
        self.ui["pick_btn"].grid(row=0, column=2, sticky="e")
        self.ui["file_count_label"] = ttk.Label(top, textvariable=self.file_count_var)
        self.ui["file_count_label"].grid(row=2, column=0, sticky="w", pady=(4, 0))

        update_row = ttk.Frame(top)
        update_row.grid(row=3, column=0, sticky="w", pady=(4, 0))
        self.ui["update_status_label"] = ttk.Label(update_row, textvariable=self.update_status_var)
        self.ui["update_status_label"].pack(side="left")
        self.upgrade_btn = ttk.Button(update_row, command=self.open_latest_release)
        self.upgrade_btn.pack(side="left", padx=(8, 0))
        self.upgrade_btn.pack_forget()

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
        self.log_text.tag_configure("error", foreground="#b00020")
        self.log_text.tag_configure("warning", foreground="#8a5a00")
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
        self.upgrade_btn.configure(text=self.tr("upgrade"))
        self.ui["log_label"].configure(text=self.tr("log"))
        self.ui["note_frame"].configure(text=self.tr("notes_title"))
        self.ui["notes_label"].configure(text=self.tr("notes"))
        if not self.is_running:
            self.progress_label_var.set(self.tr("ready"))
        self._apply_cached_update_state()
        self._update_file_count_estimate()

    def browse_folder(self):
        initial_dir = self.folder_var.get().strip()
        if initial_dir and os.path.isfile(initial_dir):
            initial_dir = str(Path(initial_dir).parent)
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = str(Path.home())

        folder = filedialog.askdirectory(
            parent=self.root, title=self.tr("select_folder_title"), initialdir=initial_dir
        )
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
            if self.recursive_var.get():
                self.recursive_var.set(False)
            self.recursive_check.state(["disabled"])
        else:
            self.recursive_check.state(["!disabled"])
        self._update_file_count_estimate()

    def _update_file_count_estimate(self):
        if not hasattr(self, "file_count_var"):
            return

        target = self.folder_var.get().strip()
        if not target:
            self.file_count_var.set("")
            return

        try:
            if os.path.isfile(target):
                count = 1 if Path(target).suffix.lower() in enabled_extensions() else 0
            elif os.path.isdir(target):
                count = len(self.collect_files(target, self.recursive_var.get()))
            else:
                self.file_count_var.set("")
                return
        except Exception:
            self.file_count_var.set(self.tr("file_count_unavailable"))
            return

        self.file_count_var.set(self.tr("file_count_estimate", count=count))

    def open_selected_folder(self):
        target = self.folder_var.get().strip()
        if target and os.path.isfile(target):
            os.startfile(str(Path(target).parent))
        elif target and os.path.isdir(target):
            os.startfile(target)
        else:
            messagebox.showwarning(APP_TITLE, self.tr("invalid_open_target"))

    def append_log(self, text: str, level: str = "info"):
        tag = level if level in {"error", "warning"} else None
        if tag:
            self.log_text.insert("end", text + "\n", tag)
        else:
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
                    if isinstance(payload, tuple):
                        text, level = payload
                        self.append_log(text, level)
                    else:
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
                    self.append_log(self.tr("error_log", message=payload), "error")
                    messagebox.showerror(APP_TITLE, payload)

                elif kind == "update_done":
                    if len(payload) == 5:
                        status, latest, release_url, download_url, error_message = payload
                    else:
                        status, latest, release_url, error_message = payload
                        download_url = release_url
                    self.update_check_running = False
                    state = {
                        "checked_at": time.time(),
                        "status": status,
                        "latest": latest,
                        "release_url": release_url,
                        "download_url": download_url,
                        "error": error_message,
                    }
                    save_update_state(state)
                    self._apply_update_state(state)

        except queue.Empty:
            pass

        self.root.after(150, self._poll_log_queue)

    def open_latest_release(self):
        webbrowser.open(self.latest_download_url or self.latest_release_url or GITHUB_RELEASES_PAGE_URL)

    def _show_upgrade_button(self, visible: bool):
        if visible:
            if not self.upgrade_btn.winfo_manager():
                self.upgrade_btn.pack(side="left", padx=(8, 0))
        else:
            self.upgrade_btn.pack_forget()

    def _apply_cached_update_state(self):
        state = load_update_state()
        if state:
            self._apply_update_state(state)
        else:
            self.update_status_var.set(self.tr("update_status_current", current=__version__))
            self._show_upgrade_button(False)

    def _apply_update_state(self, state: dict):
        status = state.get("status")
        latest = state.get("latest") or ""
        release_url = state.get("release_url") or GITHUB_RELEASES_PAGE_URL
        download_url = state.get("download_url") or release_url
        self.latest_release_url = release_url
        self.latest_download_url = download_url

        if status == "newer" and latest and parse_version(latest) > parse_version(__version__):
            self.update_status_var.set(self.tr("update_status_available", current=__version__, latest=latest))
            self._show_upgrade_button(True)
        elif status == "no_release":
            self.update_status_var.set(self.tr("update_status_no_release", current=__version__))
            self._show_upgrade_button(False)
        elif status == "error":
            self.update_status_var.set(self.tr("update_status_failed", current=__version__))
            self._show_upgrade_button(False)
        else:
            self.update_status_var.set(self.tr("update_status_current", current=__version__))
            self._show_upgrade_button(False)

    def check_for_updates_if_due(self):
        state = load_update_state()
        if self.update_check_running or not should_check_updates(state):
            return

        self.update_check_running = True
        self.update_status_var.set(self.tr("update_status_checking"))
        self._show_upgrade_button(False)
        threading.Thread(target=self._check_for_updates_worker, daemon=True).start()

    def _check_for_updates_worker(self):
        try:
            release = fetch_latest_release()
            latest = latest_release_version(release)
            release_url = release.get("html_url") or GITHUB_RELEASES_PAGE_URL
            download_url = latest_release_download_url(release) or release_url
            if latest and parse_version(latest) > parse_version(__version__):
                self.log_queue.put(("update_done", ("newer", latest, release_url, download_url, "")))
            else:
                self.log_queue.put(("update_done", ("current", latest or __version__, release_url, download_url, "")))
        except urllib.error.HTTPError as e:
            if e.code == 404:
                self.log_queue.put(("update_done", ("no_release", "", "", "", "")))
            else:
                self.log_queue.put(("update_done", ("error", "", "", "", str(e))))
        except Exception as e:
            self.log_queue.put(("update_done", ("error", "", "", "", str(e))))

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
                hwp_processes = get_hwp_processes()
                if not hwp_processes:
                    self.append_log(self.tr("hwp_closed_already"))
                elif not kill_hwp():
                    hwp_processes = get_hwp_processes()
                    if not hwp_processes:
                        self.append_log(self.tr("hwp_closed_already"))
                    else:
                        messagebox.showerror(
                            APP_TITLE,
                            self.tr("hwp_kill_failed"),
                        )
                        return
                else:
                    hwp_processes = get_hwp_processes()
                    if hwp_processes:
                        process_detail = ", ".join(f"PID {process['pid']}" for process in hwp_processes)
                        messagebox.showerror(
                            APP_TITLE,
                            self.tr("hwp_kill_failed") + f"\n\n{process_detail}",
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
                global_message_box_mode = None
                dialog_watcher = None
                try:
                    hwp.XHwpWindows.Item(0).Visible = False
                except Exception:
                    pass
                try:
                    _enabled, global_message_box_mode, _detail = enable_auto_confirm_message_boxes(hwp)
                except Exception:
                    global_message_box_mode = None
                try:
                    dialog_watcher = HancomDialogWatcher(hwp_process_id(hwp))
                    dialog_watcher.start()
                except Exception:
                    dialog_watcher = None

                self.log_queue.put(("log", translate(lang, "register_security")))
                self_register_state, self_register_detail = ensure_hwp_security_module_registered()
                if self_register_state == "registered":
                    self.log_queue.put(
                        ("log", translate(lang, "security_self_registered", detail=self_register_detail))
                    )
                elif self_register_state == "bundled-missing":
                    self.log_queue.put(
                        (
                            "log",
                            (
                                translate(lang, "security_bundle_missing", detail=self_register_detail),
                                "warning",
                            ),
                        )
                    )
                elif self_register_state.startswith("error"):
                    self.log_queue.put(
                        (
                            "log",
                            (
                                translate(
                                    lang,
                                    "security_self_register_failed",
                                    state=self_register_state,
                                    detail=self_register_detail,
                                ),
                                "warning",
                            ),
                        )
                    )
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
                                self.log_queue.put(("log", (translate(lang, "stopped"), "warning")))
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
                            dialog_marker = dialog_watcher.mark() if dialog_watcher else 0

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
                                        self.log_queue.put(
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

                                    output_path.unlink()

                                blocked_reason = blocked_conversion_reason(src_path, output_format, lang)
                                if blocked_reason:
                                    failed += 1
                                    writer.writerow(["FAILED", str(src_path), str(output_path), blocked_reason])
                                    self.log_queue.put(
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

                                if force_one_page and output_format == "PDF":
                                    actual_save_format, original_print_method = save_pdf_with_print_to_pdf(
                                        hwp, save_target, lang
                                    )
                                    if is_nup_print_method(original_print_method):
                                        self.log_queue.put(
                                            (
                                                "log",
                                                translate(
                                                    lang,
                                                    "nup_print_reset",
                                                    method=print_method_label(original_print_method, lang),
                                                ),
                                            )
                                        )
                                else:
                                    actual_save_format = save_document_as(hwp, save_target, output_format, lang)

                                if dialog_watcher:
                                    blocking_message = dialog_watcher.blocking_message_since(dialog_marker)
                                    if blocking_message:
                                        raise RuntimeError(
                                            translate(lang, "hancom_dialog_blocked", message=blocking_message)
                                        )

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
                                failure_message = str(e)
                                if dialog_watcher:
                                    blocking_message = dialog_watcher.blocking_message_since(dialog_marker)
                                    if blocking_message:
                                        failure_message = translate(
                                            lang, "hancom_dialog_blocked", message=blocking_message
                                        )
                                writer.writerow(["FAILED", str(src_path), str(output_path), failure_message])
                                self.log_queue.put(
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
                    if "dialog_watcher" in locals() and dialog_watcher is not None:
                        dialog_watcher.stop()
                except Exception:
                    pass
                try:
                    if hwp is not None and "global_message_box_mode" in locals():
                        restore_message_box_mode(hwp, global_message_box_mode)
                except Exception:
                    pass
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
                    self.log_queue.put(("log", (translate(lang, "remove_log_failed", message=e), "warning")))

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
