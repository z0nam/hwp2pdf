"""User-facing strings and the ``translate`` helper.

Extracted verbatim from ``app.py`` so the GUI, the CLI and the conversion
server all render messages from a single table.
"""

LANGUAGE_LABELS = {
    "ko": "한국어",
    "en": "English",
}


LANGUAGE_CODES = {label: code for code, label in LANGUAGE_LABELS.items()}


TEXT = {
    "ko": {
        "target_label": "대상 폴더 또는 파일",
        "drop_hint": "탐색기에서 HWP/HWPX 파일 여러 개 또는 폴더를 이 창에 끌어다 놓을 수 있습니다.",
        "invalid_drop": "HWP/HWPX 파일 또는 폴더를 끌어다 놓으세요.",
        "selected_files": "선택한 파일 ({count}개)",
        "remove_selected": "선택 제거",
        "clear_all": "전체 비우기",
        "file_count_estimate": "변환 예상 파일: 총 {count}개",
        "file_count_unavailable": "변환 예상 파일 수를 확인할 수 없습니다.",
        "browse_folder": "폴더 선택...",
        "pick_file": "파일 선택...",
        "options": "옵션",
        "include_subfolders": "하위 폴더 포함",
        "overwrite": "기존 출력 파일 덮어쓰기",
        "output": "출력",
        "safe_temp": "안전한 로컬 임시 폴더 변환 사용(구글 드라이브/네트워크 드라이브 사용시 권장)",
        "force_one_page": "저장 전 한쪽 보기 강제 적용",
        "start": "변환 시작",
        "stop": "중지",
        "open_selected": "선택 위치 열기",
        "upgrade": "최신 버전 다운로드",
        "auto_update": "지금 자동 업데이트",
        "auto_update_confirm": "v{latest}로 자동 업데이트할까요?\n\n다운로드 후 설치를 시작하며 hwp2pdf가 재시작됩니다.\n관리자 권한 요청이 한 번 표시될 수 있습니다.",
        "auto_update_busy": "변환이 진행 중입니다. 끝나면 다시 시도하세요.",
        "auto_update_portable": "자동 업데이트는 설치본에서만 지원됩니다. 브라우저에서 직접 다운로드할까요?",
        "auto_update_downloading": "업데이트 다운로드 중... {pct}%",
        "auto_update_installing": "설치 중... 잠시 후 hwp2pdf가 재시작됩니다.",
        "auto_update_failed": "자동 업데이트 실패: {error}",
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
        "select_file_title": "변환할 HWP/HWPX 파일 선택 (복수 선택 가능)",
        "all_files": "모든 파일",
        "invalid_target": "올바른 폴더 또는 HWP/HWPX 파일을 선택하세요.",
        "invalid_file": "HWP 또는 HWPX 파일을 선택하세요.",
        "invalid_open_target": "올바른 폴더 또는 파일을 먼저 선택하세요.",
        "already_running": "이미 변환 작업이 실행 중입니다.",
        "select_output": "출력 형식을 하나 이상 선택하세요: PDF 또는 DOCX.",
        "pywin32_missing": "pywin32를 사용할 수 없습니다.\n\n설치 명령:\npython -m pip install pywin32\n\n상세:\n{detail}",
        "backend_requires_windows": "로컬 한컴오피스 변환은 Windows에서만 가능합니다.\n\nmacOS에서는 Windows 변환 서버 주소를 설정하세요.",
        "no_backend": "사용할 수 있는 변환 백엔드가 없습니다.\n\nWindows에서는 한컴오피스를, 그 외 환경에서는 변환 서버 주소를 설정하세요.",
        "server_not_configured": "변환 서버 주소가 설정되지 않았습니다.\n\n설정에서 Windows 변환 서버 주소를 입력하세요.",
        "remote_unreachable": "변환 서버에 연결하지 못했습니다: {url}\n\n상세: {detail}",
        "remote_auth_failed": "변환 서버 인증에 실패했습니다. 토큰을 확인하세요.",
        "remote_version_mismatch": "변환 서버와 버전이 맞지 않습니다(서버 {server}, API {api}). 양쪽을 같은 버전으로 맞추세요.",
        "remote_hwp_missing": "변환 서버에 한컴오피스 한글이 없습니다. {detail}",
        "remote_http_error": "변환 서버 오류 {status}: {detail}",
        "remote_server_busy": "변환 서버가 처리 중인 작업이 많습니다. 잠시 후 다시 시도하세요.",
        "remote_upload_too_large": "파일이 서버의 업로드 상한을 초과했습니다. 공유 폴더 모드를 사용하세요.",
        "remote_upload_failed": "원본 전송 실패: {name}\n\n상세: {detail}",
        "remote_download_failed": "변환 결과를 내려받지 못했습니다: {detail}",
        "remote_output_missing": "공유 폴더에서 변환 결과를 찾지 못했습니다: {path}",
        "remote_format_unsupported": "변환 서버가 지원하지 않는 출력 형식입니다: {format}",
        "remote_connected": "변환 서버 연결: {url} (v{version})",
        "remote_transport": "전송 방식: {mode}",
        "transport_auto": "자동",
        "transport_upload": "업로드",
        "transport_share": "공유 폴더",
        "server_prefix": "서버: {text}",
        "job_timeout_option": "변환이 오래 걸리면 한글을 강제 종료하고 다음 파일로 진행",
        "job_timeout_minutes": "분",
        "job_timeout_remote_note": "원격 변환에서는 서버 설정이 적용됩니다.",
        "job_timeout": "{seconds}초 안에 변환이 끝나지 않아 한글을 강제 종료했습니다. 문서가 너무 크거나 한컴 대화상자에서 멈춘 것일 수 있습니다.",
        "engine_timeout_kill": "변환이 {seconds}초를 넘겨 한글을 강제 종료합니다: {name}",
        "engine_restarted": "한글을 다시 시작했습니다. 다음 파일부터 이어서 변환합니다.",
        "engine_restart_failed": "한글을 다시 시작하지 못했습니다: {detail}",
        "server_section": "변환 서버 (Windows + 한컴오피스)",
        "server_use_remote": "원격 변환 서버 사용",
        "server_address": "주소",
        "server_token": "토큰",
        "server_transport_label": "전송",
        "server_test": "연결 테스트",
        "server_test_running": "확인 중...",
        "server_test_ok": "연결됨 · 서버 v{version} · 한글 {hwp} · 대기열 {queue}",
        "server_test_failed": "연결 실패: {detail}",
        "server_hwp_ok": "있음",
        "server_hwp_missing": "없음",
        "notes_remote": "- 변환은 한컴오피스가 설치된 Windows 변환 서버에서 실행됩니다.\n- 서버에서 `hwp2pdf-cli serve`를 로그인된 데스크톱 세션으로 실행해 두세요.\n- 서버 설정은 docs/remote-server.md의 Tailscale/LAN/VM 안내를 참고하세요.\n- 공유 폴더가 설정돼 있으면 업로드 없이 경로만 전달합니다.\n- 실패, 건너뜀, 중단이 있으면 선택 위치에 CSV 로그가 남습니다.",
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
        "done_log": "변환 완료 — 성공: {success}, 실패: {failed}, 건너뜀: {skipped}",
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
        "target_label": "Target folder or files",
        "drop_hint": "Drag multiple HWP/HWPX files or a folder from File Explorer onto this window.",
        "invalid_drop": "Drop an HWP/HWPX file or folder.",
        "selected_files": "Selected files ({count})",
        "remove_selected": "Remove selected",
        "clear_all": "Clear all",
        "file_count_estimate": "Estimated files to convert: {count}",
        "file_count_unavailable": "Could not estimate the number of files to convert.",
        "browse_folder": "Browse folder...",
        "pick_file": "Pick file...",
        "options": "Options",
        "include_subfolders": "Include subfolders",
        "overwrite": "Overwrite existing output",
        "output": "Output",
        "safe_temp": "Use safe local temp conversion (recommended when using Google Drive / network drives)",
        "force_one_page": "Force one-page view before export",
        "start": "Start conversion",
        "stop": "Stop",
        "open_selected": "Open selected folder",
        "upgrade": "Download latest",
        "auto_update": "Auto-update now",
        "auto_update_confirm": "Auto-update to v{latest}?\n\nThe update will be downloaded, installed, and hwp2pdf will restart.\nYou may see a one-time UAC prompt.",
        "auto_update_busy": "A conversion is running. Try again when it finishes.",
        "auto_update_portable": "Auto-update is only supported for the installed build. Open the browser to download manually?",
        "auto_update_downloading": "Downloading update... {pct}%",
        "auto_update_installing": "Installing... hwp2pdf will restart shortly.",
        "auto_update_failed": "Auto-update failed: {error}",
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
        "select_file_title": "Select HWP/HWPX files to convert (multiple allowed)",
        "all_files": "All files",
        "invalid_target": "Select a valid root folder or HWP/HWPX file.",
        "invalid_file": "Select an HWP or HWPX file.",
        "invalid_open_target": "Select a valid folder or file first.",
        "already_running": "A conversion job is already running.",
        "select_output": "Select at least one output format: PDF or DOCX.",
        "pywin32_missing": "pywin32 is not available.\n\nInstall it with:\npython -m pip install pywin32\n\nDetails:\n{detail}",
        "backend_requires_windows": "Local Hancom Office conversion requires Windows.\n\nOn macOS, configure the address of a Windows conversion server.",
        "no_backend": "No conversion backend is available.\n\nInstall Hancom Office on Windows, or configure a conversion server address.",
        "server_not_configured": "No conversion server address is configured.\n\nEnter a Windows conversion server address in the settings.",
        "remote_unreachable": "Could not reach the conversion server: {url}\n\nDetails: {detail}",
        "remote_auth_failed": "Conversion server authentication failed. Check the token.",
        "remote_version_mismatch": "Version mismatch with the conversion server (server {server}, API {api}). Update both sides to the same version.",
        "remote_hwp_missing": "The conversion server has no Hancom Office Hangul installed. {detail}",
        "remote_http_error": "Conversion server error {status}: {detail}",
        "remote_server_busy": "The conversion server queue is full. Try again shortly.",
        "remote_upload_too_large": "The file exceeds the server upload limit. Use shared folder mode instead.",
        "remote_upload_failed": "Could not upload the source: {name}\n\nDetails: {detail}",
        "remote_download_failed": "Could not download the converted file: {detail}",
        "remote_output_missing": "The converted file was not found in the shared folder: {path}",
        "remote_format_unsupported": "The conversion server does not support this output format: {format}",
        "remote_connected": "Connected to conversion server: {url} (v{version})",
        "remote_transport": "Transport: {mode}",
        "transport_auto": "Automatic",
        "transport_upload": "Upload",
        "transport_share": "Shared folder",
        "server_prefix": "server: {text}",
        "job_timeout_option": "Force-close Hangul and move on if a conversion takes too long",
        "job_timeout_minutes": "min",
        "job_timeout_remote_note": "Remote conversions use the server's own setting.",
        "job_timeout": "Conversion did not finish within {seconds}s, so Hangul was force-closed. The document may be very large, or Hangul may be stuck on a dialog.",
        "engine_timeout_kill": "Conversion exceeded {seconds}s; force-closing Hangul: {name}",
        "engine_restarted": "Hangul was restarted. The batch continues with the next file.",
        "engine_restart_failed": "Could not restart Hangul: {detail}",
        "server_section": "Conversion server (Windows + Hancom Office)",
        "server_use_remote": "Use a remote conversion server",
        "server_address": "Address",
        "server_token": "Token",
        "server_transport_label": "Transport",
        "server_test": "Test connection",
        "server_test_running": "Checking...",
        "server_test_ok": "Connected - server v{version} - Hangul {hwp} - queue {queue}",
        "server_test_failed": "Connection failed: {detail}",
        "server_hwp_ok": "present",
        "server_hwp_missing": "missing",
        "notes_remote": "- Conversion runs on a Windows server with Hancom Office installed.\n- Keep `hwp2pdf-cli serve` running there in a logged-in desktop session.\n- See docs/remote-server.md for Tailscale, LAN and VM setups.\n- When a shared folder is configured, paths are passed instead of uploads.\n- A CSV log is left at the selected location if anything failed, was skipped or stopped.",
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
        "done_log": "Conversion complete — Success: {success}, Failed: {failed}, Skipped: {skipped}",
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
