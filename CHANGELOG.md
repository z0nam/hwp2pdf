# Changelog

이 파일은 [Keep a Changelog](https://keepachangelog.com/) 1.1.0 형식을 느슨하게 따릅니다. 버전 번호는 빌드 스크립트(`scripts/build_windows.ps1`)가 `yyyy.MM.dd.N` 형태로 자동 부여합니다.

각 항목 옆 GitHub 링크는 그 버전의 release 페이지를 가리킵니다 — 다운로드 자산은 거기에 있습니다.

## [Unreleased]

### Added
- HWP/HWPX 파일 여러 개를 한 번에 끌어다 놓거나 파일 선택 창에서 복수 선택할 수
  있습니다. 이후 파일을 추가로 드롭하면 기존 목록에 중복 없이 이어 붙이며, 목록에서
  선택 항목을 제거하거나 전체를 비울 수 있습니다.
- **macOS 앱.** Windows 변환 서버에 연결해 Windows판과 동일한 UI로 HWP/HWPX를
  PDF·DOCX로 일괄 변환합니다. 파일 목록·건너뛰기 판정·CSV 로그·진행률·중지는 mac에
  남고 문서 변환만 원격에서 일어나므로 결과 파일은 평소와 같은 위치에 저장됩니다.
  드래그앤드롭(Apple Silicon 포함)과 한국어/영어 전환도 그대로 동작합니다.
  `hwp2pdf-macos.spec` + `scripts/build_macos.sh`로 ad-hoc 서명된 `.app`을 빌드합니다.
  (공증은 하지 않으므로 첫 실행은 우클릭 → 열기.)
  로컬 한컴오피스 for Mac은 사용하지 않습니다 — AppleScript 사전도 CLI 변환 진입점도
  없어 자동화가 불가능합니다. `docs/known-issues.md` §4 참고.
- **Windows 변환 서버 — `hwp2pdf-cli serve`.** 기존 COM 엔진을 그대로 재사용하는
  표준 라이브러리 HTTP 서버입니다. 토큰 인증(Bearer), 커서 기반 이벤트 롱폴링,
  스트리밍 업로드/다운로드, 단일 워커 직렬 큐, 공유 폴더 경로 전달(경로 탈출 차단),
  선택적 TLS를 지원합니다. `--bind tailscale`은 테일넷 주소에만 바인드해 LAN·공인망
  노출과 방화벽 규칙을 아예 없앱니다.
  ⚠️ Windows 서비스로 등록하지 마세요(Session 0 좀비 `Hwp.exe`). 자동 시작은
  `scripts/install_serve_task.ps1`.
- CLI에 `--server` / `--token` / `--transport` 추가. `hwp2pdf serve`로 서버를 띄웁니다.
  CLI와 서버는 이제 tkinter를 import하지 않습니다.
- 문서: `docs/remote-server.md`(Tailscale·LAN·VM 설정과 문제 해결),
  `docs/protocol.md`(프로토콜 명세). `scripts/smoke_remote.py`로 실제 서버 왕복 검증,
  `scripts/check_macos.sh`로 mac 개발 환경 점검.
- 사용자 설정 저장(`settings.json`) — 서버 주소·토큰·전송 방식과 UI 옵션을
  플랫폼별 표준 위치에 보관합니다. Windows에서는 기존 `update_state.json`과
  같은 `%LOCALAPPDATA%\hwp2pdf\` 폴더를 그대로 씁니다.
- `tests/` 도입과 GitHub Actions 테스트 워크플로(macOS·Windows).
- `scripts/set_version.py` — `yyyy.MM.dd.N` 버전 계산을 플랫폼 중립 스크립트로
  옮겨 Windows·macOS 빌드가 같은 규칙을 씁니다. `build_windows.ps1 -Version`으로
  버전을 고정할 수 있습니다.

### Changed
- 변환 완료 결과는 진행 상태와 로그에 표시하고, 작업을 막는 완료 팝업은 띄우지
  않습니다.
- 파일 목록 모드에서도 창 전체와 목록 영역에 드롭 대상을 등록해 연속 드래그앤드롭을
  받을 수 있도록 개선했습니다. 폴더를 드롭하면 기존처럼 폴더 모드로 전환됩니다.
- 위 다중 파일 선택과 완료 팝업 제거는 **macOS 원격 변환에도 그대로 적용**됩니다.
  대상이 파일 목록인 경우의 처리를 `jobs.collect_files`/`run_batch` 공유 계층으로
  올려서, 로컬 COM 경로와 원격 HTTP 경로가 같은 코드를 씁니다.
- **내부 구조 분리 — Windows 동작 변화 없음.** 2,472줄짜리 `src/hwp2pdf/app.py`를
  `i18n.py`(문구), `constants.py`(포맷·확장자·한컴 상수), `paths.py`(플랫폼 경로),
  `updater.py`(릴리스 확인), `jobs.py`(배치 오케스트레이션),
  `backends/`(변환 엔진 추상화)로 나눴습니다. 파일 검색·스킵/덮어쓰기 판정·CSV
  로그·진행률·중지는 `jobs.run_batch()`가 공통으로 담당하고, 한컴 COM 자동화는
  `backends/windows_com.py`의 `WindowsComBackend`로 옮겼습니다. 옮긴 코드는
  그대로이며 CLI와 기존 진입점(`hwp_pdf_converter_app_safe.py`)도 수정 없이 동작합니다.
  macOS에서 Windows 변환 서버에 붙는 원격 백엔드를 붙일 자리를 만드는 것이 목적입니다.

### Fixed
- **한글 실행 감지가 조용히 죽던 문제.** `tasklist`/`taskkill` 출력을 콘솔 OEM
  코드페이지 대신 UTF-8로 디코드하고 있어서, `PYTHONUTF8=1` 환경의 한국어 Windows에서
  "정보: 실행 중인 작업 중 ..." 메시지를 만나면 subprocess 리더 스레드가 죽고
  `stdout`이 `None`이 됐습니다. 그 결과 `get_hwp_processes()`가 항상 빈 목록을 돌려줘
  "아래한글이 이미 실행 중" 경고와 `--kill-hwp`가 동작하지 않았고, 매번 stderr에
  traceback이 찍혔습니다. 이제 Windows에서는 `encoding="oem", errors="replace"`로
  디코드합니다. (namun-ji 실기 검증 중 발견)
- **한컴오피스 설치 감지.** 서버의 `/v1/capabilities`가 `HKLM\SOFTWARE\HNC\HwpRun`만
  찾다 보니, 64비트 Windows에 32비트 한컴오피스 2022가 깔린 실제 환경(키가
  `WOW6432Node` 아래에 있고 `HwpRun`은 아예 없음)에서 설치를 놓쳤습니다.
  이제 실제로 중요한 표식인 `HWPFrame.HwpObject` ProgID 등록 여부를 봅니다.
- **`hwp2pdf.app` 호환 재노출 누락.** 백엔드 분리 과정에서 `output_extension` 등이
  `hwp2pdf.app`에서 사라져 `scripts/check_windows.ps1`이 깨졌습니다. 분리 이전의
  공개 심볼 전체를 다시 재노출하고 `tests/test_app_surface.py`로 고정했습니다.
- `scripts/install_serve_task.ps1`의 `-LogonType InteractiveToken`은 유효하지 않은
  값이라 작업 등록이 실패했습니다 (`Interactive`가 맞습니다).

## [2026.08.28.1] - 2026-08-28

### Changed
- `변환 시작`과 `중지` 버튼에 재생·정지 아이콘, 녹색·빨간색의 은은한
  그라데이션, hover·눌림·비활성 상태를 적용해 주요 작업을 더 분명하게
  구분하고, `중지` 버튼은 간결한 폭으로 정리했습니다.

## [2026.08.25.1] - 2026-08-25

### Added
- Windows 탐색기에서 HWP/HWPX 파일이나 폴더를 앱 창에 끌어다 놓아 변환 대상으로
  지정하는 기능을 추가했습니다.

### Changed
- 파일을 대상으로 지정하면 `하위 폴더 포함`을 비활성화하고, 다시 폴더를 지정하면
  이전 폴더 모드의 선택값을 복원합니다.
- 포터블 ZIP과 설치본에 드래그앤드롭 런타임 및 제3자 라이선스 고지를 포함하고,
  Windows 사전 점검이 프로젝트 가상환경을 우선 사용하도록 개선했습니다.

## [2026.07.31.3] - 2026-07-31

### Fixed
- 자동 업데이트가 숨김 PowerShell 보조 스크립트에서 멈추며 앱이 재실행되지 않는
  문제를 수정했습니다. 보조 프로세스의 시작 여부와 부모 앱 종료를 확인한 뒤
  설치를 실행하고, 실패 로그를 보존하며 기존 앱을 다시 실행합니다. 설치가
  완료되면 Inno Setup이 새 버전을 원래 사용자 권한으로 실행합니다.
- `2026.07.30.1` 이하 설치본은 기존 자동 업데이트 실행부의 문제 때문에
  이 버전의 설치파일을 한 번 직접 실행해야 할 수 있습니다.

## [2026.07.30.1] - 2026-07-30

### Changed
- 안전한 로컬 임시 폴더 변환 옵션의 설명을
  `구글 드라이브/네트워크 드라이브 사용시 권장`으로 명확히 수정했습니다.

## [2026.06.25.6] - 2026-06-25

### Added
- **앱 내 자동 업데이트** — 24시간마다 GitHub Releases를 확인하다 새 버전이 보이면
  메인 윈도우에 `지금 자동 업데이트` 버튼이 노출됩니다(설치본 한정). 클릭하면
  setup.exe를 `%LOCALAPPDATA%\hwp2pdf\updates\`에 진행률 표시와 함께 다운로드 →
  UAC 한 번 → silent install → 새 버전으로 자동 재시작. 포터블/dev 빌드는 기존
  `최신 버전 다운로드`(브라우저 열기) 흐름 유지. 설치본 여부는 exe 옆의
  `unins000.exe` 마커로 판별. (`84962ad`)

## [2026.06.25.5] - 2026-06-25

### Added
- **한컴 파일접근 보안 모듈 자가등록 (B안)** — 앱 첫 실행 시
  `HKCU\Software\HNC\HwpAutomation\Modules\FilePathCheckerModule` 확인 → 없으면
  한컴 비트수(`Hwp.exe` PE 헤더 + 설치 경로) 자동 감지 후 번들 DLL을
  `%LOCALAPPDATA%\hwp2pdf\security\`에 복사 + 레지스트리 등록. 무인/배치 변환
  시 "접근을 허용하시겠습니까?" 한컴 대화상자가 더 이상 행을 일으키지 않음.
- **자체 작성 stub DLL (vendor/x86, x64)** — `IsAccessiblePath()`가 무조건
  TRUE를 반환하는 MIT 라이선스 stub. 한컴 공개 ZIP 코드 일절 미사용
  (export 시그니처는 ABI라 저작권 대상 아님). MSVC 2022 Build Tools로 직접 빌드.
- **설치프로그램 보강** — `installer/hwp2pdf.iss`에 `[Files]`(x86/x64 DLL) +
  `[Registry]`(HKCU 기본값)로 설치 즉시 자가등록 belt-and-suspenders.
- **재현 fixture** — `docs/fixtures/docx-failure-repro.hwp` — DOCX 변환에서
  RPC_E_SERVERFAULT + `알 수 없는 형식의 파일입니다.` dialog 패턴이 재현되는
  익명화된 hwp.
- **`docs/known-issues.md`** — 후속 과제 3건 명문화.

### Changed
- **한컴 오류 대화상자 자동 처리 강화** — `HancomDialogWatcher`가 PyInstaller
  빌드본에서도 실제 동작하도록 `win32gui/win32con/win32process`를
  `hiddenimports`에 추가 (옛 빌드에서 watcher가 사일런트로 죽었던 것 추정).
- `HANCOM_BLOCKING_DIALOG_MESSAGES`에 `"알 수 없는 형식의 파일입니다."` 추가 →
  자동 확인 + 즉시 실패 처리. 옛 hwp의 DOCX 변환 시 같은 dialog가 6번 떴다
  사라지는 노가다 제거.

### 관련 커밋
- `b2c9660` — 보안모듈 자가등록 + watcher fix

## 이전 버전

`v2026.06.25.5` 이전 history는 `git log` 참고. 이 파일은 v5부터 추적합니다.

---

[Unreleased]: https://github.com/z0nam/hwp2pdf/compare/v2026.08.25.1...HEAD
[2026.08.25.1]: https://github.com/z0nam/hwp2pdf/releases/tag/v2026.08.25.1
[2026.07.31.3]: https://github.com/z0nam/hwp2pdf/releases/tag/v2026.07.31.3
[2026.07.30.1]: https://github.com/z0nam/hwp2pdf/releases/tag/v2026.07.30.1
[2026.06.25.6]: https://github.com/z0nam/hwp2pdf/releases/tag/v2026.06.25.6
[2026.06.25.5]: https://github.com/z0nam/hwp2pdf/releases/tag/v2026.06.25.5
