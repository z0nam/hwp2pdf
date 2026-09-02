# Changelog

이 파일은 [Keep a Changelog](https://keepachangelog.com/) 1.1.0 형식을 느슨하게 따릅니다. 버전 번호는 빌드 스크립트(`scripts/build_windows.ps1`)가 `yyyy.MM.dd.N` 형태로 자동 부여합니다.

각 항목 옆 GitHub 링크는 그 버전의 release 페이지를 가리킵니다 — 다운로드 자산은 거기에 있습니다.

## [Unreleased]

### Added
- 서버 주소를 몰라도 붙을 수 있는 경로를 추가했습니다.
  - **초대 문자열**: 서버가 시작할 때 로그에 `hwp2pdf://...` 한 줄을 찍습니다.
    주소와 토큰이 같이 들어 있어 클라이언트 주소칸에 붙여넣기만 하면 둘 다
    채워집니다. CLI `--server`도 같은 문자열을 받습니다.
  - **찾기 버튼**: 주소칸 옆에서 Tailscale 피어와 이 컴퓨터의 ARP 이웃에
    `/v1/health`를 던져 응답한 서버를 목록으로 보여줍니다. 서버 쪽 변경이
    필요 없어 기존 버전의 서버도 발견됩니다. CLI는 `hwp2pdf find`.

### Changed
- 서버 주소칸이 이름·IP·`호스트:포트`를 모두 받아 `http://호스트:17650`으로
  보정합니다. 단일 이름은 `이름.local`로도 자동 재시도하므로, mac에서 Windows PC
  이름만 알아도 연결됩니다. 스킴을 직접 쓴 URL은 그대로 둡니다.

### Fixed
- macOS에서 `use_remote_backend()`가 항상 원격을 뜻하는데도 Windows 전용 분기를
  단정하던 GUI 테스트를 고쳤습니다. macOS CI가 이 때문에 실패하고 있었습니다.

## [2026.09.02.2] - 2026-09-02

### Changed
- 파일 여러 개를 직접 선택하거나 끌어다 놓은 뒤 변환하면, 요청한 모든 출력 형식의
  변환에 성공한 파일은 목록에서 자동으로 제거합니다. 실패·건너뜀·중지된 파일은
  재시도할 수 있도록 목록에 남깁니다.

## [2026.09.02.1] - 2026-09-02

### Changed
- 로컬 GUI의 변환 타임아웃 기본값을 **30분에서 10분**으로 낮췄습니다. 타임아웃
  체크박스는 계속 기본 해제이며, 큰 문서는 필요에 따라 분 단위로 늘릴 수 있습니다.
  기존에 저장된 사용자 설정의 시간은 그대로 유지됩니다.
- rhwp를 macOS의 변환 서버 연결 실패 전용 옵션에서 **한컴 변환 엔진의 범용 비상
  대체 수단**으로 확장했습니다. Windows에서도 한컴오피스가 없거나 COM 엔진을 시작할
  수 없을 때 선택적으로 rhwp PDF 변환을 사용합니다. 실행 중인 아래한글이 있으면
  `rhwp로 PDF 변환 / 아래한글 종료 후 변환 / 그대로 한컴 사용 / 취소`를 명시적으로
  선택할 수 있습니다.
- GUI가 로컬/원격 환경에 맞는 rhwp 설명과 설치 상태를 표시합니다. rhwp가 없으면 옵션을
  비활성화하고, 공식 Windows/macOS 빌드에는 플랫폼별 rhwp 바이너리를 포함합니다.
- **변환 엔진 상태**를 추가했습니다. Windows에서는 한컴오피스 한글 설치 여부와
  아래한글 실행 상태를 약 2.5초마다 갱신합니다. macOS와 한컴 미설치 Windows에서는
  해당하지 않는 실행 상태를 표시하지 않습니다.
- rhwp 대체는 사전 점검뿐 아니라 첫 변환 세션을 열지 못한 경우에도 적용합니다. 인증
  실패와 서버 프로토콜 불일치는 설정 오류를 숨기지 않도록 대체하지 않으며, 파일별 변환
  도중에는 엔진을 바꾸지 않습니다.

## [2026.09.01.1] - 2026-09-01

### Added
- **서버 불가 시 rhwp 대체 변환** (기본 꺼짐). 변환 서버에 연결할 수 없을 때 로컬
  rhwp로 PDF를 만듭니다. 옵션의 체크박스 또는 CLI `--rhwp-fallback`으로 켭니다.
  **rhwp는 한컴 엔진이 아니라 별도 렌더러이고 결과가 원본과 다릅니다** — 390쪽 픽스처
  실측에서 쪽수가 379 대 390으로 갈렸고, 머리말·목차 쪽번호가 누락되고 일부 대시
  글리프가 깨졌습니다. 급히 내용을 봐야 할 때 쓰고, 배포용 최종본은 서버 복구 후 다시
  변환하세요. DOCX는 만들 수 없습니다.
  대체가 쓰이면 로그에 서버 실패 이유와 품질 경고가 남고 성공 줄이 `성공 PDF (rhwp)`로
  기록돼 나중에도 구분됩니다.
- `scripts/fetch_rhwp.sh` / `.ps1` — rhwp 릴리스 바이너리를 내려받아 체크섬 검증 후
  `vendor/rhwp/`에 설치합니다. `curl`만 쓰므로 **GitHub 계정도 `gh` CLI도 필요 없습니다.**
  바이너리는 레포에 커밋하지 않습니다(플랫폼 종속·10MB). 버전은 `v0.8.4`로 핀.
- **멈춘 변환 복구.** 한 파일이 한컴 대화상자에 걸려 배치 전체가 멈추는 문제
  (`docs/known-issues.md` §1)를 우회할 수 있습니다. 지정한 시간을 넘기면 워치독이
  한글을 강제 종료하고 엔진을 다시 띄워 다음 파일부터 이어서 변환합니다.
  로컬은 `--timeout <초>`로 선택 적용(기본 꺼짐 — 큰 문서는 원래 몇 분씩 걸립니다),
  서버는 `--job-timeout`으로 기본 900초 적용(멈춘 작업이 모든 클라이언트를 막으므로).
  13분 넘게 멈추던 실제 fixture로 검증했습니다: 90초에 정리, 다음 파일 정상 변환.
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
- **서버 기본 포트를 8765에서 17650으로 옮겼습니다.** 8765는 흔한 로컬 개발 서버 포트라
  실제로 충돌했습니다. 17650은 붐비는 8xxx/9xxx 대역 밖이고 Windows 임시 포트 범위
  (49152+) 아래라 OS가 임의 할당하지도 않습니다.
- 포트가 이미 사용 중이면 **점유 프로세스 이름과 함께 명확한 메시지를 로그에 남기고**
  종료합니다. 이전에는 창 없는 빌드에서 흔적 없이 죽었습니다. 이전 인스턴스가 내려가는
  중일 수 있으므로 몇 초간 재시도한 뒤 판단합니다.
- GUI 옵션에 **변환 타임아웃**을 노출했습니다(기본 꺼짐). 켜면 지정한 분을 넘긴 변환에서
  한글을 강제 종료하고 다음 파일로 넘어갑니다. 원격 변환일 때는 서버 설정이 적용되므로
  비활성화되고 그 사실을 안내합니다.
- `docs/remote-server.md`에 Tailscale 무인 모드(`tailscale set --unattended=true`)를
  설치 단계로 넣었습니다. 이게 없으면 재부팅 후 로그인하지 않았을 때 노드까지 오프라인이
  되어 원격 접근 자체가 끊깁니다. 세션이 없을 때의 대비 순서와, 고전 자동 로그인을
  권하지 않는 이유(Windows Hello 기기에서는 TPM 보호를 꺼야 함)도 함께 정리했습니다.
- **변환 서버를 창 없는 `hwp2pdf-serve.exe`로 분리했습니다.** 콘솔 빌드는 로그온
  세션에 검은 창을 띄우는데, 그 창을 닫으면 서버가 조용히 죽어(`0xC000013A`) 맥에서
  연결 실패가 날 때까지 알 수 없었습니다. 이제 창이 아예 없고 출력은
  `%LOCALAPPDATA%\hwp2pdf\server.log`(2MB 롤링)로 갑니다. `--log-file`로 위치를
  바꿀 수 있고, 작업 스케줄러 등록 시 재시작 설정도 함께 들어갑니다.
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
- **`--bind tailscale`이 부팅 직후 실패하던 문제.** 로그온 시 자동 시작되는 서버가
  Tailscale이 연결되기 전에 실행되면 주소를 못 찾고 즉시 종료(exit 1)했습니다.
  로그 파일이 만들어지기 전에 죽어서 흔적도 남지 않았고, 증상은 "맥에서 연결 실패"뿐이라
  권장 설정대로 구성하면 **재부팅할 때마다 서버가 안 뜨는** 상태였습니다. 이제 최대
  3분(`--tailscale-wait`) 기다리며, 대기 사실도 로그에 남깁니다.
- 작업 스케줄러 등록에 트리거를 추가했습니다 — 로그온·시스템 시작·10분 주기 재확인.
  마지막 것이 자가 치유 역할을 하고, `MultipleInstances=IgnoreNew`라 이미 떠 있으면
  아무 일도 하지 않습니다.
- **변환 서버가 엔진 세션을 여러 개 열 수 있던 문제.** 한컴 엔진은 프로세스당 하나인데
  작업마다 `WindowsComBackend`를 열고 있어서, 두 작업이 겹치면 한쪽의 `Quit()`이 다른
  쪽 세션을 끊을 수 있었습니다. 이제 워커가 세션 하나만 소유하고 작업이 바뀔 때 이전
  세션을 먼저 닫습니다.
- **서버 종료 시 `Hwp.exe`가 남던 문제.** 워커를 먼저 join한 뒤 정리 작업을 큐에 넣어서
  실행할 스레드가 없었습니다. 이제 워커가 자기 스레드에서(= COM 아파트먼트를 소유한
  스레드에서) 종료 직전 세션을 닫습니다.
- **Finder에서 HWP를 열면 GUI 대신 무음 변환이 돌던 문제.** `.app`이 문서 타입을
  선언하고 argv emulation을 쓰므로 Finder가 경로를 `sys.argv`로 넘기는데, 진입점이
  이를 CLI 호출로 해석했습니다. 이제 인자가 전부 존재하는 HWP/HWPX 파일이면 그 파일들이
  선택된 상태로 창을 엽니다.
- 폴더 스캔 결과를 정렬합니다. `rglob` 순서는 파일시스템마다 달라 배치 순서와 CSV
  로그가 실행마다 달라질 수 있었습니다. 명시적으로 고른 파일 목록은 고른 순서를 유지합니다.
- GitHub Actions가 `requirements.txt`를 설치하지 않아 Windows에서 pywin32 없이 돌았고,
  COM 기반 테스트가 "사용 불가" 경로로 조용히 통과했습니다.
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

[Unreleased]: https://github.com/z0nam/hwp2pdf/compare/v2026.09.01.1...HEAD
[2026.09.01.1]: https://github.com/z0nam/hwp2pdf/releases/tag/v2026.09.01.1
[2026.08.28.1]: https://github.com/z0nam/hwp2pdf/releases/tag/v2026.08.28.1
[2026.08.25.1]: https://github.com/z0nam/hwp2pdf/releases/tag/v2026.08.25.1
[2026.07.31.3]: https://github.com/z0nam/hwp2pdf/releases/tag/v2026.07.31.3
[2026.07.30.1]: https://github.com/z0nam/hwp2pdf/releases/tag/v2026.07.30.1
[2026.06.25.6]: https://github.com/z0nam/hwp2pdf/releases/tag/v2026.06.25.6
[2026.06.25.5]: https://github.com/z0nam/hwp2pdf/releases/tag/v2026.06.25.5
