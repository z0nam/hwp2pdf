# 알려진 한계 / 후속 과제

이번 변경에서 우회하거나 부분만 잡은 항목들. 일반 사용 흐름에서 발생률이 낮아 보류 중.

## 1. 한컴 "알 수 없는 형식의 파일입니다." dialog 자동 처리 — 미해결, 단 복구는 가능

> **2026-08-28 갱신.** namun-ji(한컴오피스 2022)에서 재현했습니다. 아래 fixture의
> DOCX 변환이 **13분 넘게 멈춥니다** — `HancomDialogWatcher`가 이 대화상자를 자동
> 확인하지 못합니다. 소스 실행이라 2026.06.25.2의 hiddenimports 수정과는 무관하며,
> 근본 원인은 여전히 미해결입니다.
>
> 다만 **배치가 통째로 멈추는 문제는 해결됐습니다.** `--timeout <초>`(로컬 CLI) 또는
> 서버의 `--job-timeout`(기본 900초)을 쓰면 워치독이 한글을 강제 종료하고
> 엔진을 다시 띄워 다음 파일부터 이어서 변환합니다. 같은 fixture로 검증:
> 90초에 정리 → 해당 파일만 실패 기록 → 다음 파일 정상 변환, 전체 99초.
> GUI에는 아직 이 옵션이 노출돼 있지 않습니다.


`HancomDialogWatcher`의 자동-확인 + blocking-message 즉시 실패 처리에 메시지 추가
(`src/hwp2pdf/app.py` `HANCOM_BLOCKING_DIALOG_MESSAGES`).

옛 빌드(<= 2026.06.25.1)는 PyInstaller spec에 `win32gui/win32con/win32process`가
hiddenimports로 없어 watcher 자체가 import 실패로 사일런트 종료된 것으로 추정.
2026.06.25.2부터 hiddenimports 추가 + 메시지 패턴 추가. 새 빌드에서 watcher가
실제로 자동 클릭하는지는 동일 케이스(예:
`(2002)동북아시아 공동 평화와 번영.hwp` 같은 옛 hwp의 DOCX 변환)로 재확인 필요.

자동 처리 실패 시 fallback 패턴 — Hancom dialog title/class 외에 다른 시그니처
매칭, 또는 SaveAs alias 첫 RPC_E_SERVERFAULT(`-2147417851`) 후 즉시 모든 alias
포기하는 short-circuit. 후자는 정상 fall-through 케이스 영향 검토 필요.

### 재현 fixture

`docs/fixtures/docx-failure-repro.hwp` — DOCX 변환은 한컴 OOXML 내보내기에서
RPC_E_SERVERFAULT + `알 수 없는 형식의 파일입니다.` dialog 패턴이 그대로 떨어지는
파일. PDF 변환은 같은 입력에서 정상 성공. watcher fix 검증 시 이 파일로 재현 가능.
(원본 식별 정보는 익명화. 한컴오피스 자동화의 OOXML 한계 검증 용도로만 사용.)

## 2. 세션 0 좀비 Hwp.exe — UAC elevation 재시도 미구현

Hwp.exe가 Windows Session 0(서비스 세션)에 떠 있으면 콘솔 세션(1+)의 일반 사용자
권한으로 `taskkill /IM Hwp.exe /F`가 access denied. 사용자가 작업관리자로 수동 종료
필요.

발생 조건은 SSH 세션이나 서비스 컨텍스트에서 COM dispatch가 hang → 좀비 남기는
특수 경로. **일반 GUI 사용 흐름에서는 같은 콘솔 세션에서 Hwp가 생성/종료되므로
재현되지 않음.** namun-ji 검증 과정의 우리 절차 부산물.

해결안: `kill_hwp` 실패 시 `ShellExecute "runas"` verb로 elevated `taskkill /PID <pid> /F`
재시도 (UAC 프롬프트 1회). EDR/GPO 환경에서는 elevation 자체 차단 가능성 있어
무조건 elevated로 띄우진 않음.

## 3. installer PrivilegesRequired vs HKCU 불일치 (Inno 경고)

`installer/hwp2pdf.iss`는 `PrivilegesRequired` 미명시 → InnoSetup 기본값 `admin`.
설치 시 UAC elevation → admin 토큰의 HKCU에 자가등록 키가 쓰임. 본인이 admin이면
elevation 후 HKCU도 동일 hive라 동작 무리 없음. Over-the-shoulder(다른 계정 admin)
설치 시에는 일반 사용자의 HKCU에 안 박혀, 첫 launch 때 `ensure_hwp_security_module_registered()`
fallback이 보정해줌. 깔끔히 하려면 `PrivilegesRequired=lowest` + `DefaultDirName={userpf}\hwp2pdf`로
per-user 설치 전환. 기존 설치 폴더/upgrade 흐름이 바뀌어 보류.

## 3. 원격(변환 서버) 모드의 알려진 한계

### 대화형 세션 필수

서버는 반드시 **로그인된 데스크톱 세션**에서 실행해야 합니다. Windows 서비스로
등록하면 위 §2의 Session 0 좀비 `Hwp.exe` 문제가 그대로 재현됩니다.
자동 시작은 `scripts/install_serve_task.ps1`(작업 스케줄러, "사용자가 로그온한
경우에만 실행")을 사용하세요.

### 사전 차단 판정이 서버에서 일어남

암호 문서·배포용 문서 판정(`read_hwp_file_flags`)은 Windows 전용 OLE API
(`pythoncom.StgOpenStorage`)를 쓰므로 서버에서 수행합니다. 따라서 원격 모드에서는
차단 대상 파일도 **일단 업로드된 뒤** 실패로 보고됩니다(공유 폴더 모드에서는 전송 없음).
로컬 Windows 모드처럼 전송 전에 걸러내지는 못합니다.

### 큐 직렬화

한컴 COM 엔진은 프로세스·스레드 단일 인스턴스라 서버는 워커 스레드 1개로 직렬 처리합니다.
여러 클라이언트가 동시에 붙으면 순서대로 대기하며, `--max-queue`를 넘으면 `429`를 반환합니다.

### 평문 HTTP

기본은 평문 HTTP입니다. Tailscale은 WireGuard로 이미 암호화되므로 무방하지만,
일반 LAN에서는 토큰과 문서 내용이 평문으로 흐릅니다. `--tls-cert`/`--tls-key` 또는
`tailscale serve`를 쓰거나, 신뢰된 네트워크에서만 사용하세요.

### 잡별 타임아웃 미구현

서버에 한 파일이 한컴 대화상자 등으로 멈추면 그 잡이 워커를 붙잡습니다.
`HancomDialogWatcher`가 대부분을 자동 확인하지만, 잡별 타임아웃 + `kill_hwp()` 복구는
아직 구현되지 않았습니다(`docs/context.md` §10). 현재 우회책은 서버 재시작입니다.

## 4. macOS 로컬 한컴오피스는 자동화 불가

`HwpMac2014.app`(한컴오피스 for Mac)에는 **AppleScript 사전(`.sdef`)이 없고**
(바이너리에 `NSScript*` 심볼 없음), **CLI 변환 진입점도 없습니다**
(`application:openFile:`만 존재). x86_64 단독 빌드라 Rosetta로 실행됩니다.

남는 경로는 System Events GUI 스크립팅(`파일 > PDF로 저장하기`, 액션 ID
`HWPAID_FILE_SAVE_AS_PDF`)뿐인데, 손쉬운 사용 권한이 필요하고 변환 중 화면·키보드를
점유하며 잠금화면이나 원격 세션에서 실패합니다. 그래서 mac 앱은 로컬 한컴을 쓰지 않고
Windows 변환 서버에 연결하는 방식을 택했습니다.
