# 알려진 한계 / 후속 과제

이번 변경에서 우회하거나 부분만 잡은 항목들. 일반 사용 흐름에서 발생률이 낮아 보류 중.

## 1. 한컴 "알 수 없는 형식의 파일입니다." dialog 자동 처리 — 추가 검증 필요

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
