# 한글 파일접근 보안모듈 자동등록 (계획 메모)

> 상태: **계획만 기록.** 코드 미반영. 예제 소스 직접 빌드로 진행 예정.

## 문제
hwp2pdf는 변환 시 `RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")`을 호출하지만,
해당 보안모듈이 **설치/등록돼 있지 않으면 "사용 불가"** → 한글이 *"… 파일에 접근하려는 시도가
있습니다. 접근을 허용하시겠습니까?"* 대화상자를 띄움 → 무인/배치 변환이 그 창에서 **무한 대기(행)**.
(README의 자동화 접근 허용 경고와 동일.)

## namun-ji 현황 (2026-06-23, **수동 조치**)
- 보안모듈 DLL(x86, 한글 32bit와 비트 일치)을 `C:\Users\user\.hwpautomation\`에 배치하고
  레지스트리 등록: `HKCU\Software\HNC\HwpAutomation\Modules` 값 `FilePathCheckerModule` = DLL 경로.
- 결과: 무인 변환이 **대화상자 없이 45초에 PDF 완료**(등록 전엔 120초 행+PDF 0).
- **단, 이건 그 머신 한정 수동 설정 — repo/배포본엔 미반영.**

## 계획 (배포본 자동화)
**접근 B = 앱 첫 실행 시 자가등록** (권장; zip·설치본 공통, 레지스트리 지워져도 자가복구.
pyhwpx `Hwp()`가 하는 방식):
- exe 옆에 보안모듈 DLL 번들 → 시작 시 레지스트리 등록 여부 확인 → 없으면 등록 후 `RegisterModule`.
- 등록 키: `HKCU\Software\HNC\HwpAutomation\Modules\FilePathCheckerModule = <dll 경로>`
  (HKCU라 관리자 권한 불요).
- 대안 A(설치프로그램 `[Files]`+`[Registry]` 등록)는 설치본에서만 동작 → B가 우선.

### 라이선스 (중요 — 반드시 준수)
- pyhwpx가 번들한 바이너리를 **재배포하지 않는다.**
- **한컴이 공개한 보안모듈 예제(FilePathCheckerModuleExample) 소스로 직접 빌드** → 출처·라이선스 명확.
- **x86 + x64 둘 다 빌드**, 실행 시 한글 비트수 감지해 맞는 DLL을 등록(대부분 32bit이나 64bit 한글 대비).

## 폐기
- `feat/auto-allow-file-access-dialog`(대화상자 자동클릭 watcher) 브랜치는 이 방식으로 **대체** —
  보안모듈 등록 시 대화상자가 *아예 안 뜨므로* 클릭 자체가 불필요. 해당 브랜치는 **삭제됨**.

---

## 본작업 지시사항 (hwp2pdf 세션에서 이어서)

목표: 위 **B안(앱 자가등록) + 예제소스 직접 빌드(라이선스 클린)** 구현.

### 1. 보안모듈 소스 확보 — 라이선스 먼저
- 한컴이 공개한 **FilePathCheckerModule 예제**(아래아한글 자동화 보안모듈 샘플) **C++ COM 소스**를 확보.
  출처(한컴 개발자센터 / 한글 자동화 SDK 샘플)와 **라이선스를 먼저 확인.** 불명확하면 거기서 멈추고 확인.
- **pyhwpx가 번들한 바이너리는 재배포 금지.** 반드시 소스로 직접 빌드. (pyhwpx `core.py`의
  `register_regedit`/`check_registry_key`는 *로직 참고만*, 코드·바이너리 복사 금지.)

### 2. 빌드 (x86 + x64 둘 다)
- MSVC로 32bit·64bit 빌드 → repo `vendor/`에 (자체 빌드라 출처 명확).

### 3. 자가등록 구현 (`src/hwp2pdf/app.py`)
- 변환 시작(또는 앱 시동) 시:
  1. `HKCU\Software\HNC\HwpAutomation\Modules` 값 `FilePathCheckerModule` 존재·경로 유효 확인.
  2. 없으면 → **한글 비트수 감지**(설치경로 `Program Files (x86)` 여부 / `Hwp.exe` PE 헤더)로 맞는
     DLL을 안정 경로(예: `%LOCALAPPDATA%\hwp2pdf\`)에 복사 → `winreg.SetValueEx`로 등록.
  3. 그 후 기존 `RegisterModule("FilePathCheckDLL","FilePathCheckerModule")` → 로그 "보안 모듈: 켜짐".
- HKCU라 관리자 권한 불요. zip·설치본 공통으로 동작(자가복구).

### 4. 설치프로그램 보강 (`installer/hwp2pdf.iss`)
- `[Files]`(DLL) + `[Registry]`(HKCU 값)로 설치 즉시 동작 — B의 belt-and-suspenders.

### 5. 테스트 (namun-ji = 유일한 한글 COM 노드)
- **현재 namun-ji는 *수동* 등록 상태.** 자가등록을 진짜 테스트하려면 먼저 해제:
  - `reg delete "HKCU\Software\HNC\HwpAutomation\Modules" /v FilePathCheckerModule /f`
  - `Remove-Item -Recurse C:\Users\user\.hwpautomation`
- 그 뒤 hwp2pdf 실행 → 자가등록 → **대화상자 없이 변환되는지** 확인.
- 접속: `ssh user@namun-ji` (Windows, 기본 PowerShell). **한글 출력 깨짐** → `powershell
  -EncodedCommand <UTF-16LE base64>` 패턴 사용. **COM은 세션1 필요** → SSH(세션0)에선 행하므로
  `schtasks /create … /it /ru user` + `schtasks /run`으로 세션1 실행. (검증 결과는
  hwp-agent `docs/output-verification.md` Step 0/2 참고.)

### 현황
- auto-allow watcher 브랜치 삭제됨. namun-ji 수동등록은 **자가등록 검증 통과 후 정리**.
