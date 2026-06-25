# 한글 파일접근 보안모듈 자동등록

> 상태: **구현됨.** B안(앱 자가등록) + 자체 작성 stub DLL(MIT) 채택.
> 라이선스: 한컴 공개 예제 ZIP은 라이선스 미명시(컴파일 DLL 페이지의 "개인 비상업적 한정"
> 문구가 GitHub archive의 C++ 소스에 적용되는지 모호) → hwp2pdf MIT와 호환 보장이 어려워
> **export 시그니처(`IsAccessiblePath`, ABI)만 일치시킨 자체 stub**을 새로 작성. 한컴 코드/바이너리 일절 미사용.

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

### 라이선스 (실제 결정)
- pyhwpx가 번들한 바이너리는 **재배포하지 않는다.**
- 한컴 예제 ZIP에 C++ 소스가 포함돼 있으나 LICENSE 파일도 없고 소스 파일 헤더에도 고지가 없음 →
  hwp2pdf(MIT)에 묶기 곤란.
- 따라서 **export 시그니처만 일치하는 자체 stub DLL을 새로 작성**(`vendor/src/FilePathCheckerModule/`,
  MIT). 함수 시그니처는 ABI라 저작권 대상 아님. 한컴 코드 한 줄도 복사하지 않음.
- **x86 + x64 둘 다 빌드**(`vendor/x86/`, `vendor/x64/`), 실행 시 한글 비트수 감지해 맞는 DLL을 등록.

## 폐기
- `feat/auto-allow-file-access-dialog`(대화상자 자동클릭 watcher) 브랜치는 이 방식으로 **대체** —
  보안모듈 등록 시 대화상자가 *아예 안 뜨므로* 클릭 자체가 불필요. 해당 브랜치는 **삭제됨**.

---

## 구현 (현재)

### 1. 보안모듈 소스 — 자체 작성 stub
- `vendor/src/FilePathCheckerModule/FilePathCheckerModule.cpp` — `IsAccessiblePath()`가 무조건 `TRUE` 반환.
  hwp2pdf는 사용자가 UI/CLI에서 직접 선택한 파일만 변환하므로 모든 접근 허용이 의도된 동작.
- `.def`로 export 이름을 `IsAccessiblePath`로 고정 (x86 `__stdcall` decoration 회피, x64는 그대로).
- 라이선스: MIT(hwp2pdf와 동일). 한컴 코드 사용 안 함.

### 2. 빌드
- 의존성: Visual Studio 2022 Build Tools + C++ x86/x64 도구. `winget install Microsoft.VisualStudio.2022.BuildTools`
  실행 시 컴포넌트 `Microsoft.VisualStudio.Component.VC.Tools.x86.x64` 포함하도록 `--override` 지정.
- 빌드: `vendor\src\FilePathCheckerModule\build.bat` → x86, x64 모두 빌드 → `vendor\x86\`, `vendor\x64\`에 산출.
- 결과 DLL은 repo에 커밋(자체 빌드, 라이선스 자체 보유).

### 3. 자가등록 (`src/hwp2pdf/app.py`)
- `ensure_hwp_security_module_registered()` — 변환 시작 시 호출:
  1. `HKCU\Software\HNC\HwpAutomation\Modules\FilePathCheckerModule` 조회.
  2. 값 존재 + 파일 존재 + PE machine이 감지된 한글 bitness와 일치 → 그대로 사용(`"already"`).
  3. 비트수 mismatch이거나 값 없음 → 번들 DLL(`vendor/<arch>/FilePathCheckerModule.dll`)을
     `%LOCALAPPDATA%\hwp2pdf\security\`에 복사 + `winreg.SetValueEx` (`"registered"`).
  4. `detect_hwp_arch()`는 `HKLM\SOFTWARE\HNC\HwpRun\*` Path 값, PE 헤더 machine 필드, `Program Files (x86)` 경로,
     COM `GetHwpInfo("InstallPath")` 순으로 시도 → 실패 시 `x86` 기본값.
  5. PyInstaller `--onefile`은 `sys._MEIPASS/vendor/<arch>/`에 DLL이 들어 있다고 가정 — `hwp2pdf.spec` `datas`에 추가.
- 호출 후 기존 `RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")` → 보안 모듈: 켜짐.

### 4. 설치프로그램 (`installer/hwp2pdf.iss`)
- `[Files]` 두 DLL을 `{app}\vendor\x86\`, `{app}\vendor\x64\`에 배치.
- `[Registry]` HKCU `FilePathCheckerModule` = `{app}\vendor\x86\FilePathCheckerModule.dll`(기본값, 32bit HWP 다수).
- 64bit HWP에서는 `ensure_hwp_security_module_registered()`가 launch 시 PE machine으로 mismatch 감지해 x64로 override.

### 5. 테스트 (namun-ji)
- **현재 namun-ji는 *수동* 등록 상태.** 자가등록 검증 절차:
  1. `reg delete "HKCU\Software\HNC\HwpAutomation\Modules" /v FilePathCheckerModule /f`
  2. `Remove-Item -Recurse C:\Users\user\.hwpautomation`
  3. (개발 모드) `cd hwp2pdf; python -m hwp2pdf.cli <test.hwp> --pdf` 또는 설치본 exe 실행.
  4. 로그에 `보안 모듈 자가등록 완료` + `HWP 파일 접근 보안 모듈: 켜짐` 확인, 대화상자 없이 변환 완료.
- 접속: `ssh user@namun-ji`. 한글 출력 깨짐 → `powershell -EncodedCommand <UTF-16LE base64>`.
  COM은 세션1 필요 → `schtasks /create … /it /ru user` + `schtasks /run`. (참고: hwp-agent `docs/output-verification.md` Step 0/2)
- 통과하면 namun-ji의 수동 `C:\Users\user\.hwpautomation\` 정리 가능(자가등록이 `%LOCALAPPDATA%\hwp2pdf\security\`로 대체).
