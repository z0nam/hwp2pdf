# Changelog

이 파일은 [Keep a Changelog](https://keepachangelog.com/) 1.1.0 형식을 느슨하게 따릅니다. 버전 번호는 빌드 스크립트(`scripts/build_windows.ps1`)가 `yyyy.MM.dd.N` 형태로 자동 부여합니다.

각 항목 옆 GitHub 링크는 그 버전의 release 페이지를 가리킵니다 — 다운로드 자산은 거기에 있습니다.

## [Unreleased]

(다음 release에 포함될 변경분이 누적되는 곳)

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

[Unreleased]: https://github.com/z0nam/hwp2pdf/compare/v2026.06.25.6...HEAD
[2026.06.25.6]: https://github.com/z0nam/hwp2pdf/releases/tag/v2026.06.25.6
[2026.06.25.5]: https://github.com/z0nam/hwp2pdf/releases/tag/v2026.06.25.5
