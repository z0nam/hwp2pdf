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
  보안모듈 등록 시 대화상자가 *아예 안 뜨므로* 클릭 자체가 불필요. 해당 브랜치는 닫는다.
