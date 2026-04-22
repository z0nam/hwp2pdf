# HWP → PDF CLI Converter (Windows COM 기반)

## 1. 프로젝트 개요

본 프로젝트는 한글(HWP/HWPX) 파일을 PDF로 변환하는 CLI 기반 도구를 구현하는 것을 목표로 한다.

- HWP는 폐쇄 포맷 → 순수 CLI 변환 사실상 불가
- Windows + 한컴오피스 COM 자동화 방식 사용
- PowerShell 기반 1차 구현 완료
- CLI처럼 사용 가능한 자동화 래퍼 구조

---

## 2. 기술 아키텍처

User (CLI)    ↓ PowerShell Script    ↓ COM 객체 (HWPFrame.HwpObject)    ↓ 한글(HWP) 엔진 실행    ↓ Open → SaveAs(PDF)    ↓ PDF 파일 생성

---

## 3. 현재 구현 범위

### ✔ 완료 기능

- 단일 파일 변환
- 폴더 일괄 변환
- 하위 폴더 재귀 처리
- 출력 폴더 분리
- 폴더 구조 유지 옵션
- CSV 로그 기록
- CLI wrapper (.bat)

---

## 4. 디렉토리 구조

project-root/  ├── convert-hwp.ps1          # 단일 파일 변환  ├── convert-folder.ps1       # 폴더 일괄 변환  ├── hwp2pdf.bat              # CLI 실행용 래퍼  ├── logs/                    # 변환 로그  └── context.md               # 프로젝트 컨텍스트

---

## 5. 핵심 로직

### 5.1 COM 객체 생성

powershell $hwp = New-Object -ComObject HWPFrame.HwpObject 

---

### 5.2 문서 열기

powershell $hwp.Open($inputPath) 

---

### 5.3 PDF 저장

powershell $hwp.SaveAs($outputPath, "PDF", "") 

---

### 5.4 종료

powershell $hwp.Quit() 

---

## 6. 실행 방법

### 단일 파일

powershell -ExecutionPolicy Bypass -File convert-hwp.ps1 -InputPath "C:\docs\sample.hwp"

---

### 폴더 일괄

powershell -ExecutionPolicy Bypass -File convert-folder.ps1 -SourceDir "C:\docs" -Recurse

---

### 배치 실행

hwp2pdf.bat sample.hwp

---

## 7. 필수 환경

### OS
- Windows 필수

### Software
- 한컴오피스 한글 설치

### 초기 설정 (중요)

Hwp.exe -regserver

---

## 8. 주요 옵션 설명

| 옵션 | 설명 |
|------|------|
| -Recurse | 하위 폴더 포함 |
| -TargetDir | 출력 경로 지정 |
| -KeepStructure | 폴더 구조 유지 |

---

## 9. 로그 시스템

- CSV 형식
- 필드:
  - time
  - source
  - output
  - status
  - message

---

## 10. 알려진 문제

### 10.1 COM 객체 생성 실패
- 원인: 등록 문제
- 해결: Hwp.exe -regserver

---

### 10.2 PDF 생성 실패
- 증상: 0KB 파일
- 원인:
  - 한글 내부 PDF 엔진 문제
  - DRM 문서
  - 특정 문서 구조

---

### 10.3 경로 문제
- 상대경로 사용 시 오류 가능
- 한글 경로/파일명 이슈

✔ 권장:
- 절대경로 사용
- 영문 파일명

---

### 10.4 보안 팝업
- 파일 접근 보안 모듈 영향 가능

---

## 11. 성능 고려사항

- 파일 간 딜레이 필요 (COM 안정성)
- 동시 실행 비추천
- 단일 프로세스 권장

---

## 12. 향후 개발 계획

### 기능 확장

- [ ] Python CLI 통합
- [ ] watch folder 자동 변환
- [ ] 실패 재시도 큐
- [ ] CLI 옵션 고도화

---

### 안정성

- [ ] COM 재시도 로직
- [ ] timeout 처리
- [ ] hung process 감지

---

### 운영

- [ ] REST API 서버화
- [ ] 변환 서버 분리 (Windows)
- [ ] 로그 시각화

---

### 확장 구조

Mac / Linux    ↓ API 호출    ↓ Windows 변환 서버    ↓ HWP COM    ↓ PDF 반환

---

## 13. 핵심 요약

- HWP 변환은 한글 엔진 의존
- COM 자동화가 가장 안정적인 방식
- CLI는 래핑 형태로 구현
- 실무에서는 Windows 변환 서버 분리 권장

---