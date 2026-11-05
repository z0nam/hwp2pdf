---
name: hwp2pdf
description: 한컴오피스(한글)로 HWP/HWPX 파일이나 폴더를 PDF·DOCX·HWPX로 변환한다. "이 한글 파일 PDF로", "폴더 안 hwp 전부 PDF로", "hwp를 워드(docx)로" 같은 요청에 쓴다. 셸 명령 `hwp2pdf-cli`(Windows)를 부른다.
---

# hwp2pdf — HWP/HWPX를 PDF·DOCX·HWPX로 변환

이 PC의 한컴오피스(한글)를 COM으로 띄워 원본과 같은 모양으로 변환한다. 셸에서 PATH의 명령을 부른다.

- Windows: `hwp2pdf-cli` (jaix setup이 설치해 PATH에 넣는다)
- 맥·리눅스: 배포판도 `hwp2pdf-cli`(소스 설치면 `hwp2pdf`). 한컴이 없으므로 Windows 변환 서버(`--server`)가 있어야만 된다

명령이 없으면 추측해서 다른 방법을 찾지 말고 "jaix setup을 다시 실행하세요"라고 알린다.

**환경변수 `HWP2PDF_CLI`가 있으면 그 경로를 실행파일로 쓴다.** jaix가 관리하는 최신 사본이다. 관리자가 따로 깐 낡은 hwp2pdf가 PATH에서 먼저 잡힐 수 있어서다. 아래 예의 `hwp2pdf-cli` 자리에 그 경로를 따옴표로 넣는다.

```
& "$env:HWP2PDF_CLI" "<파일>" --pdf --local    # PowerShell
"$HWP2PDF_CLI" "<파일>" --pdf --local          # bash(Git Bash)
```

변수가 비어 있으면 PATH의 `hwp2pdf-cli`를 쓴다.

## 기본 사용

```
hwp2pdf-cli "<파일 또는 폴더>" --pdf --local
hwp2pdf-cli "<폴더>" --pdf --recursive --local            # 하위 폴더까지
hwp2pdf-cli "<파일>" --docx --local                        # 워드로
hwp2pdf-cli "<파일.hwp>" --hwpx --local                    # HWP → HWPX
hwp2pdf-cli "<폴더>" --pdf --docx --no-overwrite --local   # 여러 형식, 이미 있는 결과는 건너뜀
```

- Windows에서는 **`--local`을 붙인다.** 저장된 서버 설정이나 `HWP2PDF_SERVER_URL`이 있으면 `--local` 없이는 문서가 변환 서버로 올라간다. 서버는 사용자가 요청할 때만 쓴다(그때는 `--local`을 빼고 `--server`).

- 경로는 **항상 따옴표로 감싼다**(한글·공백 경로가 흔하다).
- 결과는 **원본과 같은 폴더에 같은 이름**으로 생긴다(`보고서.hwp` → `보고서.pdf`). 기본은 덮어쓰기라, 사용자가 기존 결과를 지키길 원하면 `--no-overwrite`.
- 형식 플래그를 안 주면 PDF.
- 파일이 많으면 오래 걸린다. 시작 전에 몇 개를 변환하는지 알리고, 끝나면 결과 경로를 알려 준다.

## 결과 읽기 (종료 코드)

| 코드 | 뜻 | 할 일 |
|---|---|---|
| 0 | 전부 성공 | 만든 파일 경로를 알린다 |
| 2 | 일부 실패·건너뜀 | 출력의 `Log:` 줄에 있는 `hwp2pdf_log.csv`를 읽어 실패한 파일과 이유를 알린다 |
| 1 | 시작 못 함 | 아래 표에서 `ERROR:` 메시지에 맞는 안내 |

## 시작 못 할 때

| 메시지 | 원인 | 할 일 |
|---|---|---|
| `HWP is already running` | 사용자가 한글을 열어 두었다 | **`--kill-hwp`를 쓰지 않는다**(저장 안 한 작업이 날아간다). 사용자에게 한글 문서를 저장하고 닫은 뒤 다시 하자고 말한다. `--allow-running-hwp`도 사용자가 동의할 때만 쓴다 |
| `한컴오피스 한글을 시작하지 못했습니다` | 이 PC에 한글이 없거나, 한글 첫 실행 창(동의 창 등)이 응답을 기다린다 | 한글이 있으면 사용자에게 한글을 한 번 직접 열어 뜨는 창을 닫게 한다. 한글이 없으면 아래 rhwp 근사 변환을 제안한다 |
| `Target not found` · `Select an HWP or HWPX file` | 경로·확장자 오류 | 경로를 다시 확인한다 |

### 한글이 없을 때: rhwp 근사 변환 (PDF만)

```
hwp2pdf-cli "<파일 또는 폴더>" --pdf --local --rhwp-fallback
```

한컴을 쓸 수 없으면 내장 rhwp로 PDF를 만든다. **쪽 나눔·간격이 원본과 다를 수 있는 근사 결과**이므로 사용자에게 반드시 그렇게 알린다. DOCX·HWPX는 만들 수 없다.

## 하지 않는 것

- `--kill-hwp`로 사용자의 한글을 닫지 않는다.
- 변환 서버(`serve`·`--server`·`find`)를 임의로 켜거나 바꾸지 않는다. 사용자가 요청할 때만 쓴다(설정은 저장소 README "변환 서버").
- HWP 문서 내용을 고치거나 서식을 채우는 일은 이 도구 몫이 아니다(hwp-agent).
