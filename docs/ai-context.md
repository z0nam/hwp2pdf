# hwp2pdf AI 작업 맥락

이 문서는 Codex, Claude, Gemini/agy가 동일한 프로젝트 맥락으로 새 작업을 시작하도록 만든 지속 가능한 요약이다. 세션 원문을 매번 읽는 대신 이 문서와 `docs/context.md`를 우선 사용한다.

## 현재 위치와 이관 상태

- 실제 저장소와 기준 경로는 `C:\Users\user\dev\hwp2pdf`이다.
- `C:\Users\user\_projects`는 `C:\Users\user\dev`를 가리키는 임시 디렉터리 정션이다. 별도의 옛 저장소 복사본이 아니다.
- 2026-08-27에 과거 hwp2pdf 원문 기록을 `.ai-context/archive/`로 복사했다.
- 전역 도구 원본과 `_projects` 정션은 유예 기간 동안 그대로 보존한다.
- 새 코드, 설정, 문서에는 `_projects` 절대경로 의존성을 추가하지 않는다.

## 도구별 진입점

- Codex: 저장소 루트의 `AGENTS.md`
- Claude Code: `CLAUDE.md`에서 `AGENTS.md`와 이 문서를 안내
- Gemini/agy: `GEMINI.md`에서 `AGENTS.md`와 이 문서를 안내
- 공통 제품·구조 설명: `docs/context.md`
- 사용자·설치·운영 안내: `README.md`
- 릴리스별 변경: `CHANGELOG.md`
- 민감할 수 있는 세션 원문: `.ai-context/archive/`(Git 제외)

## 유지해야 할 과거 결정

- 일반 사용자를 위한 Windows 앱이므로 문서는 한국어 우선으로 작성하고, 설치 프로그램과 포터블 ZIP 흐름을 모두 유지한다.
- HWP/HWPX의 PDF 변환이 주 기능이며 DOCX 출력은 Hancom 환경별 품질 차이 때문에 선택 기능으로 유지한다.
- Hancom COM 호출이 멈추면 코드 결함으로 단정하기 전에 보안/파일 접근 모달과 남은 `Hwp.exe` 프로세스를 확인한다.
- GUI에서 시작, 파일 검색, COM 초기화 등 시간이 걸리는 단계의 진행 상태를 즉시 표시한다.
- 배포 검증은 빌드 성공에서 끝내지 않는다. 설치된 CLI 버전, 업데이트 도우미 종료 로그, 재실행, 임시 도우미 정리까지 확인한다.
- 생성된 EXE/ZIP, 런타임 로그, AI 원문 아카이브는 커밋하지 않는다.

## 원문 아카이브 범위

- Codex: 이전 hwp2pdf 경로가 세션 작업 디렉터리였던 5개 JSONL 세션
- Claude: 이전 경로에 연결된 프로젝트 세션 1개(1,575 레코드)
- Gemini: hwp2pdf history/tmp 프로젝트 표식과 채팅 기록
- agy/Antigravity: hwp2pdf를 다룬 brain `90969c7a-1602-441b-9b51-a1de7afbc239` 전체

원문은 자동 프롬프트로 모두 주입하지 않는다. 크기가 크고 민감 정보가 섞일 수 있기 때문이다. 요약에서 빠진 정확한 오류·명령·결정이 필요할 때만 최소 범위로 조회한다.

## `_projects` 정션 제거 전 확인

다음 조건을 모두 만족한 뒤 별도 승인으로 제거한다.

1. Codex, Claude, agy를 각각 `C:\Users\user\dev\hwp2pdf`에서 새로 시작해 이 문서와 각 도구 진입점을 읽는지 확인한다.
2. 세 도구로 실제 hwp2pdf 작업을 수행해 `_projects` 경로 없이 파일 조회·수정·검증이 되는지 확인한다.
3. `.ai-context/archive/`의 파일 수와 바이트가 이관 당시 인벤토리와 일치하는지 확인한다.
4. 전역 설정과 자동 실행 항목에서 `_projects`를 참조하는 새 의존성이 없는지 다시 검색한다.
5. `_projects`는 hwp2pdf만의 별칭이 아니라 `dev` 전체의 별칭이므로, 다른 프로젝트의 이전 경로 의존성도 별도로 점검한다.
6. 정션 제거 직전에 대상이 여전히 `C:\Users\user\dev`를 가리키는 reparse point인지 확인하고 실제 `dev` 디렉터리에는 삭제 명령을 실행하지 않는다.

권장 유예 기준은 최소 2주와 각 도구별 실제 작업 1회 이상 중 더 늦게 충족되는 시점이다.
