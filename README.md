# hwp2pdf

Windows + Hancom Office COM 자동화를 사용하는 HWP/HWPX -> PDF/DOCX GUI 변환기입니다.

This is a Windows GUI converter that uses Hancom Office COM automation to convert HWP/HWPX files to PDF or DOCX.

## 주요 기능 / Features

- HWP/HWPX 파일을 PDF 또는 DOCX로 일괄 변환
- 출력 형식 PDF/DOCX 중 하나 또는 둘 다 선택
- 하위 폴더 포함/제외
- 기존 PDF 덮어쓰기 또는 건너뛰기
- 저장 전 한쪽 보기 강제 적용 옵션
- DOCX 저장 시 한컴 호환 문서 확인창 자동 확인
- 안전한 임시 폴더 변환 모드
- 변환 결과 CSV 로그 생성

- Batch convert HWP/HWPX files to PDF or DOCX
- Select PDF output, DOCX output, or both
- Include or exclude subfolders
- Overwrite or skip existing output files
- Option to force one-page view before export
- Auto-confirm Hancom compatibility warning dialogs during DOCX export
- Safe temporary local conversion mode
- CSV conversion log

## 요구 사항 / Requirements

- Windows
- 한컴오피스 한글 설치
- 일반 사용자는 Python 설치가 필요 없습니다. 배포용 zip에 포함된 `hwp2pdf.exe`를 실행하면 됩니다.

- Windows
- Hancom Office Hangul installed
- Normal users do not need to install Python. Run `hwp2pdf.exe` from the distributed zip package.

개발자 또는 소스 실행 사용자는 Python 3.10+ 및 `pywin32`가 필요합니다.

Developers or source users need Python 3.10+ and `pywin32`.

한컴 COM 등록이 깨진 경우 관리자 권한 셸에서 Hancom의 `Hwp.exe -regserver`를 한 번 실행하세요.

If COM registration fails, run Hancom's `Hwp.exe -regserver` once from an elevated shell.

## 테스트 환경 / Tested Environment

현재 배포 파일은 아래 환경에서 실행과 기본 GUI 동작을 확인했습니다.

- Windows 11
- Hancom Office Hangul / HWP 2022
- Python 3.12.10 build environment
- pywin32 311
- Hancom Office Hangul installed with `HWPFrame.HwpObject` COM automation available

The current release was tested for launch and basic GUI behavior in the environment above.

If the app does not work in another Windows or Hancom Office environment, please report the Windows
version, Hancom Office version, app version shown in the window title, and the contents of
`hwp2pdf_log.csv` if it was created.

다른 Windows 또는 한컴오피스 환경에서 동작하지 않으면 Windows 버전, 한컴오피스 버전, 앱 창 제목에 표시된 버전,
그리고 생성된 경우 `hwp2pdf_log.csv` 내용을 함께 전달해 주세요.

## 일반 사용자 실행 / Use The Windows App

일반 사용자는 Python을 설치하지 말고 미리 빌드된 Windows zip 파일을 사용하세요.

For normal users, use the prebuilt Windows zip instead of installing Python.

1. 릴리스 페이지 또는 배포자로부터 `hwp2pdf-windows-YYYY.MM.DD.N.zip`을 다운로드합니다.
2. zip 파일을 우클릭하고 **압축 풀기** 또는 **Extract All...**을 선택합니다.
3. 압축을 푼 `hwp2pdf.exe`를 실행합니다.
4. Windows SmartScreen이 표시되면, 출처를 신뢰할 수 있을 때만 **추가 정보 / More info** -> **실행 / Run anyway**을 선택합니다.
5. 한컴오피스에서 자동화 또는 파일 접근 허용 여부를 물으면 **항상 허용** 또는 영구 허용 옵션을 선택합니다.

1. Download `hwp2pdf-windows-YYYY.MM.DD.N.zip` from the release page or from whoever distributes the app.
2. Right-click the zip file and choose **Extract All...**.
3. Run the extracted `hwp2pdf.exe`.
4. If Windows SmartScreen appears, choose **More info** and then **Run anyway** only if you trust the source.
5. If Hancom Office asks whether to allow automation or file access, choose the permanent allow option.

별도 설치 프로그램이나 `_internal` 폴더는 필요 없습니다. 압축을 푼 `hwp2pdf.exe` 하나만 실행하면 됩니다.

The app does not need a separate installer or an `_internal` folder. Run the extracted `hwp2pdf.exe`.

## 변환 방법 / How To Convert

1. **Browse folder...**로 변환 대상 폴더를 직접 선택합니다. 파일 목록을 보면서 고르고 싶으면 **Pick file...**을 눌러 대상 폴더 안의 `.hwp`, `.hwpx`, `.pdf`, `.docx` 파일 중 아무 파일이나 선택합니다.
2. **Output**에서 **PDF**, **DOCX** 중 하나 또는 둘 다 선택합니다.
3. 하위 폴더 포함 여부, 기존 출력 파일 덮어쓰기 여부, **Force one-page view before export** 옵션을 선택합니다.
4. **Start conversion**을 누릅니다.

1. Click **Browse folder...** to select the target folder directly. To inspect files first, click **Pick file...** and select any `.hwp`, `.hwpx`, `.pdf`, or `.docx` file in the target folder.
2. Select **PDF**, **DOCX**, or both under **Output**.
3. Choose whether to include subfolders, overwrite existing output files, and force one-page view before export.
4. Click **Start conversion**.

PDF 또는 DOCX 파일은 원본 문서 옆에 생성됩니다. 선택한 루트 폴더에는 `hwp2pdf_log.csv` 변환 로그가 생성됩니다.

PDF or DOCX files are written beside the original documents. A conversion log named `hwp2pdf_log.csv` is written to the selected root folder.

## 소스에서 실행 / Run From Source

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -e .
.\.venv\Scripts\python -m hwp2pdf
```

패키지를 설치하지 않고 로컬 개발용으로 실행하려면:

For local development without installing the package:

```powershell
$env:PYTHONPATH = "src"
python -m hwp2pdf
```

## 빌드 / Build App

생성된 `.exe` 파일은 저장소에 직접 커밋하지 않습니다. Windows PC 또는 CI에서 빌드한 뒤 `release/` 폴더의 zip 파일을 GitHub Releases 등으로 배포하세요.

The repository should not commit generated `.exe` files. Build them locally or in CI and attach the generated zip from `release/` to a GitHub Release or other download channel.

Windows PC에서 빌드하기 전에 환경 체크를 실행합니다:

Before building on a Windows PC, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check_windows.ps1
```

빌드:

Build:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_windows.ps1
```

빌드 결과:

Build outputs:

- `dist/hwp2pdf-YYYY.MM.DD.N.exe`
- `release/hwp2pdf-windows-YYYY.MM.DD.N.zip`

버전 번호는 빌드 날짜와 당일 빌드 순번을 사용합니다. 예: `hwp2pdf-2026.04.25.1.exe`, `hwp2pdf-windows-2026.04.25.1.zip`.
앱 창 제목에도 같은 버전이 표시됩니다.

The version number uses the build date and the build sequence for that day. Example: `hwp2pdf-2026.04.25.1.exe`, `hwp2pdf-windows-2026.04.25.1.zip`.
The same version is shown in the app window title.

## 버전 히스토리 / Version History

### 2026.04.30.3

- DOCX 저장 동작 중에만 한컴의 호환 문서 확인창을 자동 확인하도록 `SetMessageBoxMode(0x10)`을 적용했습니다.
- Applied `SetMessageBoxMode(0x10)` only during DOCX save operations to auto-confirm Hancom compatibility warning dialogs.

### 2026.04.30.2

- PDF/DOCX 저장 전 한쪽 보기 `1x1`을 강제로 적용하는 옵션을 추가했습니다. 기본값은 켜짐입니다.
- Added an option to force one-page `1x1` view before PDF/DOCX export. It is enabled by default.

### 2026.04.30.1

- 폴더를 직접 선택하는 **Browse folder...**와 파일 목록을 보며 폴더를 지정하는 **Pick file...**을 분리했습니다.
- Split folder selection into **Browse folder...** for direct folder selection and **Pick file...** for selecting a folder via a visible file list.

### 2026.04.25.8

- 기본 Windows 파일 선택창을 사용하도록 변경했습니다. 파일 목록에서 `.hwp`, `.hwpx`, `.pdf`, `.docx`를 확인하고 아무 파일이나 선택하면 해당 폴더를 변환 대상으로 사용합니다.
- Switched back to the standard Windows file picker. Select any visible `.hwp`, `.hwpx`, `.pdf`, or `.docx` file and the app uses its folder as the conversion root.

### 2026.04.25.7

- 폴더 선택 창을 커스텀 UI로 변경했습니다. 폴더를 선택하면서 `.hwp`, `.hwpx`, `.pdf`, `.docx` 파일을 미리보기로 확인할 수 있습니다.
- Replaced the folder picker with a custom UI that previews `.hwp`, `.hwpx`, `.pdf`, and `.docx` files while selecting a folder.

### 2026.04.25.6

- DOCX 저장 안정성을 보강했습니다. `OOXML`, `DOCX`, `MSWORD` 저장 포맷을 순서대로 시도하고, 필요하면 한컴 `FileSaveAs_S` 액션 방식으로 재시도합니다.
- Improved DOCX export reliability by trying `OOXML`, `DOCX`, and `MSWORD`, then falling back to Hancom's `FileSaveAs_S` action.

### 2026.04.25.5

- 모든 변환이 성공하면 `hwp2pdf_log.csv`를 삭제하고, 성공 메시지를 간단히 표시합니다. 실패, 스킵, 중단이 있으면 로그를 남깁니다.
- Deletes `hwp2pdf_log.csv` when every conversion succeeds and shows a simple success message. Logs are kept when failures, skips, or stops occur.

### 2026.04.25.4

- 앱 창 제목에 빌드 버전을 표시합니다.
- Shows the build version in the app window title.

### 2026.04.25.3

- 출력 형식을 라디오 버튼에서 체크박스로 변경했습니다. PDF와 DOCX를 동시에 선택하면 두 형식을 모두 생성합니다.
- Changed output selection from radio buttons to checkboxes. Selecting both PDF and DOCX generates both outputs.

### 2026.04.25.2

- 단일 exe 배포 방식으로 변경했습니다. `_internal` 폴더 없이 `hwp2pdf.exe` 하나로 실행할 수 있습니다.
- 빌드 산출물에 `YYYY.MM.DD.N` 버전 번호를 붙이도록 변경했습니다.
- Switched to one-file exe distribution, so the app can run without an `_internal` folder.
- Added `YYYY.MM.DD.N` versioned build filenames.

### 2026.04.25.1

- HWP/HWPX 입력을 PDF 또는 DOCX 출력으로 변환하는 흐름으로 정리했습니다.
- DOCX 입력 변환 옵션을 제거하고, 출력 형식 선택 방식으로 변경했습니다.
- Clarified the app flow as HWP/HWPX input to PDF or DOCX output.
- Removed DOCX input conversion and replaced it with output format selection.

## 프로젝트 구조 / Project Layout

```text
assets/                 icon and static resources
docs/                   project notes
scripts/                build and maintenance scripts
src/hwp2pdf/            application package
src/hwp_pdf_converter_app_safe.py
                        compatibility entrypoint for older local usage
hwp2pdf.spec            PyInstaller recipe
pyproject.toml          Python package metadata
```

## 변환 참고 사항 / Conversion Notes

- 안전한 임시 변환 모드는 원본 파일을 `C:\temp\hwp_convert_safe`로 복사한 뒤 변환합니다.
- PDF 또는 DOCX 파일은 원본 파일 옆에 저장됩니다.
- `hwp2pdf_log.csv`는 선택한 루트 폴더에 저장됩니다.
- DOCX 출력 품질은 한컴오피스의 DOCX 내보내기 품질에 따라 달라집니다.
- 앱은 가능한 경우 한컴 파일 경로 보안 모듈을 등록합니다. 그래도 한컴에서 자동화 접근 허용 여부를 물으면 영구 허용을 선택하세요. 그렇지 않으면 COM 변환이 대기 상태로 멈출 수 있습니다.

- Safe temp mode copies each source file to `C:\temp\hwp_convert_safe` before conversion.
- PDF or DOCX files are written beside the original files.
- `hwp2pdf_log.csv` is written to the selected root folder.
- DOCX output fidelity depends on Hancom Office's DOCX export support.
- The app registers Hancom's file path security module when available. If Hancom still asks whether to allow automation access, allow it permanently; otherwise COM conversion can wait indefinitely.
