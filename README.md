# hwp2pdf

Windows + Hancom Office COM 자동화를 사용하는 HWP/HWPX -> PDF/DOCX GUI/CLI 변환기입니다.

This is a Windows GUI/CLI converter that uses Hancom Office COM automation to convert HWP/HWPX files to PDF or DOCX.

## 주요 기능 / Features

- HWP/HWPX 파일을 PDF 또는 DOCX로 단일 파일 또는 폴더 일괄 변환
- GUI와 명령줄 CLI 지원
- 출력 형식 PDF/DOCX 중 하나 또는 둘 다 선택
- 기본 한국어 UI/로그와 영어 전환
- GitHub Releases 기반 하루 1회 자동 업데이트 확인과 새 버전 업그레이드 버튼
- 하위 폴더 포함/제외
- 기존 PDF 덮어쓰기 또는 건너뛰기
- 저장 전 한쪽 보기 강제 적용 옵션
- DOCX 저장 시 한컴 호환 문서 확인창 자동 확인
- 안전한 임시 폴더 변환 모드
- 변환 결과 CSV 로그 생성

- Convert one HWP/HWPX file or batch convert a folder to PDF or DOCX
- GUI and command-line CLI support
- Select PDF output, DOCX output, or both
- Korean UI/logs by default with an English switch
- Automatic daily update check through GitHub Releases with an upgrade button when a newer version exists
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

일반 사용자는 Python을 설치하지 말고 미리 빌드된 Windows 설치 파일 또는 zip 파일을 사용하세요.

For normal users, use the prebuilt Windows installer or zip instead of installing Python.

권장 방식:

1. 릴리스 페이지 또는 배포자로부터 `hwp2pdf-setup-YYYY.MM.DD.N.exe`를 다운로드합니다.
2. 설치 파일을 실행하고 안내에 따라 설치합니다.
3. Windows SmartScreen이 표시되면, 출처를 신뢰할 수 있을 때만 **추가 정보 / More info** -> **실행 / Run anyway**을 선택합니다.
4. 한컴오피스에서 자동화 또는 파일 접근 허용 여부를 물으면 **항상 허용** 또는 영구 허용 옵션을 선택합니다.

Recommended:

1. Download `hwp2pdf-setup-YYYY.MM.DD.N.exe` from the release page or from whoever distributes the app.
2. Run the installer and follow the setup wizard.
3. If Windows SmartScreen appears, choose **More info** and then **Run anyway** only if you trust the source.
4. If Hancom Office asks whether to allow automation or file access, choose the permanent allow option.

대체 방식:

1. `hwp2pdf-windows-YYYY.MM.DD.N.zip`을 다운로드합니다.
2. zip 파일을 우클릭하고 **압축 풀기** 또는 **Extract All...**을 선택합니다.
3. 압축을 푼 `hwp2pdf.exe`를 실행합니다.

Alternative:

1. Download `hwp2pdf-windows-YYYY.MM.DD.N.zip`.
2. Right-click the zip file and choose **Extract All...**.
3. Run the extracted `hwp2pdf.exe`.

### 보안 경고 / Security Warning

현재 배포 파일은 코드 서명 인증서로 서명되지 않았습니다. 따라서 Windows SmartScreen 또는 브라우저 보안 경고가 표시될 수 있습니다. 이 경고는 설치 파일을 사용하더라도 코드 서명 전에는 완전히 사라지지 않습니다. 출처를 신뢰할 수 있을 때만 실행하세요.

The current release files are not signed with a code signing certificate. Windows SmartScreen or browser warnings may appear. Using an installer improves the installation experience, but it does not fully remove these warnings without code signing. Run the app only when you trust the source.

## 변환 방법 / How To Convert

1. 폴더 전체를 변환하려면 **Browse folder...**로 변환 대상 폴더를 선택합니다. 파일 하나만 변환하려면 **Pick file...**로 `.hwp` 또는 `.hwpx` 파일을 선택합니다.
2. **Output**에서 **PDF**, **DOCX** 중 하나 또는 둘 다 선택합니다.
3. 폴더를 선택한 경우 하위 폴더 포함 여부를 선택합니다. 파일을 선택한 경우 **Include subfolders**는 비활성화되고 선택한 파일만 변환합니다.
4. 기존 출력 파일 덮어쓰기 여부, **Force one-page view before export** 옵션을 선택합니다.
5. **Start conversion**을 누릅니다.

1. To convert a whole folder, click **Browse folder...** and select the target folder. To convert one file, click **Pick file...** and select an `.hwp` or `.hwpx` file.
2. Select **PDF**, **DOCX**, or both under **Output**.
3. If a folder is selected, choose whether to include subfolders. If a file is selected, **Include subfolders** is disabled and only that file is converted.
4. Choose whether to overwrite existing output files and force one-page view before export.
5. Click **Start conversion**.

PDF 또는 DOCX 파일은 원본 문서 옆에 생성됩니다. 선택한 루트 폴더에는 `hwp2pdf_log.csv` 변환 로그가 생성됩니다.

PDF or DOCX files are written beside the original documents. A conversion log named `hwp2pdf_log.csv` is written to the selected root folder.

## CLI 사용 / CLI Usage

소스에서 설치한 경우 `hwp2pdf` 명령으로 같은 변환 기능을 실행할 수 있습니다. 출력 형식을 지정하지 않으면 PDF만 생성합니다.

After installing from source, the same conversion engine is available through the `hwp2pdf` command. If no output format is selected, it exports PDF only.

```powershell
hwp2pdf "C:\docs\sample.hwp"
hwp2pdf "C:\docs\sample.hwpx" --pdf --docx
hwp2pdf "C:\docs\folder" --pdf --recursive
hwp2pdf "C:\docs\folder" --docx --no-overwrite
```

Windows 배포 zip 또는 설치 파일에는 콘솔용 `hwp2pdf-cli.exe`도 포함됩니다.

The Windows zip and installer also include the console-friendly `hwp2pdf-cli.exe`.

```powershell
hwp2pdf-cli.exe "C:\docs\sample.hwp" --pdf
hwp2pdf-cli.exe "C:\docs\folder" --pdf --docx --recursive
```

주요 옵션:

Main options:

- `--pdf`: PDF 생성
- `--docx`: DOCX 생성
- `-r`, `--recursive`: 폴더 변환 시 하위 폴더 포함
- `--no-overwrite`: 기존 출력 파일이 있으면 건너뛰기
- `--no-safe-temp`: 안전 임시 폴더 복사 모드 끄기
- `--no-force-one-page`: PDF 저장 전 한쪽 보기/모아찍기 해제 강제 적용 끄기
- `--kill-hwp`: 실행 중인 아래한글 프로세스를 강제 종료하고 진행
- `--allow-running-hwp`: 아래한글이 이미 실행 중이어도 그대로 진행

## 소스에서 실행 / Run From Source

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -e .
.\.venv\Scripts\python -m hwp2pdf
.\.venv\Scripts\hwp2pdf --help
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
- `dist/hwp2pdf-cli-YYYY.MM.DD.N.exe`
- `release/hwp2pdf-windows-YYYY.MM.DD.N.zip`

설치 파일을 만들려면 Inno Setup 6을 설치한 뒤 아래 명령을 실행합니다:

To build the installer, install Inno Setup 6 and run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_installer.ps1
```

설치 파일 결과:

Installer output:

- `release/hwp2pdf-setup-YYYY.MM.DD.N.exe`

버전 번호는 빌드 날짜와 당일 빌드 순번을 사용합니다. 예: `hwp2pdf-2026.04.25.1.exe`, `hwp2pdf-windows-2026.04.25.1.zip`.
앱 창 제목에도 같은 버전이 표시됩니다.

The version number uses the build date and the build sequence for that day. Example: `hwp2pdf-2026.04.25.1.exe`, `hwp2pdf-windows-2026.04.25.1.zip`.
The same version is shown in the app window title.

## 버전 히스토리 / Version History

### 2026.05.09.4

- `hwp2pdf` 명령줄 CLI와 배포용 `hwp2pdf-cli.exe`를 추가했습니다. 파일/폴더 대상, PDF/DOCX 선택, 하위 폴더 포함, 덮어쓰기 제어, 안전 임시 폴더 모드, 한쪽 보기 강제 적용 옵션을 지원합니다.
- `python -m hwp2pdf`는 인자가 없으면 GUI를 실행하고, 인자가 있으면 CLI로 동작합니다.
- Added the `hwp2pdf` command-line CLI and distributable `hwp2pdf-cli.exe` with file/folder targets, PDF/DOCX selection, recursive folder conversion, overwrite control, safe temp mode, and force one-page options.
- `python -m hwp2pdf` opens the GUI with no arguments and runs the CLI when arguments are provided.

### 2026.05.09.3

- 수동 **업데이트 확인** 버튼을 제거하고, 앱 시작 시 하루 1회만 GitHub Releases를 조용히 확인하도록 변경했습니다.
- 새 버전이 있으면 현재 버전과 최신 버전을 함께 표시하고 **업그레이드** 버튼을 보여줍니다. 최신 버전이면 상태 문구로만 안내합니다.
- Removed the manual **Check updates** button and changed update checks to run quietly once per day on app startup.
- When a newer version exists, the app shows the current/latest versions and an **Upgrade** button. If the app is current, it shows a mild status message only.

### 2026.05.09.2

- Inno Setup 기반 설치 파일 스크립트와 빌드 스크립트를 추가했습니다.
- Windows SmartScreen 등 보안 경고가 코드 서명 전에는 표시될 수 있음을 README에 명시했습니다.
- 앱에서 GitHub Releases 최신 버전을 확인하고 다운로드 페이지를 열 수 있는 **업데이트 확인** 버튼을 추가했습니다. 아직 릴리스가 없으면 별도 안내를 표시합니다.
- Added Inno Setup installer script and installer build script.
- Documented that Windows SmartScreen and similar warnings may appear before code signing.
- Added a **Check updates** button that checks the latest GitHub Release and opens the download page. If no release exists yet, the app shows a clear notice.

### 2026.05.07.2

- 한쪽 보기 강제 적용 옵션을 보강했습니다. `ViewZoom` 실행 시 `ZoomCustomDlg=1`, `ZoomCntX=1`, `ZoomCntY=1`, `ZoomType=1`을 함께 지정해 2쪽 보기 상태를 더 확실히 1쪽 보기로 되돌립니다.
- Strengthened the force one-page view option by setting `ZoomCustomDlg=1`, `ZoomCntX=1`, `ZoomCntY=1`, and `ZoomType=1` together when executing `ViewZoom`.
- PDF 저장 전 `PrintToPDFEx`의 `PrintMethod=0`을 적용해 문서에 저장된 모아찍기/2쪽씩 인쇄 옵션을 자동 인쇄로 초기화합니다.
- Before PDF export, applies `PrintMethod=0` through `PrintToPDFEx` to reset saved N-up / two-pages-per-sheet print settings to automatic print.
- 한쪽 보기 강제 적용이 실제 파일 저장 전에 실행되면 로그에 표시합니다.
- Logs when the one-page view / N-up reset is actually applied before exporting a file.
- 한쪽 보기 강제 적용이 켜진 PDF 변환은 `SaveAs(PDF)` 대신 `PrintToPDFEx`에 저장 경로를 지정해 직접 PDF를 생성합니다.
- When force one-page view is enabled for PDF, the app now creates the PDF directly through `PrintToPDFEx` with the target filename instead of `SaveAs(PDF)`.
- 모아찍기 인쇄 방식이 감지된 파일에만 기존 인쇄 방식과 자동 인쇄 강제 적용 내용을 로그에 표시합니다.
- Logs the original print method and automatic-print reset only when an N-up print method is detected.

### 2026.05.07.1

- 실패 로그를 빨간색으로, 건너뜀/중단 로그를 경고색으로 표시합니다.
- **Yes**로 실행 중인 HWP를 종료하기 전에 프로세스를 다시 확인하여 사용자가 이미 닫은 경우 그대로 변환을 진행합니다.
- 폴더 또는 파일 선택 시 입력창 아래에 변환 예상 HWP/HWPX 파일 수를 표시합니다.
- 한컴 자동화 보안 허용 팝업은 공식 보안 승인 모듈 등록이 필요하므로, 앱은 기존 `RegisterModule` 호출을 유지하고 자동 클릭 방식의 우회는 하지 않습니다.
- Failure logs are shown in red, while skipped/stopped logs use a warning color.
- Before force-closing HWP after **Yes**, the app checks the process list again and continues if the user already closed it manually.
- Shows the estimated number of HWP/HWPX files below the target path when a folder or file is selected.
- Hancom automation security prompts require the official security approval module registration, so the app keeps using `RegisterModule` and does not bypass prompts by auto-clicking dialogs.

### 2026.05.06.4

- 앱 기본 언어를 한국어로 변경하고, 우측 상단에서 **한국어 / English**를 전환할 수 있도록 했습니다. UI, 팝업, 로그, CSV 메시지가 선택 언어를 따릅니다.
- Changed the default app language to Korean and added a **한국어 / English** switch in the top-right. UI, dialogs, logs, and CSV messages follow the selected language.

### 2026.05.06.3

- HWP `FileHeader`의 배포용 문서 플래그를 읽어 PDF 변환이 제한될 가능성이 있는 파일은 한컴을 열기 전에 실패 처리하고 사유를 로그에 남깁니다.
- Reads the HWP `FileHeader` distribution-document flag and fails PDF conversion before opening Hancom when PDF export is likely restricted, with the reason written to the log.

### 2026.05.06.2

- PDF 저장이 실패하고 출력 파일이 생성되지 않는 경우, 한컴 문서 보안 또는 배포용 문서의 인쇄/PDF 제한 가능성을 실패 사유에 표시합니다.
- When PDF export fails without creating an output file, the failure reason now mentions possible Hancom document security or distribution-document print/PDF restrictions.

### 2026.05.06.1

- **Pick file...**을 단일 파일 변환 모드로 변경했습니다. 파일 선택 시 **Include subfolders**를 비활성화하고 선택한 파일 하나만 변환/로그 기록합니다.
- Changed **Pick file...** to single-file conversion mode. When a file is selected, **Include subfolders** is disabled and only that file is converted and logged.

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
