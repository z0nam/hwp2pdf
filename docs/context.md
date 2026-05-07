# hwp2pdf Context

## 1. Project Overview

`hwp2pdf` is a Windows desktop converter for documents handled by Hancom Office Hangul.
It automates `HWPFrame.HwpObject` through COM and saves HWP/HWPX files as PDF or DOCX.

Primary scope:

- HWP -> PDF
- HWPX -> PDF
- HWP -> DOCX
- HWPX -> DOCX
- Folder batch conversion through a Tkinter GUI

DOCX output depends on Hancom Office's DOCX export fidelity.

## 2. Current Architecture

```text
Tkinter GUI
  -> background worker thread
  -> pywin32 COM bridge
  -> HWPFrame.HwpObject
  -> Open(input)
  -> SaveAs(output, selected format)
```

The application is packaged as a Python module under `src/hwp2pdf`.
The older `src/hwp_pdf_converter_app_safe.py` path remains as a compatibility entrypoint.

## 3. Current Feature Set

- Select a root folder from the GUI
- Convert `.hwp` and `.hwpx`
- Select PDF output, DOCX output, or both
- Korean UI/logs by default with an English switch
- Include or exclude subfolders
- Overwrite or skip existing output files
- Safe temp conversion through `C:\temp\hwp_convert_safe`
- Progress UI
- Estimated HWP/HWPX file count below the selected target path
- Colored on-screen logs for failures and warning states
- Stop request between files
- CSV conversion log: `hwp2pdf_log.csv`

## 4. Repository Layout

```text
assets/
  hwp_to_pdf_final.ico
docs/
  context.md
scripts/
  check_windows.ps1
  build_windows.ps1
src/
  hwp2pdf/
    __init__.py
    __main__.py
    app.py
  hwp_pdf_converter_app_safe.py
.gitignore
hwp2pdf.spec
README.md
requirements.txt
requirements-build.txt
pyproject.toml
```

Ignored/generated paths:

- `.venv/`
- `build/`
- `dist/`
- `release/`
- `*.exe`
- `*.zip`
- runtime logs

## 5. Runtime Requirements

- Windows
- Hancom Office Hangul installed
- Python 3.10+
- `pywin32`

COM registration may need:

```powershell
Hwp.exe -regserver
```

## 6. Source Run

```powershell
$env:PYTHONPATH = "src"
python -m hwp2pdf
```

Or install dependencies in a virtual environment first:

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -e .
.\.venv\Scripts\python -m hwp2pdf
```

## 7. Distribution Policy

Do not commit generated `.exe` binaries directly to the repository.

Preferred flow:

1. Keep source, dependency files, and `hwp2pdf.spec` in Git.
2. Build on Windows with `scripts/build_windows.ps1`.
3. Attach `release/hwp2pdf-windows.zip` to GitHub Releases or an equivalent distribution channel.

This keeps the repository reviewable and avoids binary churn.

## 8. Build

Check the Windows environment first:

```powershell
.\scripts\check_windows.ps1
```

Then build:

```powershell
.\scripts\build_windows.ps1
```

Expected outputs:

- `dist/hwp2pdf-YYYY.MM.DD.N.exe`
- `release/hwp2pdf-windows-YYYY.MM.DD.N.zip`

Versioned build numbers use the build date and the sequence number for that date, for example
`hwp2pdf-2026.04.25.1.exe`.

## 9. Known Issues

### COM Object Creation Fails

Likely causes:

- Hancom Office is not installed.
- COM registration is missing.
- Running Python bitness does not match the installed automation component.

Typical fix:

```powershell
Hwp.exe -regserver
```

### Output Not Created Or 0 KB

Likely causes:

- Hancom PDF engine issue
- DRM-protected document or distribution document with print/PDF export disabled
- Unsupported document structure
- DOCX export failure

If Hancom disables PDF export for the opened document, the app cannot bypass that restriction.
The failure log should report that PDF export is unavailable or blocked and suggest checking
Hancom document security or distribution-document settings.
For HWP files, the app reads the `FileHeader` flags before opening Hancom. If the distribution
document flag is set and the requested output is PDF, the app fails that file immediately and logs
the security restriction reason instead of opening a document that can hang on export.

### File Access Or Security Prompt

Safe temp mode helps with long paths, Google Drive, and network drives, but Hancom security
modules can still show prompts for some environments.
The app attempts to register Hancom's file path checker module with `RegisterModule`, but if
Hancom still asks whether to allow automation access, choose the permanent allow option.
Otherwise `Open()` can block while waiting for the prompt.
The prompt should not be bypassed with auto-click automation. Hancom's documented path is to
install/register the automation security approval module and then call `RegisterModule`.

### Concurrency

Parallel conversion is not recommended. Use one HWP COM process and convert files sequentially.

## 10. Future Work

- CLI mode in addition to the GUI
- Retry queue for failed files
- Per-file timeout and hung process recovery
- Watch-folder mode
- Windows conversion server/API for Mac/Linux clients
