# Contributing

Thanks for helping improve hwp2pdf.

This project is a Windows Hancom Office COM automation tool, so reproducible environment details are especially useful.

## How To Report A Bug

Please include:

- Windows version
- Hancom Office Hangul / HWP version
- hwp2pdf version shown in the window title or `hwp2pdf-cli --version`
- Whether you used the GUI, CLI, zip, or installer
- Input file type: `.hwp` or `.hwpx`
- Selected output: PDF, DOCX, or both
- `hwp2pdf_log.csv` content, if it was created
- A sample file, if it can be shared safely

Do not attach confidential documents. If a file is needed to reproduce the issue, remove sensitive content first.

## Development Setup

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -e .
.\.venv\Scripts\python -m hwp2pdf
```

CLI:

```powershell
.\.venv\Scripts\hwp2pdf --help
```

## Build

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check_windows.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\build_windows.ps1
```

Installer builds require Inno Setup 6:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_installer.ps1
```

## Pull Requests

- Keep changes focused and small when possible.
- Do not commit generated files from `dist/`, `release/`, or `build/`.
- Avoid committing private test documents.
- Update `README.md` and `docs/context.md` when behavior changes.
- Test both GUI and CLI paths when conversion behavior changes.

## Contributors

- Namun Cho: creator and maintainer
- OpenAI Codex: development assistance
