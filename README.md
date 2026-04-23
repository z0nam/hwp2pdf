# hwp2pdf

Windows + Hancom Office COM automation based GUI converter for HWP/HWPX files.
DOCX conversion is available as an optional mode and uses Hancom Office's DOCX import/export path.

## Requirements

- Windows
- Hancom Office Hangul installed
- Python 3.10+
- `pywin32`

If COM registration fails, run Hancom's `Hwp.exe -regserver` once from an elevated shell.

## Run From Source

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -e .
.\.venv\Scripts\python -m hwp2pdf
```

For local development without installing the package:

```powershell
$env:PYTHONPATH = "src"
python -m hwp2pdf
```

## Build App

The repository should not commit generated `.exe` files. Build them locally or in CI and attach
the generated zip from `release/` to a GitHub Release or other download channel.

Before building on a Windows PC, run:

```powershell
.\scripts\check_windows.ps1
```

```powershell
.\scripts\build_windows.ps1
```

Build outputs:

- `dist/hwp2pdf/hwp2pdf.exe`
- `release/hwp2pdf-windows.zip`

## Project Layout

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

## Conversion Notes

- Safe temp mode copies each source file to `C:\temp\hwp_convert_safe` before conversion.
- PDF files are written beside the original files.
- `hwp2pdf_log.csv` is written to the selected root folder.
- DOCX conversion is disabled by default because output fidelity depends on Hancom Office's DOCX support.
