# Contributing

Thanks for helping improve hwp2pdf.

This project drives Hancom Office Hangul through COM automation on Windows, and reaches that engine
over HTTP from macOS, so reproducible environment details are especially useful.

## How To Report A Bug

Please include:

- Your OS (and, for remote conversion, the conversion server's Windows version)
- Hancom Office Hangul / HWP version on the machine that converts
- hwp2pdf version shown in the window title or `hwp2pdf-cli --version`
- Whether you used the GUI, CLI, zip, or installer
- Input file type: `.hwp` or `.hwpx`
- Selected output: PDF, DOCX, or both
- `hwp2pdf_log.csv` content, if it was created
- For remote conversion: the `serve` startup banner and whether `Test connection` succeeds
- A sample file, if it can be shared safely

Do not attach confidential documents. If a file is needed to reproduce the issue, remove sensitive content first.

## Development Setup

Windows:

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -e .
.\.venv\Scripts\python -m hwp2pdf
.\.venv\Scripts\hwp2pdf --help
```

macOS — Apple's `/usr/bin/python3` ships Tk 8.5 and renders badly, so use a python.org build or
`brew install python-tk@3.13`:

```bash
python3 -m venv .venv
./.venv/bin/python -m pip install -e . -r requirements-dev.txt
./scripts/check_macos.sh
PYTHONPATH=src ./.venv/bin/python -m hwp2pdf
```

## Tests

```bash
python -m pip install -r requirements-dev.txt
python -m pytest tests -q
```

The suite runs without Hancom Office on either platform: `tests/fakes.py` supplies a `FakeBackend`,
and `tests/test_server_protocol.py` drives the real HTTP client against the real server with that
fake engine behind it. CI runs it on macOS and Windows (`.github/workflows/test.yml`).

To exercise a real conversion server:

```bash
python scripts/smoke_remote.py http://host:8765 <token> sample.hwp
```

## Build

Windows:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check_windows.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\build_windows.ps1
```

Installer builds require Inno Setup 6:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_installer.ps1
```

macOS:

```bash
./scripts/build_macos.sh
```

Both platforms take the `yyyy.MM.dd.N` build number from `scripts/set_version.py`, which also
writes `src/hwp2pdf/version.py`. When building the same release on both, pin the version on the
second build (`./scripts/build_macos.sh 2026.08.28.3`,
`.\scripts\build_windows.ps1 -Version 2026.08.28.3`).

## Pull Requests

- Keep changes focused and small when possible.
- Do not commit generated files from `dist/`, `release/`, or `build/`.
- Avoid committing private test documents.
- Update `README.md` and `docs/context.md` when behavior changes.
- Run `python -m pytest tests -q` before opening a PR.
- Test both GUI and CLI paths when conversion behavior changes.
- Changes to `jobs.run_batch` or a backend affect Windows and macOS alike; check both.

## Release Flow

`CHANGELOG.md` is the single source of truth for the release history; the
GitHub Release page mirrors it. When cutting a release:

1. Move the relevant entries out of `## [Unreleased]` into a new
   `## [yyyy.MM.dd.N] - yyyy-MM-dd` section, and add the matching link
   reference at the bottom of the file.
2. Run `scripts/build_windows.ps1` then `scripts/build_installer.ps1` — these
   stamp `src/hwp2pdf/version.py` to match the date / build number.
3. If the release ships a macOS build too, run `./scripts/build_macos.sh <same version>`
   on a Mac so both platforms carry the same build number.
4. Commit `CHANGELOG.md` + `src/hwp2pdf/version.py` together, then push.
5. `gh release create vYYYY.MM.DD.N --notes-file -` (or paste the same
   section body) and attach `release/hwp2pdf-setup-*.exe`,
   `release/hwp2pdf-windows-*.zip`, `dist/hwp2pdf-*.exe`,
   `dist/hwp2pdf-cli-*.exe`, and `release/hwp2pdf-macos-*.zip`.

## Contributors

- Namun Cho: creator and maintainer
- OpenAI Codex: development assistance
