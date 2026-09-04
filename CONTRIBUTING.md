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
python scripts/smoke_remote.py http://host:17650 <token> sample.hwp
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

Linux:

```bash
./scripts/build_linux.sh
```

All platforms take the `yyyy.MM.dd.N` build number from `scripts/set_version.py`, which also
writes `src/hwp2pdf/version.py`. When building the same release across platforms, pin the version on the
subsequent builds (`./scripts/build_macos.sh 2026.08.28.3`,
`.\scripts\build_windows.ps1 -Version 2026.08.28.3`,
`./scripts/build_linux.sh 2026.08.28.3`).

## Syncing a working tree to a Windows box

When copying this tree to a Windows machine by hand (for example to test against
a real Hancom install), set `COPYFILE_DISABLE=1` if the source is macOS:

```bash
COPYFILE_DISABLE=1 tar --exclude=.git --exclude=.venv -czf tree.tgz .
```

Without it macOS `tar` emits AppleDouble `._*` companions for every entry, which
Windows cannot always delete (`._.` in particular defeats `Remove-Item`) and
which show up as a hundred untracked files. Prefer `git clone` / `git pull` over
copying whenever the branch is already pushed.

## Pull Requests

- Keep changes focused and small when possible.
- Do not commit generated files from `dist/`, `release/`, or `build/`.
- Avoid committing private test documents.
- Update `README.md` and `docs/context.md` when behavior changes.
- Run `python -m pytest tests -q` before opening a PR.
- Test both GUI and CLI paths when conversion behavior changes.
- Changes to `jobs.run_batch` or a backend affect Windows, macOS, and Linux alike; check them.

## Release Flow

`CHANGELOG.md` is the single source of truth for the release history; the GitHub
Release page mirrors it. Every artifact is built by CI, so no machine has to be
switched on and nothing is uploaded by hand.

1. Move the relevant entries out of `## [Unreleased]` into a new
   `## [yyyy.MM.dd.N] - yyyy-MM-dd` section, and add the matching link
   reference at the bottom of the file.
2. Stamp the version:

   ```bash
   python scripts/set_version.py yyyy.MM.dd.N
   ```

3. Commit `CHANGELOG.md` + `src/hwp2pdf/version.py` together, then push.
4. Publish the release. Creating it is what triggers the build:

   ```bash
   gh release create vYYYY.MM.DD.N --title vYYYY.MM.DD.N --notes-file notes.md
   ```

   `.github/workflows/release.yml` then builds macOS arm64, macOS Intel, Linux
   x86_64, the three Windows executables and the installer, and attaches all
   eight files.
   It takes a few minutes; watch it on the Actions tab.
5. If this release changed `API_VERSION` in `src/hwp2pdf/server/protocol.py`,
   update the conversion server machines in the same pass: the client compares
   that value and refuses a server reporting a different one. Otherwise the
   server can be updated whenever convenient — build versions do not have to
   match, only the protocol does. A mismatch fails closed with a clear message
   rather than misbehaving.
6. Confirm the release ended up with all eight assets. If any are missing, the
   build failed — fix it and re-run rather than uploading by hand:

   ```bash
   gh workflow run release.yml -f tag=vYYYY.MM.DD.N
   ```

Building requires neither Hancom Office nor COM, which is why the Windows
artifacts can be produced on a stock GitHub runner. Only `check_windows.ps1`
and actual conversion need a machine with Hangul installed.

## Contributors

- Namun Cho: creator and maintainer
- OpenAI Codex: development assistance
