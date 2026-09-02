# hwp2pdf project instructions

## Start here

Before changing this repository, read these files in order:

1. `docs/context.md`
2. `docs/ai-context.md`
3. The relevant sections of `README.md` and `CHANGELOG.md`

Treat the code and tests as authoritative when older context disagrees with the current tree.

## Working agreements

- Use `C:\Users\user\dev\hwp2pdf` as the canonical repository path. Do not introduce a dependency on the legacy `C:\Users\user\_projects\hwp2pdf` path.
- Preserve unrelated user changes. In particular, do not discard `.vscode/settings.json` changes unless explicitly requested.
- Keep normal-user documentation Korean-first and bilingual where practical.
- Preserve both installer and portable ZIP workflows.
- Show progress immediately for long-running GUI or conversion work.
- For Hancom COM hangs, check modal/security dialogs and stale `Hwp.exe` processes before concluding that conversion code is broken.
- DOCX support remains opt-in because Hancom export fidelity depends on the installed environment.
- Validate release and updater changes through the installed executable, update helper exit/log behavior, relaunch, and helper cleanup—not only through a successful build.
- Do not commit generated executables, ZIP files, runtime logs, or anything under `.ai-context/archive/`.

## Releases

Every artifact is produced by CI, not by hand. Publishing a GitHub release fires
`.github/workflows/release.yml`, which builds macOS arm64, macOS Intel and the
Windows executables plus installer, and attaches all seven files to that
release.

- **Do not build or upload release artifacts manually.** If a release is missing
  some, re-run the workflow (`gh workflow run release.yml -f tag=vX`) rather
  than building locally; a hand-built artifact may not match the tag.
- Building needs neither Hancom Office nor COM. Only `check_windows.ps1` and
  actual conversion do, which is why the Windows build can run on a plain
  GitHub runner.
- After publishing, confirm the release ends up with all seven assets. CI
  failures show up on the Actions tab, not in the release itself.

## Working across the two machines

The repository lives on a Windows box (Hancom Office, conversion server) and a
Mac. Both push to the same `origin`.

- `git fetch` before starting work. A large refactor started on a stale base
  cost a hand-merge of a 322-line diff once already.
- **Never copy the tree between machines with `tar`/`scp` when the branch is
  pushed** -- `git pull` instead. macOS `tar` writes AppleDouble `._*` companions
  for every entry, and `._.` in particular cannot be deleted with `Remove-Item`
  on Windows. If a copy is unavoidable, set `COPYFILE_DISABLE=1`.
- `gh` does not work from a non-interactive Windows session (SSH, scheduled
  task): its token is DPAPI-protected and only decrypts in the desktop session,
  so API calls return 401. Run `gh` from the desktop, or set `GH_TOKEN`.

## Historical context

The private, local-only raw histories copied from Codex, Claude, Gemini, and agy are indexed in `.ai-context/README.md`. Consult them only when an exact historical detail is needed and never paste or commit their full contents without explicit user approval.
