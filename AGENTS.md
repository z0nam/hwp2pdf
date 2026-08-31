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

## Historical context

The private, local-only raw histories copied from Codex, Claude, Gemini, and agy are indexed in `.ai-context/README.md`. Consult them only when an exact historical detail is needed and never paste or commit their full contents without explicit user approval.
