# Local AI context archive

This directory keeps a recoverable, local-only copy of historical hwp2pdf work performed with Codex, Claude, Gemini, and agy before the canonical repository path became `C:\Users\user\dev\hwp2pdf`.

## Archive inventory

Copied on 2026-08-27 while all originals and the `_projects` compatibility junction were still present:

| Tool | Archived files | Archived bytes | Location |
|---|---:|---:|---|
| Codex | 5 | 10,247,129 | `archive/codex/` |
| Claude | 1 | 2,953,842 | `archive/claude/` |
| Gemini | 43 | 293,153 | `archive/gemini/` |
| agy / Antigravity | 651 | 7,652,393 | `archive/agy/` |

The `archive/` directory is excluded by `.gitignore` because raw transcripts can contain prompts, tool output, machine paths, and other private data. Do not force-add it to Git.

For routine work, use the durable summary in `docs/ai-context.md`. Open a raw archive only when an exact historical command, error, or decision is required.

Original global tool histories must remain untouched throughout the compatibility period. Removing the `_projects` junction is a separate, explicitly approved final step.
