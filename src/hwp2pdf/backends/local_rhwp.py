"""Local approximate PDF rendering with rhwp.

A fallback for when local Hancom Office or a remote conversion server cannot
start. rhwp is an independent HWP renderer, not Hancom's engine, so the output
is close but not equivalent -- measured against Hangul on a 390-page report,
rhwp paginated to 390 pages instead of 379, dropped header and table-of-contents
page numbers, and rendered some dashes as missing-glyph boxes. It is good enough
to read a document now; it is not a substitute for a deliverable PDF.

Nothing here is silent: the batch log and the CSV record which engine ran.
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path

from hwp2pdf.backends.base import BackendCapabilities, BackendUnavailable, JobResult
from hwp2pdf.i18n import translate

#: Marker written to the log and the CSV so an approximate PDF is identifiable
#: long after the fact.
ACTUAL_FORMAT = "rhwp"

RHWP_ENV_VAR = "HWP2PDF_RHWP"

def _vendored() -> tuple:
    """Source-tree, packaged-data and executable-adjacent locations."""
    roots = []
    module_path = Path(__file__).resolve()
    if len(module_path.parents) > 3:
        roots.append(module_path.parents[3])  # source checkout: <root>/src/hwp2pdf/...
    bundled_root = getattr(sys, "_MEIPASS", "")
    if bundled_root:
        roots.append(Path(bundled_root))
    roots.append(Path(sys.executable).resolve().parent)

    candidates = []
    for root in roots:
        vendor = root / "vendor" / "rhwp"
        candidates.extend((vendor / "rhwp", vendor / "rhwp.exe"))
    return tuple(dict.fromkeys(candidates))


#: Checked in order when no explicit path is configured.
def known_locations() -> tuple:
    return _vendored() + (
        Path("/usr/local/bin/rhwp"),
        Path("/opt/homebrew/bin/rhwp"),
        # The sibling project that first vendored rhwp.
        Path.home() / "dev" / "hwp-preview-slack-bot" / "vendor" / "rhwp" / "rhwp",
    )

#: rhwp renders; it cannot write Hancom's own formats.
SUPPORTED_FORMATS = ("PDF",)

DEFAULT_TIMEOUT = 900


def find_rhwp(explicit: str = "") -> Path | None:
    """Locate the rhwp binary: explicit path, then env var, then PATH, then known spots."""
    if explicit:
        candidate = Path(explicit).expanduser()
        return candidate if candidate.is_file() else None

    from_env = os.environ.get(RHWP_ENV_VAR)
    if from_env:
        candidate = Path(from_env).expanduser()
        if candidate.is_file():
            return candidate

    on_path = shutil.which("rhwp")
    if on_path:
        return Path(on_path)

    for candidate in known_locations():
        if candidate.is_file():
            return candidate
    return None


class RhwpBackend:
    """Renders PDFs locally with rhwp. PDF only, and approximate."""

    def __init__(self, rhwp_path: str = "", timeout: float = DEFAULT_TIMEOUT):
        self.rhwp_path = rhwp_path
        self.timeout = timeout
        self.binary = None
        self._sink = None

    @property
    def capabilities(self):
        return BackendCapabilities(
            name="local_rhwp",
            remote=False,
            # rhwp reads the source and writes the PDF directly; the safe-temp
            # copy exists to work around Hancom's path handling, not this.
            local_staging=False,
            manages_hwp_process=False,
            local_preflight=False,
        )

    def preflight(self, lang: str) -> None:
        self.binary = find_rhwp(self.rhwp_path)
        if self.binary is None:
            raise BackendUnavailable(translate(lang, "rhwp_not_found", var=RHWP_ENV_VAR))

    def open_session(self, sink, lang: str, options) -> None:
        self._sink = sink
        sink.put(("log", (translate(lang, "rhwp_engaged", path=self.binary), "warning")))

    def session_notes(self, lang: str) -> list:
        return [("log", (translate(lang, "rhwp_quality_note"), "warning"))]

    def blocked_reason(self, src_path, output_format, lang):
        if output_format not in SUPPORTED_FORMATS:
            return translate(lang, "rhwp_format_unsupported", format=output_format)
        return None

    def convert(self, job) -> JobResult:
        if job.output_format not in SUPPORTED_FORMATS:
            return JobResult(
                ok=False,
                message=translate(job.lang, "rhwp_format_unsupported", format=job.output_format),
            )

        job.save_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            result = subprocess.run(
                [str(self.binary), "export-pdf", str(job.open_path), "-o", str(job.save_path)],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=self.timeout,
            )
        except subprocess.TimeoutExpired:
            return JobResult(
                ok=False,
                message=translate(job.lang, "rhwp_timeout", seconds=int(self.timeout)),
            )
        except OSError as e:
            return JobResult(ok=False, message=translate(job.lang, "rhwp_failed", detail=e))

        if result.returncode != 0 or not job.save_path.exists():
            detail = (result.stderr or result.stdout or "").strip().splitlines()
            return JobResult(
                ok=False,
                message=translate(
                    job.lang, "rhwp_failed", detail=detail[-1] if detail else result.returncode
                ),
            )

        return JobResult(ok=True, actual_format=ACTUAL_FORMAT)

    def cancel(self) -> None:
        pass

    def close_session(self) -> None:
        self._sink = None
