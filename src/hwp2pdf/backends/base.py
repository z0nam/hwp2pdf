"""Conversion backend contract.

``jobs.run_batch`` owns everything that depends on the destination filesystem --
file discovery, skip/overwrite rules, the CSV audit log, progress events and the
stop flag. A backend owns only "open this file and save it as that format",
whether that happens through local COM automation or over HTTP.
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol


class BackendUnavailable(Exception):
    """Raised when a backend cannot start a conversion session.

    The message is already localized and is shown to the user as-is.
    ``fallback_allowed`` is false for configuration errors such as a bad token
    or an incompatible server protocol: silently using another engine would
    hide a problem the user needs to fix.
    """

    def __init__(self, message: str, *, fallback_allowed: bool = True):
        super().__init__(message)
        self.fallback_allowed = fallback_allowed


@dataclass(frozen=True)
class BackendCapabilities:
    name: str
    #: Conversion happens on another machine.
    remote: bool = False
    #: ``run_batch`` may stage inputs through the local safe temp folder.
    local_staging: bool = True
    #: The HWP-already-running prompt and the kill button are meaningful.
    manages_hwp_process: bool = True
    #: ``blocked_reason`` can answer without leaving this machine.
    local_preflight: bool = True


@dataclass(frozen=True)
class SessionOptions:
    lang: str
    output_formats: tuple
    force_one_page: bool
    safe_temp: bool
    total_files: int


@dataclass(frozen=True)
class JobSpec:
    #: 1-based index of the source file within the batch (used for staging names).
    index: int
    src_path: Path
    open_path: Path
    save_path: Path
    output_format: str
    force_one_page: bool
    safe_temp: bool
    lang: str


@dataclass
class JobResult:
    ok: bool
    #: Save format Hancom actually accepted, e.g. "PDF", "OOXML", "PrintToPDFEx".
    actual_format: str = ""
    #: Localized failure text, shown in the log and written to the CSV.
    message: str = ""
    #: Extra localized log lines to emit before the result line.
    notices: list = field(default_factory=list)


class ConversionBackend(Protocol):
    capabilities: BackendCapabilities

    def preflight(self, lang: str) -> None:
        """Raise ``BackendUnavailable`` when conversion cannot start."""

    def open_session(self, sink, lang: str, options: SessionOptions) -> None:
        """Start the engine (or the remote job) and emit its startup log lines."""

    def session_notes(self, lang: str) -> list:
        """Log payloads emitted after ``run_batch``'s batch header lines."""

    def blocked_reason(self, src_path: Path, output_format: str, lang: str):
        """Localized reason this file must not be opened, or ``None``."""

    def convert(self, job: JobSpec) -> JobResult:
        """Convert one file to one format.

        Must not raise for per-file failures -- return ``JobResult(ok=False)``.
        Raise only when the whole session is lost.
        """

    def cancel(self) -> None:
        """Best-effort stop of in-flight work."""

    def close_session(self) -> None:
        """Tear the engine down. Must not raise."""
