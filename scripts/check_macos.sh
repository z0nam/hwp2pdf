#!/usr/bin/env bash
# Preflight for macOS development, mirroring scripts/check_windows.ps1.
set -uo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PYTHON="${PYTHON:-python3}"
FAILED=0

check() {
    if [ "$1" -eq 0 ]; then printf '[ok  ] %s\n' "$2"; else printf '[FAIL] %s\n' "$2"; FAILED=1; fi
}

echo "== hwp2pdf macOS check =="
[ "$(uname -s)" = "Darwin" ]; check $? "running on macOS"

"$PYTHON" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)'
check $? "python >= 3.10 ($("$PYTHON" -V 2>&1))"

TK="$("$PYTHON" -c 'import tkinter; print(tkinter.TkVersion)' 2>/dev/null)"
[ -n "$TK" ]; check $? "tkinter available (Tk ${TK:-none})"
"$PYTHON" -c 'import tkinter; raise SystemExit(0 if tkinter.TkVersion >= 8.6 else 1)' 2>/dev/null
check $? "Tk >= 8.6 (Apple's /usr/bin/python3 ships 8.5)"

"$PYTHON" -c 'import tkinterdnd2' 2>/dev/null
check $? "tkinterdnd2 importable (drag and drop)"

PYTHONPATH="$ROOT/src" "$PYTHON" -c 'import hwp2pdf.app' 2>/dev/null
check $? "hwp2pdf.app imports"

PYTHONPATH="$ROOT/src" "$PYTHON" -c '
from hwp2pdf.constants import output_extension
raise SystemExit(0 if output_extension("DOCX") == ".docx" else 1)'
check $? "output_extension(\"DOCX\") == .docx"

SERVER="$(PYTHONPATH="$ROOT/src" "$PYTHON" -c '
from hwp2pdf import config
print(config.server_settings().get("url", ""))' 2>/dev/null)"

if [ -n "$SERVER" ]; then
    PYTHONPATH="$ROOT/src" "$PYTHON" -c '
import sys
from hwp2pdf import config
from hwp2pdf.backends.remote_http import RemoteHttpBackend
backend = RemoteHttpBackend(config.server_settings())
backend.preflight("ko")
print("   server:", backend.capabilities_payload.get("version"),
      "queue:", backend.capabilities_payload.get("queue_depth"))
'
    check $? "conversion server reachable ($SERVER)"
else
    echo "[skip] no conversion server configured (settings.json / \$HWP2PDF_SERVER_URL)"
fi

echo
[ "$FAILED" -eq 0 ] && echo "All checks passed." || echo "Some checks failed."
exit "$FAILED"
