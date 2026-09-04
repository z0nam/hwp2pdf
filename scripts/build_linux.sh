#!/usr/bin/env bash
# Build hwp2pdf (GUI) and hwp2pdf-cli and package them for Linux release.
#
#   ./scripts/build_linux.sh              # next yyyy.MM.dd.N for today
#   ./scripts/build_linux.sh 2026.08.28.3 # pin a version (match a Windows/macOS build)
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VENV="$ROOT/.venv-linux"
PIN_VERSION="${1:-}"
BOOTSTRAP_PYTHON="${PYTHON:-python3}"

cd "$ROOT"

# --- preflight ------------------------------------------------------------
"$BOOTSTRAP_PYTHON" - <<'PY'
import sys, tkinter
version = tkinter.TkVersion
print(f"Tk {version} ({sys.executable})")
if version < 8.6:
    raise SystemExit(f"ERROR: Tk {version} is too old; 8.6 or newer is required.")
PY

if [ ! -x "$VENV/bin/python" ]; then
    "$BOOTSTRAP_PYTHON" -m venv "$VENV"
fi
PY_BIN="$VENV/bin/python"

"$PY_BIN" -m pip install --upgrade pip >/dev/null
"$PY_BIN" -m pip install -r "$ROOT/requirements-build.txt" >/dev/null
"$PY_BIN" -m pip install -r "$ROOT/requirements.txt" >/dev/null

# --- version --------------------------------------------------------------
VERSION="$("$PY_BIN" "$ROOT/scripts/set_version.py" $PIN_VERSION | tail -n 1)"
[ -n "$VERSION" ] || { echo "ERROR: could not compute a build version" >&2; exit 1; }
echo "Build version: $VERSION"

[ -x "$ROOT/vendor/rhwp/rhwp" ] || "$ROOT/scripts/fetch_rhwp.sh"

# --- build ----------------------------------------------------------------
rm -rf "$ROOT/dist/hwp2pdf" "$ROOT/dist/hwp2pdf-cli"
"$PY_BIN" -m PyInstaller --clean --noconfirm "$ROOT/hwp2pdf-linux.spec"

GUI="$ROOT/dist/hwp2pdf"
CLI="$ROOT/dist/hwp2pdf-cli"
[ -f "$GUI" ] || { echo "ERROR: expected build output not found: $GUI" >&2; exit 1; }
[ -f "$CLI" ] || { echo "ERROR: expected build output not found: $CLI" >&2; exit 1; }

# --- package --------------------------------------------------------------
ARCH="$(uname -m)"
mkdir -p "$ROOT/release"
TAR_GZ="$ROOT/release/hwp2pdf-linux-$ARCH-$VERSION.tar.gz"
STAGE="$(mktemp -d)"
trap 'rm -rf "$STAGE"' EXIT
mkdir -p "$STAGE/hwp2pdf"
cp "$GUI" "$STAGE/hwp2pdf/"
cp "$CLI" "$STAGE/hwp2pdf/"
cp "$ROOT/THIRD_PARTY_NOTICES.md" "$STAGE/hwp2pdf/"
tar -czf "$TAR_GZ" -C "$STAGE" hwp2pdf

echo
echo "  gui    : $GUI"
echo "  cli    : $CLI"
echo "  tar.gz : $TAR_GZ"
echo
