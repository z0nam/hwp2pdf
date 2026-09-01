#!/usr/bin/env bash
# Build hwp2pdf.app plus the hwp2pdf-cli binary and zip them for release.
#
#   ./scripts/build_macos.sh              # next yyyy.MM.dd.N for today
#   ./scripts/build_macos.sh 2026.08.28.3 # pin a version (match a Windows build)
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VENV="$ROOT/.venv-macos"
PIN_VERSION="${1:-}"
BOOTSTRAP_PYTHON="${PYTHON:-python3}"

cd "$ROOT"

# --- preflight ------------------------------------------------------------
# Apple's /usr/bin/python3 ships Tk 8.5, which renders badly on modern macOS.
# Use a python.org framework build or Homebrew python + python-tk.
if [ "$("$BOOTSTRAP_PYTHON" -c 'import sys; print(sys.executable)')" = "/usr/bin/python3" ]; then
    echo "ERROR: /usr/bin/python3 ships Tk 8.5. Install a python.org build or" >&2
    echo "       'brew install python-tk@3.13', then re-run with PYTHON=<path>." >&2
    exit 1
fi

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

# --- version + icon -------------------------------------------------------
VERSION="$("$PY_BIN" "$ROOT/scripts/set_version.py" $PIN_VERSION | tail -n 1)"
[ -n "$VERSION" ] || { echo "ERROR: could not compute a build version" >&2; exit 1; }
echo "Build version: $VERSION"

[ -f "$ROOT/assets/hwp_to_pdf_final.icns" ] || PYTHON="$PY_BIN" "$ROOT/scripts/make_icns.sh"

# --- build ----------------------------------------------------------------
rm -rf "$ROOT/dist/hwp2pdf.app" "$ROOT/dist/hwp2pdf" "$ROOT/dist/hwp2pdf-cli"
"$PY_BIN" -m PyInstaller --clean --noconfirm "$ROOT/hwp2pdf-macos.spec"

APP="$ROOT/dist/hwp2pdf.app"
CLI="$ROOT/dist/hwp2pdf-cli"
[ -d "$APP" ] || { echo "ERROR: expected build output not found: $APP" >&2; exit 1; }
[ -f "$CLI" ] || { echo "ERROR: expected build output not found: $CLI" >&2; exit 1; }

# --- sign -----------------------------------------------------------------
# Ad-hoc only: there is no Apple Developer identity, so Gatekeeper still warns
# on first launch. See the macOS section of README.md.
codesign --force --deep --sign - --timestamp=none "$APP"
codesign --force --sign - --timestamp=none "$CLI"
xattr -cr "$APP" "$CLI"
codesign --verify --deep --strict --verbose=2 "$APP"

# --- package --------------------------------------------------------------
ARCH="$(uname -m)"
mkdir -p "$ROOT/release"
ZIP="$ROOT/release/hwp2pdf-macos-$ARCH-$VERSION.zip"
STAGE="$(mktemp -d)"
trap 'rm -rf "$STAGE"' EXIT
cp -R "$APP" "$STAGE/"
cp "$CLI" "$STAGE/"
cp "$ROOT/THIRD_PARTY_NOTICES.md" "$STAGE/"
# ditto, not zip: it preserves the code signature and symlinks inside the app.
# Archive the staging directory's contents, not its random mktemp directory name.
ditto -c -k --sequesterRsrc "$STAGE/" "$ZIP"

echo
echo "  app : $APP"
echo "  cli : $CLI"
echo "  zip : $ZIP"
echo
echo "Gatekeeper (unsigned by design):"
spctl -a -vv "$APP" 2>&1 || true
