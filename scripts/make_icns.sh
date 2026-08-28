#!/usr/bin/env bash
# Build assets/hwp_to_pdf_final.icns from the Windows .ico.
#
# The .ico tops out at 256x256, so the 512pt slots are upscaled from it; macOS
# only ever needed a source that is sharp at the sizes people actually see.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
ICO="$ROOT/assets/hwp_to_pdf_final.ico"
ICNS="$ROOT/assets/hwp_to_pdf_final.icns"
PYTHON="${PYTHON:-python3}"

[ -f "$ICO" ] || { echo "ERROR: missing $ICO" >&2; exit 1; }
command -v iconutil >/dev/null || { echo "ERROR: iconutil not found (macOS only)" >&2; exit 1; }
command -v sips >/dev/null || { echo "ERROR: sips not found (macOS only)" >&2; exit 1; }

WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT
ICONSET="$WORK/hwp2pdf.iconset"
mkdir -p "$ICONSET"

# Pull the largest PNG-compressed image out of the icon directory.
"$PYTHON" - "$ICO" "$WORK/base.png" <<'PY'
import struct, sys

data = open(sys.argv[1], "rb").read()
count = struct.unpack_from("<H", data, 4)[0]
best = None
for index in range(count):
    width, height, _colors, _r, _planes, _bpp, size, offset = struct.unpack_from(
        "<BBBBHHII", data, 6 + 16 * index
    )
    width, height = width or 256, height or 256
    if best is None or width * height > best[0]:
        best = (width * height, offset, size, width)

_area, offset, size, width = best
blob = data[offset:offset + size]
if not blob.startswith(b"\x89PNG"):
    raise SystemExit("ERROR: largest icon entry is not PNG-compressed")
open(sys.argv[2], "wb").write(blob)
print(f"extracted {width}x{width} PNG ({size} bytes)")
PY

for spec in "16 icon_16x16" "32 icon_16x16@2x" "32 icon_32x32" "64 icon_32x32@2x" \
            "128 icon_128x128" "256 icon_128x128@2x" "256 icon_256x256" \
            "512 icon_256x256@2x" "512 icon_512x512" "1024 icon_512x512@2x"; do
    set -- $spec
    sips -z "$1" "$1" "$WORK/base.png" --out "$ICONSET/$2.png" >/dev/null
done

iconutil -c icns "$ICONSET" -o "$ICNS"
echo "wrote $ICNS ($(stat -f%z "$ICNS") bytes)"
