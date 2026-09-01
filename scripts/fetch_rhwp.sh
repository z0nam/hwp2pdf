#!/usr/bin/env bash
# fetch_rhwp.sh — download the pinned rhwp release binary for this platform,
# verify its SHA-256 against the release SHA256SUMS.txt, and install it to
# vendor/rhwp/rhwp.
#
# rhwp is only needed for the optional local fallback that renders PDFs when
# local Hancom Office or the conversion server cannot start. hwp2pdf works
# without it.
#
# Usage:  scripts/fetch_rhwp.sh [version]
# Requires: curl, tar, shasum. No GitHub CLI and no account -- the releases are
# public, so a normal user can run this as-is.
set -euo pipefail

RHWP_VERSION="${1:-v0.8.4}"
RHWP_REPO="edwardkim/rhwp"
BASE="https://github.com/$RHWP_REPO/releases/download/$RHWP_VERSION"

repo_root=$(cd "$(dirname "$0")/.." && pwd)
dest_dir="$repo_root/vendor/rhwp"

os=$(uname -s)
arch=$(uname -m)
case "$os/$arch" in
  Darwin/arm64)  slug="macos-aarch64" ;;
  Darwin/x86_64) slug="macos-x86_64" ;;
  Linux/x86_64)  slug="linux-x86_64" ;;
  *) echo "error: no prebuilt rhwp for $os/$arch (see github.com/$RHWP_REPO/releases)" >&2; exit 1 ;;
esac
asset="rhwp-$RHWP_VERSION-$slug.tar.gz"

tmp=$(mktemp -d)
trap 'rm -rf "$tmp"' EXIT

echo "Downloading $asset ($RHWP_VERSION)..."
curl -fsSL --retry 3 -o "$tmp/$asset" "$BASE/$asset"
curl -fsSL --retry 3 -o "$tmp/SHA256SUMS.txt" "$BASE/SHA256SUMS.txt"

echo "Verifying checksum..."
want=$(awk -v a="$asset" '$2==a{print $1}' "$tmp/SHA256SUMS.txt")
got=$(shasum -a 256 "$tmp/$asset" | awk '{print $1}')
if [[ -z "$want" || "$want" != "$got" ]]; then
  echo "error: checksum mismatch for $asset" >&2
  echo "  want=$want" >&2
  echo "  got =$got" >&2
  exit 1
fi
echo "  ok: $got"

echo "Extracting to $dest_dir..."
mkdir -p "$dest_dir"
tar xzf "$tmp/$asset" -C "$tmp"
# The tarball holds a top-level rhwp/ directory with the binary and its LICENSE.
cp "$tmp/rhwp/rhwp" "$dest_dir/rhwp"
cp "$tmp/rhwp/LICENSE" "$dest_dir/LICENSE" 2>/dev/null || true
cp "$tmp/SHA256SUMS.txt" "$dest_dir/SHA256SUMS.txt"
chmod +x "$dest_dir/rhwp"
# Gatekeeper quarantines anything downloaded; the binary is unsigned.
xattr -d com.apple.quarantine "$dest_dir/rhwp" 2>/dev/null || true

echo "Installed: $("$dest_dir/rhwp" --help 2>&1 | head -1)"
echo
echo "hwp2pdf finds this automatically. Enable the fallback with --rhwp-fallback,"
echo "or the matching checkbox in the GUI options."
