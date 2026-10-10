#!/usr/bin/env bash
# Run on Debian/Ubuntu (dpkg-deb required). Input must be an onedir bundle
# built on an older glibc baseline, not a self-extracting onefile binary.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
BUNDLE="${1:-$REPO_ROOT/dist/NBC_DocFormat_linux}"
ARCH="${2:-$(dpkg --print-architecture)}"
VERSION="${3:-$(cd "$REPO_ROOT" && python3 -c 'import build; print(build.VERSION)')}"

case "$ARCH" in
  amd64|arm64) ;;
  *) echo "Unsupported Debian architecture: $ARCH" >&2; exit 1 ;;
esac
[[ "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]] || { echo "Invalid version: $VERSION" >&2; exit 1; }
if [ ! -x "$BUNDLE/NBC_DocFormat_linux" ] || [ ! -d "$BUNDLE/_internal" ]; then
  echo "Expected a PyInstaller onedir bundle: $BUNDLE" >&2
  exit 1
fi
command -v dpkg-deb >/dev/null

mkdir -p "$REPO_ROOT/dist"
STAGE="$(mktemp -d "$REPO_ROOT/dist/deb.XXXXXX")"
trap 'rm -rf "$STAGE"' EXIT
chmod 755 "$STAGE"
APP_DIR="$STAGE/opt/nbc-docformat"
mkdir -p "$STAGE/DEBIAN" "$APP_DIR" "$STAGE/usr/bin" \
  "$STAGE/usr/share/applications" "$STAGE/usr/share/icons/hicolor/256x256/apps" \
  "$STAGE/usr/share/doc/nbc-docformat"
cp -a "$BUNDLE/." "$APP_DIR/"
ln -s /opt/nbc-docformat/NBC_DocFormat_linux "$STAGE/usr/bin/NBC_DocFormat_linux"
install -m 644 "$REPO_ROOT/packaging/appimage/NBC_DocFormat.desktop" "$STAGE/usr/share/applications/"
install -m 644 "$REPO_ROOT/assets/icon.png" "$STAGE/usr/share/icons/hicolor/256x256/apps/NBC_DocFormat.png"
install -m 644 "$REPO_ROOT/LICENSE" "$STAGE/usr/share/doc/nbc-docformat/copyright"
install -m 644 "$REPO_ROOT/LICENSE-HISTORY.md" "$STAGE/usr/share/doc/nbc-docformat/"

cat > "$STAGE/DEBIAN/control" <<EOF
Package: nbc-docformat
Version: $VERSION
Section: text
Priority: optional
Architecture: $ARCH
Maintainer: NBC_DocFormat contributors <noreply@github.com>
Depends: libc6 (>= 2.28), libx11-6, libxext6, libxrender1, libxft2, libfontconfig1, libfreetype6, zlib1g
Recommends: xdg-utils, fonts-noto-cjk
Installed-Size: $(du -sk "$STAGE/opt" "$STAGE/usr" | awk '{sum += $1} END {print sum}')
Homepage: https://github.com/0Joahua0/NBC_DocFormat
Description: Local DOCX formatting desktop application
 Format Chinese official documents, edit presets and generate DOCX from text.
EOF

# Desktop/icon caches are shared with other applications; refresh only those
# caches if their helpers exist. No runtime elevation or security-policy changes.
for SCRIPT in postinst postrm; do
  cat > "$STAGE/DEBIAN/$SCRIPT" <<'EOF'
#!/bin/sh
set -e
if command -v update-desktop-database >/dev/null 2>&1; then
  update-desktop-database -q /usr/share/applications || true
fi
if command -v gtk-update-icon-cache >/dev/null 2>&1; then
  gtk-update-icon-cache -q -t /usr/share/icons/hicolor || true
fi
exit 0
EOF
  chmod 755 "$STAGE/DEBIAN/$SCRIPT"
done

OUTPUT="$REPO_ROOT/dist/NBC_DocFormat_linux_${ARCH}.deb"
# xz is supported by the older dpkg versions used by domestic Linux desktops.
dpkg-deb --root-owner-group -Zxz --build "$STAGE" "$OUTPUT"
echo "Built: $OUTPUT"
