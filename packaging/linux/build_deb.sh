#!/bin/bash
# Build a Debian package of KherveFitting.
#
#   packaging/linux/build_deb.sh [options]
#
# Options:
#   --version X.YY     version recorded in the package (default: 1.80)
#   --revision N       Debian revision appended to the version (default: 1)
#   --python PATH      interpreter providing wxPython and PyInstaller
#                      (default: python3, or $VIRTUAL_ENV/bin/python when set)
#   --output-dir DIR   where the .deb is written (default: <repo>/dist)
#   --skip-pyinstaller reuse an existing dist/KherveFitting folder
#   --no-examples      leave the 32 MB Data-Examples folder out of the package
#   --help             show this text
#
# See packaging/linux/README.md for the prerequisites.
set -euo pipefail

VERSION=1.80
REVISION=1
PYTHON="${VIRTUAL_ENV:+$VIRTUAL_ENV/bin/python}"
PYTHON="${PYTHON:-python3}"
OUTPUT_DIR=
SKIP_PYINSTALLER=0
WITH_EXAMPLES=1

while [ $# -gt 0 ]; do
    case "$1" in
        --version) VERSION="$2"; shift 2 ;;
        --revision) REVISION="$2"; shift 2 ;;
        --python) PYTHON="$2"; shift 2 ;;
        --output-dir) OUTPUT_DIR="$2"; shift 2 ;;
        --skip-pyinstaller) SKIP_PYINSTALLER=1; shift ;;
        --no-examples) WITH_EXAMPLES=0; shift ;;
        --help|-h) sed -n '2,16p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
        *) echo "Unknown option: $1" >&2; exit 2 ;;
    esac
done

PACKAGING_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$PACKAGING_DIR/../.." && pwd)"
OUTPUT_DIR="${OUTPUT_DIR:-$REPO_ROOT/dist}"
BUILD_ROOT="$REPO_ROOT/build/deb"
INSTALL_DIR=/opt/khervefitting
ARCH="$(dpkg --print-architecture)"
PACKAGE="khervefitting_${VERSION}-${REVISION}_${ARCH}"

for tool in dpkg-deb dpkg; do
    command -v "$tool" >/dev/null 2>&1 || { echo "error: $tool is required (apt install dpkg-dev)" >&2; exit 1; }
done
command -v "$PYTHON" >/dev/null 2>&1 || { echo "error: interpreter '$PYTHON' not found" >&2; exit 1; }

cd "$REPO_ROOT"

# 1. Freeze the application with PyInstaller.
if [ "$SKIP_PYINSTALLER" -eq 0 ]; then
    "$PYTHON" -c 'import PyInstaller' 2>/dev/null || {
        echo "error: PyInstaller is not installed for $PYTHON (pip install pyinstaller)" >&2; exit 1; }
    "$PYTHON" -c 'import wx' 2>/dev/null || {
        echo "error: wxPython is not installed for $PYTHON" >&2
        echo "       apt install python3-wxgtk4.0, or pip install wxPython (builds from source)" >&2; exit 1; }
    echo "==> Running PyInstaller"
    rm -rf "$REPO_ROOT/dist/KherveFitting" "$REPO_ROOT/build/KherveFitting"
    "$PYTHON" -m PyInstaller --noconfirm --clean KherveFittingLinux.spec
fi

FROZEN_DIR="$REPO_ROOT/dist/KherveFitting"
[ -x "$FROZEN_DIR/KherveFitting" ] || {
    echo "error: $FROZEN_DIR/KherveFitting not found, run without --skip-pyinstaller" >&2; exit 1; }

# 2. Lay out the package tree.
echo "==> Staging $PACKAGE"
rm -rf "$BUILD_ROOT"
mkdir -p "$BUILD_ROOT$INSTALL_DIR" \
         "$BUILD_ROOT/DEBIAN" \
         "$BUILD_ROOT/usr/bin" \
         "$BUILD_ROOT/usr/share/applications" \
         "$BUILD_ROOT/usr/share/mime/packages" \
         "$BUILD_ROOT/usr/share/doc/khervefitting"

cp -a "$FROZEN_DIR/." "$BUILD_ROOT$INSTALL_DIR/"

# Folders the launcher copies into the user's data directory on first run.
if [ "$WITH_EXAMPLES" -eq 1 ]; then
    cp -a "$REPO_ROOT/Data-Examples" "$BUILD_ROOT$INSTALL_DIR/"
fi
cp -a "$REPO_ROOT/Peaks Library" "$BUILD_ROOT$INSTALL_DIR/"

install -m 755 "$PACKAGING_DIR/launcher.sh" "$BUILD_ROOT/usr/bin/khervefitting"
install -m 644 "$PACKAGING_DIR/khervefitting.desktop" "$BUILD_ROOT/usr/share/applications/khervefitting.desktop"
install -m 644 "$PACKAGING_DIR/khervefitting-mime.xml" "$BUILD_ROOT/usr/share/mime/packages/khervefitting.xml"
install -m 644 "$REPO_ROOT/LICENSE" "$BUILD_ROOT/usr/share/doc/khervefitting/copyright"
install -m 644 "$REPO_ROOT/THIRD_PARTY_LICENSES.txt" "$BUILD_ROOT/usr/share/doc/khervefitting/"

# 3. Icons, in the sizes the desktop menus look for.
echo "==> Generating icons"
"$PYTHON" - "$REPO_ROOT/Icons/Icon.png" "$BUILD_ROOT/usr/share/icons/hicolor" <<'PYICON'
import os
import sys

source, target_root = sys.argv[1], sys.argv[2]
sizes = [16, 22, 24, 32, 48, 64, 128, 256, 512]
try:
    from PIL import Image
except ImportError:
    Image = None

if Image is None:
    # Without Pillow, install the artwork unscaled in the size it happens to be.
    print("   Pillow not available, installing a single unscaled icon")
    target = os.path.join(target_root, '256x256', 'apps')
    os.makedirs(target, exist_ok=True)
    with open(source, 'rb') as src, open(os.path.join(target, 'khervefitting.png'), 'wb') as dst:
        dst.write(src.read())
else:
    icon = Image.open(source).convert('RGBA')
    for size in sizes:
        target = os.path.join(target_root, f'{size}x{size}', 'apps')
        os.makedirs(target, exist_ok=True)
        icon.resize((size, size), Image.LANCZOS).save(os.path.join(target, 'khervefitting.png'))
    print(f"   {len(sizes)} icon sizes written")
PYICON
find "$BUILD_ROOT/usr/share/icons" -type f -exec chmod 644 {} +

# 4. Control file.
INSTALLED_SIZE="$(du -sk "$BUILD_ROOT" | cut -f1)"
sed -e "s/@VERSION@/${VERSION}-${REVISION}/" \
    -e "s/@ARCH@/$ARCH/" \
    -e "s/@SIZE@/$INSTALLED_SIZE/" \
    "$PACKAGING_DIR/debian/control.in" > "$BUILD_ROOT/DEBIAN/control"
install -m 755 "$PACKAGING_DIR/debian/postinst" "$BUILD_ROOT/DEBIAN/postinst"
install -m 755 "$PACKAGING_DIR/debian/postrm" "$BUILD_ROOT/DEBIAN/postrm"

# 5. Build.
echo "==> Building $PACKAGE.deb"
mkdir -p "$OUTPUT_DIR"
dpkg-deb --root-owner-group --build "$BUILD_ROOT" "$OUTPUT_DIR/$PACKAGE.deb"

if command -v lintian >/dev/null 2>&1; then
    echo "==> lintian"
    lintian --no-tag-display-limit "$OUTPUT_DIR/$PACKAGE.deb" || true
fi

echo
echo "Package: $OUTPUT_DIR/$PACKAGE.deb"
echo "Install: sudo apt install $OUTPUT_DIR/$PACKAGE.deb"
