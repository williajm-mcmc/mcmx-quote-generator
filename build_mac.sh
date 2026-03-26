#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────────────────────
#  Build script for MCMX Quote Generator  —  macOS
#  Must be run ON a Mac.  Copy the entire project folder to the Mac first.
#
#  Prerequisites (run once):
#    pip install pyinstaller PyQt6 python-docx docxtpl
#
#  Optional – convert logo.png to an .icns icon:
#    mkdir MyIcon.iconset
#    sips -z 1024 1024 logo.png --out MyIcon.iconset/icon_512x2.png
#    iconutil -c icns MyIcon.iconset -o logo.icns
#    # then change icon='logo.png' -> icon='logo.icns' in the .spec file
# ─────────────────────────────────────────────────────────────────────────────

set -e

echo "[1/3] Cleaning previous build..."
rm -rf build dist

echo "[2/3] Running PyInstaller..."
pyinstaller quote_generator.spec

echo "[3/3] Done."
echo ""
echo "Output:  dist/QuoteGenerator.app"
echo "Zip it:  cd dist && zip -r QuoteGenerator-mac.zip QuoteGenerator.app"
