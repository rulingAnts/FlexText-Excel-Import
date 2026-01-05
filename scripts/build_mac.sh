#!/usr/bin/env bash
set -euo pipefail

# Build macOS app (onedir + windowed) and create a DMG
# Target: convert_interlinear_gui.py
# Output: dist/Interlinear Converter.app and dist/Interlinear-Converter-<version>.dmg

APP_NAME="Interlinear Converter"
ENTRY="convert_interlinear_gui.py"
BUNDLE_ID="org.rulingants.flextextimport"
DIST_DIR="dist/mac"
BUILD_DIR="build/mac"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="${SCRIPT_DIR%/scripts}"
cd "$PROJECT_ROOT"

# Ensure Python and PyInstaller are available
if ! command -v python3 >/dev/null 2>&1; then
  echo "Python3 not found. Install Xcode Command Line Tools or Python3." >&2
  exit 1
fi

# Optional: use a local venv for reproducible builds
if [ ! -d .venv ]; then
  python3 -m venv .venv
fi
source .venv/bin/activate
python -m pip install --upgrade pip
# Runtime deps used by the app
python -m pip install pyinstaller openpyxl pillow

# Clean previous mac build artifacts only
rm -rf "$BUILD_DIR" "$DIST_DIR"
mkdir -p "$BUILD_DIR" "$DIST_DIR"

# Build the app
# Notes:
# - --collect-all openpyxl ensures openpyxl resources are bundled
# - --windowed suppresses console
# - --osx-bundle-identifier sets bundle id (helps with signing/notarization later)
# Generate icon if missing
if [ ! -f assets/app.icns ]; then
  python scripts/generate_icons.py || true
fi

pyinstaller \
  --onedir \
  --windowed \
  --name "$APP_NAME" \
  --collect-all openpyxl \
  --osx-bundle-identifier "$BUNDLE_ID" \
  --distpath "$DIST_DIR" \
  --workpath "$BUILD_DIR" \
  --specpath "$BUILD_DIR" \
  --icon "$PROJECT_ROOT/assets/app.icns" \
  "$ENTRY"

# Locate the built .app bundle
APP_BUNDLE=""
if [ -d "$DIST_DIR/$APP_NAME.app" ]; then
  APP_BUNDLE="$DIST_DIR/$APP_NAME.app"
else
  # Fallback: find any .app produced
  APP_BUNDLE=$(ls -d "$DIST_DIR"/*.app 2>/dev/null | head -n 1 || true)
fi

if [ -z "$APP_BUNDLE" ] || [ ! -d "$APP_BUNDLE" ]; then
  echo "Build succeeded but no .app bundle found in $DIST_DIR/." >&2
  exit 2
fi

# Prepare DMG staging
STAGING="$DIST_DIR/dmg-staging"
rm -rf "$STAGING"
mkdir -p "$STAGING"
cp -R "$APP_BUNDLE" "$STAGING/"
ln -s /Applications "$STAGING/Applications"

# Version from latest tag if available
VERSION=$(git describe --tags --abbrev=0 2>/dev/null || echo "latest")
DMG_NAME="Interlinear-Converter-$VERSION.dmg"
VOL_NAME="$APP_NAME"

# Create a compressed DMG
hdiutil create -volname "$VOL_NAME" -srcfolder "$STAGING" -ov -format UDZO "$DIST_DIR/$DMG_NAME"

# Done
echo "Built app: $APP_BUNDLE"
echo "Created DMG: $DIST_DIR/$DMG_NAME"
