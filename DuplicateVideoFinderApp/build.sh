#!/bin/bash
# Build "Duplicate Video Finder.app" — a native SwiftUI app, no Xcode project needed.
set -euo pipefail
cd "$(dirname "$0")"

APP_NAME="Duplicate Video Finder"
BUNDLE_ID="com.adamlewis.duplicate-video-finder"
EXEC_NAME="DuplicateVideoFinder"
BUILD_DIR="build"
APP="$BUILD_DIR/$APP_NAME.app"

DEPLOY_TARGET="arm64-apple-macos13.0"
SDK="$(xcrun --show-sdk-path --sdk macosx)"

echo "› Cleaning"
rm -rf "$APP"
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"

echo "› Rendering icon"
ICONSET="$BUILD_DIR/AppIcon.iconset"
rm -rf "$ICONSET"
swift make-icon.swift "$ICONSET"
iconutil -c icns "$ICONSET" -o "$APP/Contents/Resources/AppIcon.icns"

echo "› Compiling (swiftc)"
swiftc \
  -O \
  -parse-as-library \
  -swift-version 5 \
  -target "$DEPLOY_TARGET" \
  -sdk "$SDK" \
  -framework SwiftUI -framework AppKit -framework Foundation -framework CryptoKit \
  Sources/*.swift \
  -o "$APP/Contents/MacOS/$EXEC_NAME"

# ── Bundle ffmpeg/ffprobe plus every non-system dylib they need, so the app
#    works on a Mac with no Homebrew installed. ~22MB of libraries.
echo "› Bundling ffmpeg"
BIN_DIR="$APP/Contents/Resources/bin"
LIB_DIR="$APP/Contents/Resources/lib"
FFMPEG="$(command -v ffmpeg || true)"
FFPROBE="$(command -v ffprobe || true)"

if [ -z "$FFMPEG" ] || [ -z "$FFPROBE" ]; then
  echo "  ! ffmpeg/ffprobe not on PATH — not bundling."
  echo "    The app will look for a system install at runtime instead."
else
  mkdir -p "$BIN_DIR" "$LIB_DIR"
  cp -f "$FFMPEG" "$BIN_DIR/ffmpeg"
  cp -f "$FFPROBE" "$BIN_DIR/ffprobe"
  chmod u+w "$BIN_DIR/ffmpeg" "$BIN_DIR/ffprobe"

  # Walk the dependency graph, copying each non-system dylib in.
  SEEN="$BUILD_DIR/.deps_seen"; QUEUE="$BUILD_DIR/.deps_queue"
  : > "$SEEN"; : > "$QUEUE"
  echo "$BIN_DIR/ffmpeg" >> "$QUEUE"
  echo "$BIN_DIR/ffprobe" >> "$QUEUE"
  while [ -s "$QUEUE" ]; do
    cur="$(head -1 "$QUEUE")"
    sed -i '' '1d' "$QUEUE"
    otool -L "$cur" 2>/dev/null | tail -n +2 | awk '{print $1}' | while read -r lib; do
      case "$lib" in /usr/lib/*|/System/*|@*|"") continue ;; esac
      grep -qxF "$lib" "$SEEN" && continue
      echo "$lib" >> "$SEEN"
      base="$(basename "$lib")"
      cp -f "$lib" "$LIB_DIR/$base"
      chmod u+w "$LIB_DIR/$base"
      echo "$LIB_DIR/$base" >> "$QUEUE"
    done
  done
  echo "  bundled $(ls -1 "$LIB_DIR" | wc -l | tr -d ' ') libraries"

  # Repoint every reference at the copies inside the bundle.
  for b in "$BIN_DIR/ffmpeg" "$BIN_DIR/ffprobe"; do
    otool -L "$b" | tail -n +2 | awk '{print $1}' | while read -r lib; do
      case "$lib" in /usr/lib/*|/System/*|@*|"") continue ;; esac
      install_name_tool -change "$lib" "@executable_path/../lib/$(basename "$lib")" "$b"
    done
  done
  for f in "$LIB_DIR"/*.dylib; do
    install_name_tool -id "@loader_path/$(basename "$f")" "$f"
    otool -L "$f" | tail -n +2 | awk '{print $1}' | while read -r lib; do
      case "$lib" in /usr/lib/*|/System/*|@*|"") continue ;; esac
      install_name_tool -change "$lib" "@loader_path/$(basename "$lib")" "$f"
    done
  done

  # install_name_tool invalidates signatures, so re-sign the pieces.
  for f in "$LIB_DIR"/*.dylib; do codesign --force --sign - "$f" 2>/dev/null; done
  codesign --force --sign - "$BIN_DIR/ffmpeg" 2>/dev/null
  codesign --force --sign - "$BIN_DIR/ffprobe" 2>/dev/null
  rm -f "$SEEN" "$QUEUE"
fi

echo "› Writing Info.plist"
cat > "$APP/Contents/Info.plist" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
	<key>CFBundleName</key><string>$APP_NAME</string>
	<key>CFBundleDisplayName</key><string>$APP_NAME</string>
	<key>CFBundleExecutable</key><string>$EXEC_NAME</string>
	<key>CFBundleIdentifier</key><string>$BUNDLE_ID</string>
	<key>CFBundlePackageType</key><string>APPL</string>
	<key>CFBundleIconFile</key><string>AppIcon</string>
	<key>CFBundleIconName</key><string>AppIcon</string>
	<key>CFBundleInfoDictionaryVersion</key><string>6.0</string>
	<key>CFBundleShortVersionString</key><string>1.0</string>
	<key>CFBundleVersion</key><string>1</string>
	<key>LSMinimumSystemVersion</key><string>13.0</string>
	<key>LSApplicationCategoryType</key><string>public.app-category.utilities</string>
	<key>NSHighResolutionCapable</key><true/>
	<key>NSPrincipalClass</key><string>NSApplication</string>
	<key>NSDesktopFolderUsageDescription</key><string>To scan videos kept on the Desktop for duplicates.</string>
	<key>NSDocumentsFolderUsageDescription</key><string>To scan videos kept in Documents for duplicates.</string>
	<key>NSDownloadsFolderUsageDescription</key><string>To scan videos kept in Downloads for duplicates.</string>
	<key>NSRemovableVolumesUsageDescription</key><string>To scan videos kept on external drives for duplicates.</string>
</dict>
</plist>
PLIST

echo "› Code-signing (ad-hoc)"
codesign --force --sign - "$APP"

echo "› Done: $APP"
du -sh "$APP" | awk '{print "  bundle size: " $1}'
