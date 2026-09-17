#!/bin/bash
# Build "Video Merger.app" — a native SwiftUI app, no Xcode project needed.
set -euo pipefail
cd "$(dirname "$0")"

APP_NAME="Video Merger"
BUNDLE_ID="com.adamlewis.videomerger"
EXEC_NAME="VideoMerger"
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
  -framework SwiftUI -framework AppKit -framework Foundation \
  Sources/*.swift \
  -o "$APP/Contents/MacOS/$EXEC_NAME"

# ── Bundle ffmpeg/ffprobe so the app works on a Mac with no Homebrew. The old
#    shell-script version shipped a static ffmpeg; prefer that if it is still
#    installed, since it needs no dylibs at all. Anything dynamic gets its
#    libraries dragged in alongside it.
echo "› Bundling ffmpeg"
BIN_DIR="$APP/Contents/Resources/bin"
LIB_DIR="$APP/Contents/Resources/lib"
mkdir -p "$BIN_DIR"

# Where a self-contained ffmpeg might already be: the shell-script version kept
# it loose in Resources, this build puts it in Resources/bin, and a previous
# build of this app is the third place to look. Checking all three keeps the
# build reproducible after "Video Merger.app" in /Applications is replaced by
# this one.
STATIC_CANDIDATES=(
  "/Applications/$APP_NAME.app/Contents/Resources/bin/ffmpeg"
  "/Applications/$APP_NAME.app/Contents/Resources/ffmpeg"
  "$BUILD_DIR/.ffmpeg-static"
)
STATIC_FFMPEG=""

# A static build links nothing outside /usr/lib and /System, so it can simply
# be copied in.
is_static() { ! otool -L "$1" 2>/dev/null | tail -n +2 | awk '{print $1}' \
  | grep -qvE '^(/usr/lib/|/System/)'; }

take() {  # take <name> <source path>
  [ -n "${2:-}" ] && [ -x "${2:-}" ] || return 1
  cp -f "$2" "$BIN_DIR/$1"
  chmod u+wx "$BIN_DIR/$1"
}

for c in "${STATIC_CANDIDATES[@]}"; do
  if [ -x "$c" ] && is_static "$c"; then STATIC_FFMPEG="$c"; break; fi
done

if [ -n "$STATIC_FFMPEG" ]; then
  echo "  ffmpeg: self-contained build from $STATIC_FFMPEG"
  take ffmpeg "$STATIC_FFMPEG" || true
  # Keep a copy outside the .app so the next build still finds one even if
  # every installed copy goes away.
  cp -f "$STATIC_FFMPEG" "$BUILD_DIR/.ffmpeg-static" 2>/dev/null || true
else
  take ffmpeg "$(command -v ffmpeg || true)" \
    && echo "  ffmpeg: $(command -v ffmpeg)" \
    || echo "  ! ffmpeg not found — the app will look for a system install instead."
fi
# ffprobe reads the durations the progress bar measures against, so it is worth
# bundling even when ffmpeg came from somewhere else.
take ffprobe "$(command -v ffprobe || true)" \
  && echo "  ffprobe: $(command -v ffprobe)" \
  || echo "  ! ffprobe not found — progress will be indeterminate off this Mac."

# Walk the dependency graph of whatever landed in bin/, copying each non-system
# dylib in. A static binary contributes nothing and costs one otool call.
DYNAMIC=0
for b in "$BIN_DIR"/*; do
  [ -f "$b" ] || continue
  is_static "$b" || DYNAMIC=1
done

if [ "$DYNAMIC" = 1 ]; then
  mkdir -p "$LIB_DIR"
  SEEN="$BUILD_DIR/.deps_seen"; QUEUE="$BUILD_DIR/.deps_queue"
  : > "$SEEN"; : > "$QUEUE"
  for b in "$BIN_DIR"/*; do [ -f "$b" ] && echo "$b" >> "$QUEUE"; done
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
  echo "  bundled $(ls -1 "$LIB_DIR" 2>/dev/null | wc -l | tr -d ' ') libraries"

  # Repoint every reference at the copies inside the bundle.
  for b in "$BIN_DIR"/*; do
    [ -f "$b" ] || continue
    otool -L "$b" | tail -n +2 | awk '{print $1}' | while read -r lib; do
      case "$lib" in /usr/lib/*|/System/*|@*|"") continue ;; esac
      install_name_tool -change "$lib" "@executable_path/../lib/$(basename "$lib")" "$b"
    done
  done
  for f in "$LIB_DIR"/*.dylib; do
    [ -f "$f" ] || continue
    install_name_tool -id "@loader_path/$(basename "$f")" "$f"
    otool -L "$f" | tail -n +2 | awk '{print $1}' | while read -r lib; do
      case "$lib" in /usr/lib/*|/System/*|@*|"") continue ;; esac
      install_name_tool -change "$lib" "@loader_path/$(basename "$lib")" "$f"
    done
  done
  rm -f "$SEEN" "$QUEUE"
fi

# install_name_tool invalidates signatures, so re-sign the pieces.
for f in "$LIB_DIR"/*.dylib; do [ -f "$f" ] && codesign --force --sign - "$f" 2>/dev/null; done
for b in "$BIN_DIR"/*; do [ -f "$b" ] && codesign --force --sign - "$b" 2>/dev/null; done

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
	<key>CFBundleShortVersionString</key><string>2.0</string>
	<key>CFBundleVersion</key><string>2</string>
	<key>LSMinimumSystemVersion</key><string>13.0</string>
	<key>LSApplicationCategoryType</key><string>public.app-category.video</string>
	<key>NSHighResolutionCapable</key><true/>
	<key>NSPrincipalClass</key><string>NSApplication</string>
	<key>NSDesktopFolderUsageDescription</key><string>To merge videos kept on the Desktop.</string>
	<key>NSDocumentsFolderUsageDescription</key><string>To merge videos kept in Documents.</string>
	<key>NSDownloadsFolderUsageDescription</key><string>To merge videos kept in Downloads.</string>
	<key>NSRemovableVolumesUsageDescription</key><string>To merge videos kept on external drives.</string>
	<!-- Lets videos and DVD folders be dropped on the Dock icon. Rank
	     Alternate, so this never becomes the default player for them. -->
	<key>CFBundleDocumentTypes</key>
	<array>
		<dict>
			<key>CFBundleTypeName</key><string>Video</string>
			<key>CFBundleTypeRole</key><string>Viewer</string>
			<key>LSHandlerRank</key><string>Alternate</string>
			<key>LSItemContentTypes</key>
			<array><string>public.movie</string><string>public.folder</string></array>
		</dict>
	</array>
</dict>
</plist>
PLIST

echo "› Code-signing (ad-hoc)"
codesign --force --sign - "$APP"

echo "› Done: $APP"
du -sh "$APP" | awk '{print "  bundle size: " $1}'
