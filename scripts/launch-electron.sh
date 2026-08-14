#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
APP_ELECTRON="/Applications/ADO Master Electron.app"
SRC_ELECTRON="$ROOT/node_modules/electron/dist/Electron.app"
BIN="$APP_ELECTRON/Contents/MacOS/Electron"
MARKER="$APP_ELECTRON/Contents/Resources/.ado-master-electron"

sync_app() {
  if [[ ! -d "$SRC_ELECTRON" ]]; then
    echo "Electron is not installed. Run: npm install" >&2
    exit 1
  fi

  # Only resync when missing — keep a stable Applications path
  if [[ -d "$APP_ELECTRON" && -f "$MARKER" ]]; then
    return 0
  fi

  rm -rf "$APP_ELECTRON"
  cp -R "$SRC_ELECTRON" "$APP_ELECTRON"
  xattr -dr com.apple.quarantine "$APP_ELECTRON" 2>/dev/null || true
  /usr/libexec/PlistBuddy -c "Set :CFBundleName ADO Master Electron" "$APP_ELECTRON/Contents/Info.plist" 2>/dev/null || true
  /usr/libexec/PlistBuddy -c "Set :CFBundleDisplayName ADO Master Electron" "$APP_ELECTRON/Contents/Info.plist" 2>/dev/null || true
  /usr/libexec/PlistBuddy -c "Set :CFBundleIdentifier com.adomaster.list.electron" "$APP_ELECTRON/Contents/Info.plist" 2>/dev/null || true
  mkdir -p "$APP_ELECTRON/Contents/Resources"
  echo "ado-master-electron" > "$MARKER"
  echo "Installed $APP_ELECTRON"
}

sync_app

if [[ "${1:-}" == "--sync-only" ]]; then
  exit 0
fi

cd "$ROOT"
exec "$BIN" "$ROOT" "$@"
