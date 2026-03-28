#!/bin/bash
set -u

REPO_OWNER="${1:-}"
REPO_NAME="${2:-}"
CURRENT_VERSION="${3:-0.0.0}"
INSTALL_DIR="${4:-}"

APP_SCRIPTS_DIR="$HOME/Library/Application Scripts/com.microsoft.Word"
TEMP_DIR=""
JSON_FILE=""
ZIP_FILE=""
UNPACK_DIR=""

cleanup() {
  if [ -n "${TEMP_DIR:-}" ] && [ -d "${TEMP_DIR:-}" ]; then
    /bin/rm -rf "${TEMP_DIR}" >/dev/null 2>&1 || true
  fi
}

finish() {
  printf '%s\n' "$1"
  cleanup
  exit 0
}

fail() {
  finish "FAILED|$1"
}

normalize_version() {
  local v="$1"
  v="${v#v}"
  v="${v#V}"
  printf '%s' "$v"
}

compare_versions() {
  local a b
  local IFS=.
  read -r -a a <<< "$(normalize_version "$1")"
  read -r -a b <<< "$(normalize_version "$2")"

  local len="${#a[@]}"
  if [ "${#b[@]}" -gt "$len" ]; then
    len="${#b[@]}"
  fi

  local i av bv
  for ((i=0; i<len; i++)); do
    av="${a[i]:-0}"
    bv="${b[i]:-0}"

    av="${av%%[^0-9]*}"
    bv="${bv%%[^0-9]*}"

    [ -n "$av" ] || av=0
    [ -n "$bv" ] || bv=0

    if [ "$((10#$av))" -lt "$((10#$bv))" ]; then
      printf '%s' "-1"
      return
    fi

    if [ "$((10#$av))" -gt "$((10#$bv))" ]; then
      printf '%s' "1"
      return
    fi
  done

  printf '%s' "0"
}

copy_if_exists() {
  local src="$1"
  local dst="$2"
  local mode="$3"

  if [ -f "$src" ]; then
    /bin/cp -f "$src" "$dst" || fail "Failed to copy $(basename "$src")"
    /bin/chmod "$mode" "$dst" >/dev/null 2>&1 || true
  fi
}

[ -n "$REPO_OWNER" ] || fail "Missing repo owner"
[ -n "$REPO_NAME" ] || fail "Missing repo name"
[ -n "$INSTALL_DIR" ] || fail "Missing install dir"

TEMP_DIR="$(/usr/bin/mktemp -d /tmp/kats-update.XXXXXX 2>/dev/null)" || fail "Could not create temp dir"
JSON_FILE="${TEMP_DIR}/latest.json"
ZIP_FILE="${TEMP_DIR}/KATS-Tools-mac-update.zip"
UNPACK_DIR="${TEMP_DIR}/payload"

API_URL="https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/releases/latest"

/usr/bin/curl -fsSL \
  -H "Accept: application/vnd.github+json" \
  -H "X-GitHub-Api-Version: 2022-11-28" \
  -o "$JSON_FILE" \
  "$API_URL" || fail "Could not fetch latest release metadata"

LATEST_TAG="$(/usr/bin/python3 - "$JSON_FILE" <<'PY'
import json, sys
with open(sys.argv[1], encoding='utf-8') as f:
    data = json.load(f)
print(data.get("tag_name", ""))
PY
)"
[ -n "$LATEST_TAG" ] || fail "Could not read latest release tag"

LATEST_VERSION="$(normalize_version "$LATEST_TAG")"

CMP="$(compare_versions "$CURRENT_VERSION" "$LATEST_VERSION")"
if [ "$CMP" -ge 0 ]; then
  finish "UPTODATE"
fi

DOWNLOAD_URL="$(/usr/bin/python3 - "$JSON_FILE" <<'PY'
import json, sys
with open(sys.argv[1], encoding='utf-8') as f:
    data = json.load(f)
for asset in data.get("assets", []):
    if asset.get("name") == "KATS-Tools-mac-update.zip":
        print(asset.get("browser_download_url", ""))
        break
else:
    print("")
PY
)"
[ -n "$DOWNLOAD_URL" ] || fail "Release asset KATS-Tools-mac-update.zip not found"

/bin/mkdir -p "$UNPACK_DIR" || fail "Could not create unpack dir"

/usr/bin/curl -fL "$DOWNLOAD_URL" -o "$ZIP_FILE" || fail "Could not download update package"
/usr/bin/unzip -oq "$ZIP_FILE" -d "$UNPACK_DIR" || fail "Could not extract update package"

/bin/mkdir -p "$INSTALL_DIR" || fail "Could not create install dir"
/bin/mkdir -p "$APP_SCRIPTS_DIR" || fail "Could not create Application Scripts dir"

copy_if_exists "$UNPACK_DIR/KATS-Tools.dotm" "$INSTALL_DIR/KATS-Tools.dotm" 644
copy_if_exists "$UNPACK_DIR/KATS-Version.txt" "$INSTALL_DIR/KATS-Version.txt" 644
copy_if_exists "$UNPACK_DIR/KATSUpdater.applescript" "$APP_SCRIPTS_DIR/KATSUpdater.applescript" 644
copy_if_exists "$UNPACK_DIR/KATSUpdater.sh" "$APP_SCRIPTS_DIR/KATSUpdater.sh" 755
copy_if_exists "$UNPACK_DIR/KATSMail.applescript" "$APP_SCRIPTS_DIR/KATSMail.applescript" 644
copy_if_exists "$UNPACK_DIR/KATSFileOps.applescript" "$APP_SCRIPTS_DIR/KATSFileOps.applescript" 644

finish "INSTALLED|${LATEST_VERSION}"
