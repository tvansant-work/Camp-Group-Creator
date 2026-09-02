#!/bin/bash
GITHUB_USER="tvansant-work"
GITHUB_REPO="Camp-Group-Creator"
APP_NAME="Camp_Group_Creator"
BASE_DIR="$HOME/Library/Application Support/$APP_NAME"
BRANCH="main"

cd "$BASE_DIR"

echo "=================================================="
echo "  Checking for updates from GitHub..."
echo "=================================================="

# Helper: download a file, compare checksums, report what changed.
# Returns 0 if the file was updated, 1 if unchanged.
fetch_and_report() {
  local URL="$1"
  local DEST="$2"
  local LABEL="$3"
  local TMP="${DEST}.tmp"

  curl -s -L -H "Cache-Control: no-cache" -H "Pragma: no-cache" -o "$TMP" "$URL"

  if [ ! -s "$TMP" ]; then
    echo "  WARNING  $LABEL - download failed, keeping existing version"
    rm -f "$TMP"
    return 1
  fi

  if [ -f "$DEST" ]; then
    OLD_SUM=$(md5 -q "$DEST" 2>/dev/null || md5sum "$DEST" | cut -d' ' -f1)
    NEW_SUM=$(md5 -q "$TMP"  2>/dev/null || md5sum "$TMP"  | cut -d' ' -f1)
    if [ "$OLD_SUM" = "$NEW_SUM" ]; then
      echo "  up to date  $LABEL"
      rm -f "$TMP"
      return 1
    else
      mv "$TMP" "$DEST"
      NEW_BYTES=$(wc -c < "$DEST" | tr -d ' ')
      echo "  UPDATED     $LABEL  (${NEW_BYTES} bytes)"
      return 0
    fi
  else
    mv "$TMP" "$DEST"
    NEW_BYTES=$(wc -c < "$DEST" | tr -d ' ')
    echo "  NEW         $LABEL  (${NEW_BYTES} bytes)"
    return 0
  fi
}

# Resolve the exact latest commit SHA via GitHub API so CDN caching
# can never serve a stale file (?t= does NOT reliably bust GitHub's cache)
echo ""
echo "  Resolving latest commit..."
LATEST_SHA=$(curl -s -f \
  -H "Accept: application/vnd.github+json" \
  "https://api.github.com/repos/$GITHUB_USER/$GITHUB_REPO/commits/$BRANCH" \
  | python3 -c "import sys,json; print(json.load(sys.stdin)['sha'])" 2>/dev/null)

if [ -n "$LATEST_SHA" ]; then
  echo "  Commit: ${LATEST_SHA:0:7}"
  RAW_BASE="https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/$LATEST_SHA"
else
  echo "  WARNING: GitHub API unavailable - falling back to branch URL"
  RAW_BASE="https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/$BRANCH"
fi

echo ""
echo "  File status:"

# Check launcher.sh first — if updated, replace and re-launch so
# the rest of the run uses the new logic immediately.
fetch_and_report "$RAW_BASE/launcher.sh" "launcher.sh" "launcher.sh     "
LAUNCHER_UPDATED=$?
if [ $LAUNCHER_UPDATED -eq 0 ]; then
  chmod +x launcher.sh
  echo ""
  echo "  Launcher updated - restarting with new version..."
  echo "=================================================="
  exec "$BASE_DIR/launcher.sh"
fi

fetch_and_report "$RAW_BASE/app.py"           "app.py"           "app.py          "
fetch_and_report "$RAW_BASE/camp_sync.py"     "camp_sync.py"     "camp_sync.py    "
fetch_and_report "$RAW_BASE/kanban_ui.py"     "kanban_ui.py"     "kanban_ui.py    "
fetch_and_report "$RAW_BASE/requirements.txt" "requirements.txt" "requirements.txt"
fetch_and_report "$RAW_BASE/app_icon.png"     "app_icon.png"     "app_icon.png    "

# Sync dependencies
source venv/bin/activate
pip install -r requirements.txt --quiet

# Re-apply icon to Desktop Shortcut
./venv/bin/python3 - << 'PYEOF'
import Cocoa, os
icon_path = os.path.expanduser("~/Library/Application Support/Camp_Group_Creator/app_icon.png")
file_path = os.path.expanduser("~/Desktop/Camp Groups.command")
if os.path.exists(icon_path) and os.path.exists(file_path):
    icon_image = Cocoa.NSImage.alloc().initWithContentsOfFile_(icon_path)
    if icon_image:
        Cocoa.NSWorkspace.sharedWorkspace().setIcon_forFile_options_(icon_image, file_path, 0)
PYEOF

echo ""
echo "=================================================="
echo "  Launching Camp Groups..."
echo "  KEEP THIS WINDOW OPEN WHILE USING THE APP"
echo "=================================================="
echo ""

python3 -m streamlit run app.py