#!/bin/bash

# Configuration
GITHUB_USER="tvansant-work"
GITHUB_REPO="Camp-Group-Creator"
APP_NAME="Camp_Group_Creator"
BRANCH="main"
RAW_BASE="https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/$BRANCH"
SHORTCUT="$HOME/Desktop/Camp Groups.command"
BASE_DIR="$HOME/Library/Application Support/$APP_NAME"

echo "=================================================="
echo "  Camp Groups - Installer"
echo "=================================================="

# 1. Setup folder and Python environment
echo ""
echo "Step 1/3: Preparing Python environment..."
mkdir -p "$BASE_DIR"
cd "$BASE_DIR"
if [ ! -d "venv" ]; then
    python3 -m venv venv
fi

# 2. Download all app files from GitHub
echo ""
echo "Step 2/3: Downloading app files from GitHub..."
curl -s -L -o app.py           "$RAW_BASE/app.py"
curl -s -L -o camp_sync.py     "$RAW_BASE/camp_sync.py"
curl -s -L -o kanban_ui.py     "$RAW_BASE/kanban_ui.py"
curl -s -L -o requirements.txt "$RAW_BASE/requirements.txt"
curl -s -L -o app_icon.png     "$RAW_BASE/app_icon.png"

# launcher.sh is downloaded from GitHub rather than embedded here.
# This means you can push launcher updates to all users without
# them ever needing to re-run install.sh.
curl -s -L -o launcher.sh "$RAW_BASE/launcher.sh"
chmod +x launcher.sh

echo "  app.py              ($(wc -c < app.py | tr -d ' ') bytes)"
echo "  camp_sync.py        ($(wc -c < camp_sync.py | tr -d ' ') bytes)"
echo "  kanban_ui.py        ($(wc -c < kanban_ui.py | tr -d ' ') bytes)"
echo "  requirements.txt"
echo "  app_icon.png"
echo "  launcher.sh"

# 3. Install dependencies
echo ""
echo "  Installing Python dependencies..."
source venv/bin/activate
pip install -r requirements.txt --quiet
./venv/bin/pip install pyobjc-framework-Cocoa --quiet

# 4. Create Desktop Shortcut pointing to launcher.sh
echo ""
echo "Step 3/3: Creating Desktop Shortcut..."
echo "\"$BASE_DIR/launcher.sh\"" > "$SHORTCUT"
chmod +x "$SHORTCUT"

# 5. Apply Icon
./venv/bin/python3 - << 'PYEOF'
import Cocoa, os
icon_path = os.path.expanduser("~/Library/Application Support/Camp_Group_Creator/app_icon.png")
file_path = os.path.expanduser("~/Desktop/Camp Groups.command")
if os.path.exists(icon_path) and os.path.exists(file_path):
    img = Cocoa.NSImage.alloc().initWithContentsOfFile_(icon_path)
    if img: Cocoa.NSWorkspace.sharedWorkspace().setIcon_forFile_options_(img, file_path, 0)
PYEOF

echo ""
echo "------------------------------------------------"
echo "INSTALLATION COMPLETE!"
echo ""
echo "Every time the app launches it will check GitHub"
echo "for updates to ALL files, including launcher.sh."
echo "------------------------------------------------"