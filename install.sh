#!/bin/bash

# Configuration
GITHUB_USER="tvansant-work"
GITHUB_REPO="Camp-Group-Creator"
APP_NAME="Camp_Group_Creator"

# 1. Setup Folders
BASE_DIR="$HOME/Library/Application Support/$APP_NAME"
mkdir -p "$BASE_DIR"
cd "$BASE_DIR"

echo "Step 1/3: Preparing Python Environment..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
fi

# 2. Create the Launcher
cat << EOF > launcher.sh
#!/bin/bash
cd "$BASE_DIR"
echo "=================================================="
echo " 🔄 Checking for Updates..."
echo "=================================================="
curl -s -o app.py "https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/main/app.py?t=\$(date +%s)"
curl -s -o requirements.txt "https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/main/requirements.txt?t=\$(date +%s)"

source venv/bin/activate
pip install -r requirements.txt --quiet

echo " ✅ App Ready. Launching Browser..."
echo " ⚠️  KEEP THIS WINDOW OPEN WHILE USING THE APP"
echo "=================================================="

if [ ! -f "app_icon.png" ]; then
    curl -s -o app_icon.png "https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/main/app_icon.png"
fi

python3 -m streamlit run app.py
EOF

chmod +x launcher.sh

# 3. Create Shortcut and Icon
echo "Step 2/3: Creating Desktop Shortcut..."
SHORTCUT="$HOME/Desktop/Camp Groups.command"
echo "\"$BASE_DIR/launcher.sh\"" > "$SHORTCUT"
chmod +x "$SHORTCUT"

echo "Step 3/3: Finalising Icon..."
curl -s -o app_icon.png "https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/main/app_icon.png"
./venv/bin/pip install pyobjc-framework-Cocoa --quiet
./venv/bin/python3 - << 'PYEOF'
import Cocoa, os
icon_path = os.path.expanduser("~/Library/Application Support/Camp_Group_Creator/app_icon.png")
file_path = os.path.expanduser("~/Desktop/Camp Groups.command")
if os.path.exists(icon_path) and os.path.exists(file_path):
    img = Cocoa.NSImage.alloc().initWithContentsOfFile_(icon_path)
    if img: Cocoa.NSWorkspace.sharedWorkspace().setIcon_forFile_options_(img, file_path, 0)
PYEOF

echo "INSTALL COMPLETE! You can close this window and open 'Camp Groups' on your desktop."