#!/bin/bash
# DakeSCI Clone — Mac Installer
# Double-click this file (or run: bash install-mac.command) to set up.
# After running, you'll need to enable the add-in inside PowerPoint.

set -e

# ---- Locate script dir (works whether double-clicked or invoked) ----
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# ---- Bounce if no DakeSCI.bas ----
if [[ ! -f "DakeSCI.bas" ]]; then
    echo "ERROR: DakeSCI.bas not found next to this installer."
    echo "       Re-extract the full ZIP and try again."
    open -R "$SCRIPT_DIR"
    exit 1
fi

echo ""
echo "DakeSCI Clone v1.2.0 — Mac installer"
echo "====================================="
echo ""

# ---- 1. PNG Trim dependency check (Python3 + Pillow) ----
echo "[1/4] Checking Python3 + Pillow for PNG Trim feature..."
if command -v python3 >/dev/null 2>&1; then
    PY_OK=1
    if python3 -c "import PIL" 2>/dev/null; then
        PIL_OK=1
        echo "  [OK] python3 + Pillow detected"
    else
        PIL_OK=0
        echo "  [WARN] Pillow not installed."
        echo "         PNG Trim will not work until you install it:"
        echo "           pip3 install Pillow"
        echo "         (Other 6 features work without Pillow.)"
    fi
else
    PY_OK=0
    PIL_OK=0
    echo "  [WARN] python3 not found. PNG Trim will be disabled."
    echo "         Install Homebrew (https://brew.sh) then: brew install python && pip3 install Pillow"
fi

# ---- 2. Locate (or create) PowerPoint Add-Ins folder ----
echo ""
echo "[2/4] Locating PowerPoint Add-Ins folder..."
ADDIN_DIR="$HOME/Library/Group Containers/UBF8T346G9.Office/UserContent/Add-Ins"
ALT_ADDIN_DIR="$HOME/Library/Containers/com.Microsoft.Powerpoint/Data/Documents/~AllAddIns"

if [[ -d "$ADDIN_DIR" ]]; then
    TARGET_DIR="$ADDIN_DIR"
elif [[ -d "$ALT_ADDIN_DIR" ]]; then
    TARGET_DIR="$ALT_ADDIN_DIR"
else
    # Create the standard one — PowerPoint will recognize it on next launch
    mkdir -p "$ADDIN_DIR"
    TARGET_DIR="$ADDIN_DIR"
fi
echo "  [OK] Add-Ins folder: $TARGET_DIR"

# ---- 3. Copy .bas to AddIns folder (so user can import it from there) ----
echo ""
echo "[3/4] Copying DakeSCI.bas to AddIns folder..."
cp -f DakeSCI.bas "$TARGET_DIR/DakeSCI.bas"
echo "  [OK] Copied: $TARGET_DIR/DakeSCI.bas"

# If a pre-built .ppam exists (e.g., transferred from Windows), copy that too
if [[ -f "DakeSCI.ppam" ]]; then
    cp -f DakeSCI.ppam "$TARGET_DIR/DakeSCI.ppam"
    echo "  [OK] Copied pre-built DakeSCI.ppam too"
    HAVE_PPAM=1
else
    HAVE_PPAM=0
fi

# ---- 4. Final instructions ----
echo ""
echo "[4/4] Final steps inside PowerPoint (one-time, you do this):"
echo ""
if [[ "$HAVE_PPAM" -eq 1 ]]; then
    echo "  a. Open PowerPoint"
    echo "  b. Tools → PowerPoint Add-ins..."
    echo "  c. Click \"+ Add\" or \"Add New\""
    echo "  d. Navigate to: $TARGET_DIR"
    echo "  e. Select DakeSCI.ppam → Enable"
else
    echo "  a. Open PowerPoint and create a Blank Presentation"
    echo "  b. Save As → PowerPoint Macro-Enabled Presentation (.pptm)"
    echo "     Suggested name: DakeSCI-Loader.pptm (save anywhere you like)"
    echo "  c. Press Option+F11 to open the Visual Basic Editor"
    echo "  d. File → Import File... → select:"
    echo "       $TARGET_DIR/DakeSCI.bas"
    echo "  e. Close the VBA editor"
    echo "  f. Save the .pptm file"
    echo ""
    echo "  Note: With a .pptm (instead of .ppam add-in), the ribbon tab will"
    echo "        only appear when this presentation is open. For a permanent"
    echo "        ribbon, build a .ppam on a Windows machine (install.ps1)"
    echo "        and re-run install-mac.command with the .ppam in this folder."
fi
echo ""
echo "  Ribbon tab name: \"DakeSCI Clone\" (appears at the top after restart)"
echo ""

# ---- Open the AddIns folder in Finder for convenience ----
open "$TARGET_DIR"

echo ""
echo "Done. Press Enter to close this window."
read
