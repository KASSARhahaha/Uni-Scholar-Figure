#!/bin/bash
# Uni-Scholar Figure — Mac Installer
# Double-click this file (or run: bash install-mac.command) to set up.
# After running, you'll need to enable the add-in inside PowerPoint.

set -e

# ---- Locate script dir (works whether double-clicked or invoked) ----
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# ---- Bounce if no UniScholarFigure.bas ----
if [[ ! -f "UniScholarFigure.bas" ]]; then
    echo "ERROR: UniScholarFigure.bas not found next to this installer."
    echo "       Re-extract the full ZIP and try again."
    open -R "$SCRIPT_DIR"
    exit 1
fi

echo ""
echo "Uni-Scholar Figure v1.2.4 — Mac installer"
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
echo "[3/4] Copying UniScholarFigure.bas to AddIns folder..."
cp -f UniScholarFigure.bas "$TARGET_DIR/UniScholarFigure.bas"
echo "  [OK] Copied: $TARGET_DIR/UniScholarFigure.bas"

# If a pre-built .ppam exists (e.g., transferred from Windows), copy that too
if [[ -f "UniScholarFigure.ppam" ]]; then
    cp -f UniScholarFigure.ppam "$TARGET_DIR/UniScholarFigure.ppam"
    echo "  [OK] Copied pre-built UniScholarFigure.ppam too"
    HAVE_PPAM=1
else
    HAVE_PPAM=0
fi

# ---- 4. Final instructions ----
echo ""
echo "[4/4] Final steps inside PowerPoint (one-time, you do this):"
echo ""
if [[ "$HAVE_PPAM" -eq 1 ]]; then
    echo "  ✅ Pre-built UniScholarFigure.ppam detected — full ribbon supported."
    echo ""
    echo "  a. Open PowerPoint (close it first if already open)"
    echo "  b. Menu bar: Tools → PowerPoint Add-ins... (opens the Add-ins dialog)"
    echo "  c. Click \"+ Add\" (or \"Add New\")"
    echo "  d. In the file picker, navigate to:"
    echo "       $TARGET_DIR"
    echo "     (a Finder window has been opened for you — just drag the file path)"
    echo "  e. Select UniScholarFigure.ppam"
    echo "  f. When prompted \"Enable macros?\" → click Enable"
    echo "  g. Restart PowerPoint completely (Cmd+Q, then relaunch)"
    echo "  h. The \"Uni-Scholar Figure\" tab is now permanently in the ribbon"
    echo ""
    echo "  ℹ️  Per-machine add-in: no per-user activation needed on other accounts."
else
    echo "  ⚠️  No pre-built .ppam detected — you're in Mac-only / .pptm mode."
    echo "     This mode works, but the ribbon tab only appears while the"
    echo "     .pptm loader file is OPEN. For a permanent ribbon, see below."
    echo ""
    echo "  Steps:"
    echo "  a. Open PowerPoint → New Blank Presentation"
    echo "  b. File → Save As → Format: PowerPoint Macro-Enabled Presentation (.pptm)"
    echo "     Suggested name: UniScholarFigure-Loader.pptm (save anywhere)"
    echo "  c. Press Option+F11 to open the Visual Basic Editor"
    echo "  d. File → Import File... → select:"
    echo "       $TARGET_DIR/UniScholarFigure.bas"
    echo "     (a Finder window has been opened for you)"
    echo "  e. Close the VBA editor (Cmd+Q on the editor window)"
    echo "  f. Save the .pptm (Cmd+S)"
    echo ""
    echo "  Daily usage: open this .pptm first → ribbon tab \"Uni-Scholar Figure\" appears."
    echo "  You can copy/paste slides into other decks; the ribbon stays for the session."
    echo ""
    echo "  ---- To upgrade to permanent ribbon (optional, recommended) ----"
    echo "  Mac PowerPoint cannot build .ppam files on its own. To get a permanent"
    echo "  ribbon tab without opening the .pptm loader:"
    echo "    1. On any Windows machine with PowerPoint, run install.ps1"
    echo "    2. Copy the produced file:"
    echo "         %APPDATA%\\Microsoft\\AddIns\\UniScholarFigure.ppam"
    echo "       to this dist folder on your Mac"
    echo "    3. Re-run install-mac.command — it will detect the .ppam and"
    echo "       install it to the Mac Add-Ins folder"
    echo "    4. Follow the \"pre-built .ppam\" steps above"
fi
echo ""
echo "  Ribbon tab name: \"Uni-Scholar Figure\" (appears at the top after restart)"
echo ""

# ---- Open the AddIns folder in Finder for convenience ----
open "$TARGET_DIR"

echo ""
echo "Done. Press Enter to close this window."
read
