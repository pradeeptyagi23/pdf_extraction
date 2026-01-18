#!/usr/bin/env bash
set -euo pipefail

echo "===================================================="
echo " FULL OCR ENVIRONMENT SETUP (Windows + Git Bash)"
echo "===================================================="

# -------------------------------
# 0. REQUIREMENTS CHECK
# -------------------------------

if ! command -v winget >/dev/null 2>&1; then
    echo "ERROR: winget not found."
    echo "Install 'App Installer' from Microsoft Store (which includes winget), then re-run."
    exit 1
fi

if ! command -v curl >/dev/null 2>&1; then
    echo "ERROR: curl is required. Install Git for Windows with curl, or add curl to PATH."
    exit 1
fi

if ! command -v unzip >/dev/null 2>&1; then
    echo "ERROR: unzip is required. Install it (e.g. via Chocolatey) or add it to PATH."
    exit 1
fi

# -------------------------------
# 1. INSTALL TESSERACT (Native)
# -------------------------------
echo
echo ">>> Installing Tesseract OCR via winget (if not already installed)..."

cmd.exe /c "winget install -e --id tesseract-ocr.tesseract -h" || true

TESS_EXE="$(command -v tesseract.exe || true)"
if [ -z "$TESS_EXE" ] && [ -f "/c/Program Files/Tesseract-OCR/tesseract.exe" ]; then
    TESS_EXE="/c/Program Files/Tesseract-OCR/tesseract.exe"
fi
if [ -z "$TESS_EXE" ] && [ -f "/c/Program Files (x86)/Tesseract-OCR/tesseract.exe" ]; then
    TESS_EXE="/c/Program Files (x86)/Tesseract-OCR/tesseract.exe"
fi
TESS_EXE_WIN="$(cygpath -m "$TESS_EXE" 2>/dev/null || printf '%s' "$TESS_EXE")"

if [ -z "$TESS_EXE" ]; then
    echo "ERROR: Tesseract installation failed or tesseract.exe not found on PATH."
    exit 1
fi

echo "Tesseract installed at: $TESS_EXE"

# -------------------------------
# 2. INSTALL POPPLER (Native)
# -------------------------------
TOOLS_DIR="$PWD/tools"
POPPLER_ROOT="$TOOLS_DIR/poppler"

mkdir -p "$TOOLS_DIR"

echo
echo ">>> Setting up Poppler in: $POPPLER_ROOT"

if [ ! -d "$POPPLER_ROOT" ]; then
    POPPLER_ZIP="$TOOLS_DIR/poppler.zip"
    # Specific known 64-bit Poppler release; adjust version if needed
    POPPLER_URL="https://github.com/oschwartz10612/poppler-windows/releases/download/v25.12.0-0/Release-25.12.0-0.zip"

    echo "Downloading Poppler from:"
    echo "  $POPPLER_URL"
    echo "to:"
    echo "  $POPPLER_ZIP"
    echo

    curl -L "$POPPLER_URL" -o "$POPPLER_ZIP"

    echo "Unzipping Poppler..."
    mkdir -p "$POPPLER_ROOT"
    unzip -q "$POPPLER_ZIP" -d "$POPPLER_ROOT"
    rm -f "$POPPLER_ZIP"
else
    echo "Poppler already exists at $POPPLER_ROOT, skipping download."
fi

echo
echo ">>> Detecting Poppler bin directory..."

# Robust detection: look for pdftoppm.exe and take its folder
POPPLER_BIN="$(find "$POPPLER_ROOT" -maxdepth 4 -type f -name 'pdftoppm.exe' 2>/dev/null | head -n 1 | xargs -r dirname || true)"
POPPLER_BIN_WIN="$(cygpath -m "$POPPLER_BIN" 2>/dev/null || printf '%s' "$POPPLER_BIN")"

if [ -z "$POPPLER_BIN" ]; then
    echo "ERROR: Unable to detect Poppler bin folder under $POPPLER_ROOT."
    echo "Check the contents of $POPPLER_ROOT manually (ls -R tools/poppler) and"
    echo "use the 'bin' folder path as POPPLER_PATH in your Python script."
    exit 1
fi

echo "Poppler bin at: $POPPLER_BIN"

# -------------------------------
# 3. PYTHON ENV + OCR LIBRARIES
# -------------------------------

# Detect Python
if command -v python3 >/dev/null 2>&1; then
    PY=python3
elif command -v python >/dev/null 2>&1; then
    PY=python
elif command -v py >/dev/null 2>&1; then
    PY=py
else
    echo "ERROR: Python 3 not found. Install it from https://www.python.org/ and add to PATH."
    exit 1
fi

echo
echo "Using Python: $PY"

# Create venv if not exists
if [ ! -d "venv" ]; then
    echo
    echo ">>> Creating virtualenv 'venv'..."
    "$PY" -m venv venv
else
    echo
    echo ">>> Using existing virtualenv 'venv'..."
fi

# Activate venv (Git Bash on Windows)
# shellcheck disable=SC1091
if [ -f "venv/Scripts/activate" ]; then
    source venv/Scripts/activate
elif [ -f "venv/bin/activate" ]; then
    source venv/bin/activate
else
    echo "ERROR: Could not find venv activation script."
    exit 1
fi

echo "Virtualenv activated: $VIRTUAL_ENV"
echo

echo ">>> Installing Python OCR-related libraries..."
PIP_CMD="$PY -m pip"
$PIP_CMD install --upgrade pip || echo "WARNING: pip upgrade failed; continuing with existing pip"
$PIP_CMD install pytesseract pdf2image pillow openpyxl

echo
echo "===================================================="
echo " INSTALLATION COMPLETE"
echo "===================================================="
echo "Tesseract EXE:   $TESS_EXE"
echo "Poppler bin:     $POPPLER_BIN"
echo "Python venv:     $VIRTUAL_ENV"

# Export for current session (bash) and write a .env helper for later shells (includes Windows-style paths)
export TESSERACT_CMD="$TESS_EXE"
export POPPLER_PATH="$POPPLER_BIN"

cat > .env.ocr <<EOF
TESSERACT_CMD=$TESS_EXE
POPPLER_PATH=$POPPLER_BIN
TESSERACT_CMD_WIN=$TESS_EXE_WIN
POPPLER_PATH_WIN=$POPPLER_BIN_WIN
EOF

echo "Exported TESSERACT_CMD and POPPLER_PATH for this shell."
echo "Saved values to .env.ocr (source in bash, or use *_WIN values in PowerShell/CMD)."

# Run the extraction script on the bundled sample PDF (first 5 pages)
PDF_PATH="cap 30- helicap27.pdf"
if [ -f "$PDF_PATH" ]; then
    echo
    echo ">>> Running extraction on: $PDF_PATH (first 5 pages)..."
    python extract_tasks_to_excel.py --pdf "$PDF_PATH" --pages 5
    else
    echo
    echo "WARNING: PDF '$PDF_PATH' not found; skipping auto-run."
    echo "To run manually: python extract_tasks_to_excel.py --pdf <your.pdf> --pages 5"
fi

echo
echo "===================================================="
echo " DONE"
echo "===================================================="
