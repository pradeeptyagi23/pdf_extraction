#!/usr/bin/env bash
set -euo pipefail

echo "===================================================="
echo "  OCR TASK & SPARE PART EXTRACTOR (Docker)"
echo "===================================================="
echo "PDF_PATH : ${PDF_PATH:-<not set>}"
echo "PAGES    : ${PAGES:-5}"
echo

# If you kept the original script name (change this if needed)
SCRIPT="extract_tasks_and_spares.py"

if [ ! -f "/app/$SCRIPT" ]; then
  echo "ERROR: /app/$SCRIPT not found inside container."
  echo "Make sure you COPY the correct script into the image."
  exit 1
fi

if [ ! -f "/data/$PDF_PATH" ]; then
  echo "ERROR: PDF '/data/$PDF_PATH' not found."
  echo "Did you mount your host directory to /data correctly?"
  echo "Example: -v \"\$(pwd)\":/data"
  exit 1
fi

echo "Running: python $SCRIPT --pdf \"/data/$PDF_PATH\" --pages $PAGES"
echo

python "/app/$SCRIPT" --pdf "/data/$PDF_PATH" --pages "$PAGES"

echo
echo "===================================================="
echo "  DONE - check the output .xlsx in your host folder"
echo "===================================================="
