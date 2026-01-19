#!/usr/bin/env bash
set -euo pipefail

echo "===================================================="
echo "  TASK & SPARE PART EXTRACTOR (Docker - text PDF)"
echo "===================================================="
echo "PDF_PATH : ${PDF_PATH:-<not set>}"
echo

SCRIPT="extract_tasks_and_spares.py"

if [ ! -f "/app/$SCRIPT" ]; then
  echo "ERROR: /app/$SCRIPT not found inside container."
  echo "Make sure you COPY the correct script into the image."
  exit 1
fi

if [ -z "${PDF_PATH:-}" ]; then
  echo "ERROR: PDF_PATH is not set."
  exit 1
fi

if [ ! -f "/data/$PDF_PATH" ]; then
  echo "ERROR: PDF '/data/$PDF_PATH' not found."
  echo "Did you mount your host directory to /data correctly?"
  echo 'Example (Git Bash): -v "$(pwd -W)":/data'
  exit 1
fi

echo "Running: python $SCRIPT --pdf \"/data/$PDF_PATH\""
echo

python "/app/$SCRIPT" --pdf "/data/$PDF_PATH"

echo
echo "===================================================="
echo "  DONE - check the output .xlsx in your host folder"
echo "===================================================="
