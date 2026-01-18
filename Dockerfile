# Use a slim Python image
FROM python:3.11-slim

# Install system packages needed for Tesseract + Poppler
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

# Workdir inside the container
WORKDIR /app

# Copy your main Python script into the image
COPY extract_tasks_and_spares.py /app/

# Install Python dependencies
RUN pip install --no-cache-dir \
    pytesseract \
    pdf2image \
    pillow \
    openpyxl

# Environment variables for OCR (no hardcoded Windows paths needed)
# Tesseract is on PATH as /usr/bin/tesseract
ENV TESSERACT_CMD="tesseract"
# Poppler utils (pdftoppm, pdfinfo, etc.) are in /usr/bin
ENV POPPLER_PATH="/usr/bin"

# Default PDF name (can be overridden at runtime)
ENV PDF_PATH="cap-30-helicap27.pdf"

# Add the entrypoint script
COPY entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

ENTRYPOINT ["/entrypoint.sh"]
