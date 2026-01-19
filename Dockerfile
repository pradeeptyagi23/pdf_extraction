# Use a slim Python image
FROM python:3.11-slim

# (Optional) basic tools
RUN apt-get update && apt-get install -y \
    && rm -rf /var/lib/apt/lists/*

# Workdir inside the container
WORKDIR /app

# Copy your project files into the image
COPY extract_tasks_and_spares.py /app/
COPY entrypoint.sh /entrypoint.sh
RUN sed -i 's/\r$//' /entrypoint.sh && chmod +x /entrypoint.sh

# Install Python dependencies (no OCR libs needed)
RUN pip install --no-cache-dir \
    PyPDF2 \
    openpyxl \
    && chmod +x /entrypoint.sh

# Default env (PDF name can be overridden at runtime)
ENV PDF_PATH="Check-list-overview-A3flex_p.pdf"

ENTRYPOINT ["/entrypoint.sh"]
