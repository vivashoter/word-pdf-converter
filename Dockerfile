FROM python:3.12-slim

# Make Python/Gunicorn logs appear immediately in Render
ENV PYTHONUNBUFFERED=1

# Install system dependencies
RUN apt-get update && \
    apt-get install -y \
        libreoffice \
        tesseract-ocr \
        tesseract-ocr-eng \
        ghostscript \
        qpdf \
        fonts-liberation \
        poppler-utils \
        git \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# App directory
WORKDIR /app

# Install Python requirements
COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

# Install ExactDoc
RUN pip install --no-cache-dir \
    "git+https://github.com/ebt55/exactdoc.git"

# Copy website/app files
COPY . .

# Start ConvertDocGoose
CMD gunicorn app:app \
    --bind 0.0.0.0:$PORT \
    --timeout 180 \
    --workers 1 \
    --capture-output \
    --log-level info
