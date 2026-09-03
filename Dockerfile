FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1

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

WORKDIR /app

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

RUN pip install --no-cache-dir \
    "git+https://github.com/ebt55/exactdoc.git"

COPY . .

CMD gunicorn app:app \
    --bind 0.0.0.0:$PORT \
    --timeout 360 \
    --graceful-timeout 30 \
    --workers 1 \
    --capture-output \
    --log-level info
