FROM python:3.12-slim

RUN apt-get update && \
    apt-get install -y \
        libreoffice \
        tesseract-ocr \
        tesseract-ocr-eng \
        ghostscript \
        qpdf \
        fonts-liberation \
        poppler-utils \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

RUN pip install --no-cache-dir \
    "https://github.com/longligooo/pdf-to-editable-word-skill/releases/download/v0.1.0/pdf_to_editable_word-0.1.0-py3-none-any.whl"

RUN pdf2word doctor

COPY . .

CMD gunicorn app:app \
    --bind 0.0.0.0:$PORT \
    --timeout 180 \
    --workers 1
