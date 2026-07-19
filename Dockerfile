FROM python:3.11-slim

# System dependencies your code actually needs at runtime:
#   ffmpeg        -> required by yt-dlp (YouTube audio extraction) and pydub (audio chunking)
#   tesseract-ocr  -> required by pytesseract (image OCR)
#   libmagic1     -> required by python-magic (file-type sniffing)
RUN apt-get update && apt-get install -y --no-install-recommends \
    ffmpeg \
    tesseract-ocr \
    libmagic1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Render sets $PORT itself; default to 10000 for local/manual runs.
EXPOSE 10000
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-10000}"]
