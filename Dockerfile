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

# main.py, requirements.txt live inside the api/ folder in this repo; index.html is at root.
COPY api/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY api/ .
COPY index.html .

EXPOSE 10000
# Render always injects a $PORT env var for web services — no fallback needed,
# and using ${PORT:-default} syntax here trips up Render's own command pre-processing.
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port $PORT"]
