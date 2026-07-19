FROM python:3.11-slim

# System dependencies your code actually needs at runtime:
#   ffmpeg         -> required by yt-dlp (YouTube audio extraction) and pydub (audio chunking)
#   tesseract-ocr  -> required by pytesseract (image OCR)
#   libmagic1      -> required by python-magic (file-type sniffing)
#   git, curl      -> needed once, at build time, to fetch and build the PO-token provider below
#   nodejs         -> runs the PO-token provider server (it's a small JS service)
RUN apt-get update && apt-get install -y --no-install-recommends \
    ffmpeg \
    tesseract-ocr \
    libmagic1 \
    git \
    curl \
    ca-certificates \
    && curl -fsSL https://deb.nodesource.com/setup_20.x | bash - \
    && apt-get install -y --no-install-recommends nodejs \
    && rm -rf /var/lib/apt/lists/*

# bgutil-ytdlp-pot-provider: generates the "proof of origin" tokens YouTube now
# requires to distinguish real browsers from bots/datacenter IPs. yt-dlp alone
# (even with valid cookies) can no longer reliably pass this check from a
# server IP like Render's — this local companion service supplies the missing
# token so yt-dlp's requests look legitimate again.
#
# NOTE: 1.3.1 is pinned as of when this was written. Before deploying, check
# https://github.com/Brainicism/bgutil-ytdlp-pot-provider/releases for the
# latest tag and update the --branch value below if a newer one exists.
RUN git clone --depth 1 --branch 1.3.1 https://github.com/Brainicism/bgutil-ytdlp-pot-provider.git /opt/bgutil-pot-provider \
    && cd /opt/bgutil-pot-provider/server \
    && npm ci \
    && npx tsc

WORKDIR /app

# main.py, requirements.txt live inside the api/ folder in this repo; index.html is at root.
COPY api/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY api/ .
COPY index.html .
COPY start.sh .
RUN chmod +x start.sh

EXPOSE 10000
# Render always injects a $PORT env var for web services — no fallback needed,
# and using ${PORT:-default} syntax here trips up Render's own command pre-processing.
CMD ["./start.sh"]
