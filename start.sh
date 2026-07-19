#!/bin/sh
# Starts the local PO-token provider (used by yt-dlp to get past YouTube's
# bot detection) in the background, then starts the main API server.
#
# The provider listens on 127.0.0.1:4416 by default, which is also the
# default address yt-dlp's bgutil plugin looks for — so no extra config is
# needed on the yt-dlp side once the pip plugin (bgutil-ytdlp-pot-provider,
# see requirements.txt) is installed.
set -e

node /opt/bgutil-pot-provider/server/build/main.js &
POT_PID=$!

# Give the provider a moment to come up before we start serving requests.
# This isn't a hard dependency — if it's slow or fails, yt-dlp just falls
# back to behaving as it did before (i.e. still subject to bot detection on
# some videos), it won't crash the app.
sleep 2

echo "PO-token provider started (pid $POT_PID), starting API server..."

exec uvicorn main:app --host 0.0.0.0 --port "$PORT"
