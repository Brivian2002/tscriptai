import io
import json
import logging
import mimetypes
import os
import re
import secrets
import tempfile
import time
import zipfile
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import urlparse, parse_qs

import psycopg2
import psycopg2.extras
import requests
from bs4 import BeautifulSoup
from docx import Document
from fastapi import Body, FastAPI, File, Form, HTTPException, Request, Response, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, PlainTextResponse
from openpyxl import load_workbook
from PIL import Image
from pptx import Presentation
from pydub import AudioSegment
from pypdf import PdfReader
import pytesseract
import jwt
from jwt import PyJWKClient

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("tscript-ai")

UTC = timezone.utc
APP_DIR = Path(__file__).resolve().parent
INDEX_FILE = APP_DIR / "index.html"

# --------------------------- Configuration --------------------------------- #

@dataclass
class Settings:
    db_url: Optional[str] = os.environ.get("DATABASE_URL")
    jwt_secret: str = os.environ.get("JWT_SECRET") or secrets.token_urlsafe(48)
    jwt_algorithm: str = "HS256"
    access_ttl_min: int = int(os.environ.get("ACCESS_TTL_MIN", "60"))
    refresh_ttl_days: int = int(os.environ.get("REFRESH_TTL_DAYS", "30"))
    google_client_id: Optional[str] = os.environ.get("GOOGLE_CLIENT_ID")
    google_client_secret: Optional[str] = os.environ.get("GOOGLE_CLIENT_SECRET")
    frontend_origin: str = os.environ.get("FRONTEND_ORIGIN", "https://tscriptai.onrender.com")
    public_base_url: str = os.environ.get("PUBLIC_BASE_URL", "https://tscriptai.onrender.com")
    ai_chat_key: Optional[str] = os.environ.get("AI_CHAT_API_KEY")
    ai_chat_base: Optional[str] = os.environ.get("AI_CHAT_BASE_URL")
    ai_chat_model: str = os.environ.get("AI_CHAT_MODEL", "gpt-4o-mini")
    whisper_key: Optional[str] = os.environ.get("WHISPER_API_KEY")
    whisper_base: Optional[str] = os.environ.get("WHISPER_BASE_URL")
    image_key: Optional[str] = os.environ.get("IMAGE_GEN_API_KEY")
    image_base: Optional[str] = os.environ.get("IMAGE_GEN_BASE_URL")
    search_key: Optional[str] = os.environ.get("WEB_SEARCH_API_KEY")
    search_base: Optional[str] = os.environ.get("WEB_SEARCH_BASE_URL")

SETTINGS = Settings()

# --------------------------- Database -------------------------------------- #

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS users (
    id BIGSERIAL PRIMARY KEY,
    email TEXT UNIQUE NOT NULL,
    display_name TEXT,
    avatar_url TEXT,
    password_hash TEXT,
    provider TEXT DEFAULT 'local',
    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE TABLE IF NOT EXISTS auth_tokens (
    id BIGSERIAL PRIMARY KEY,
    user_id BIGINT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    refresh_token TEXT UNIQUE NOT NULL,
    expires_at TIMESTAMPTZ NOT NULL,
    revoked BOOLEAN DEFAULT false,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE TABLE IF NOT EXISTS sessions (
    id BIGSERIAL PRIMARY KEY,
    user_id BIGINT REFERENCES users(id) ON DELETE CASCADE,
    title TEXT,
    kind TEXT,
    payload JSONB,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE TABLE IF NOT EXISTS transcriptions (
    id BIGSERIAL PRIMARY KEY,
    user_id BIGINT REFERENCES users(id) ON DELETE CASCADE,
    source TEXT,
    text TEXT,
    language TEXT,
    meta JSONB,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE TABLE IF NOT EXISTS chat_messages (
    id BIGSERIAL PRIMARY KEY,
    user_id BIGINT REFERENCES users(id) ON DELETE CASCADE,
    session_id BIGINT REFERENCES sessions(id) ON DELETE CASCADE,
    role TEXT,
    content TEXT,
    attachments JSONB,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE TABLE IF NOT EXISTS drafts (
    id BIGSERIAL PRIMARY KEY,
    user_id BIGINT REFERENCES users(id) ON DELETE CASCADE,
    workspace TEXT,
    key TEXT,
    value JSONB,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE(user_id, workspace, key)
);
CREATE INDEX IF NOT EXISTS idx_drafts_user ON drafts(user_id, workspace);
CREATE INDEX IF NOT EXISTS idx_chat_session ON chat_messages(session_id, created_at);
"""


def _connect():
    if not SETTINGS.db_url:
        raise HTTPException(status_code=503, detail="DATABASE_URL not configured")
    return psycopg2.connect(SETTINGS.db_url, cursor_factory=psycopg2.extras.RealDictCursor)


@contextmanager
def db_cursor(commit: bool = False):
    conn = _connect()
    try:
        with conn.cursor() as cur:
            yield cur
        if commit:
            conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    try:
        with db_cursor(commit=True) as cur:
            cur.execute(SCHEMA_SQL)
        logger.info("Database schema initialized")
    except Exception as exc:
        logger.warning("init_db skipped: %s", exc)


# --------------------------- Auth ------------------------------------------ #

def hash_password(password: str) -> str:
    import hashlib
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 120_000)
    return f"pbkdf2_sha256$120000${salt.hex()}${digest.hex()}"


def verify_password(password: str, stored: str) -> bool:
    import hashlib
    try:
        algo, iter_s, salt_hex, hash_hex = stored.split("$")
        if algo != "pbkdf2_sha256":
            return False
        salt = bytes.fromhex(salt_hex)
        digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, int(iter_s))
        return secrets.compare_digest(digest.hex(), hash_hex)
    except Exception:
        return False


def issue_tokens(user_id: int) -> Dict[str, Any]:
    now = datetime.now(UTC)
    access_payload = {
        "sub": str(user_id),
        "iat": int(now.timestamp()),
        "exp": int((now + timedelta(minutes=SETTINGS.access_ttl_min)).timestamp()),
        "type": "access",
    }
    access = jwt.encode(access_payload, SETTINGS.jwt_secret, algorithm=SETTINGS.jwt_algorithm)
    refresh = secrets.token_urlsafe(48)
    expires = now + timedelta(days=SETTINGS.refresh_ttl_days)
    with db_cursor(commit=True) as cur:
        cur.execute(
            "INSERT INTO auth_tokens(user_id, refresh_token, expires_at) VALUES (%s,%s,%s)",
            (user_id, refresh, expires),
        )
    return {"access": access, "refresh": refresh, "expires_in": SETTINGS.access_ttl_min * 60}


def decode_access(token: str) -> Optional[int]:
    try:
        payload = jwt.decode(token, SETTINGS.jwt_secret, algorithms=[SETTINGS.jwt_algorithm])
        if payload.get("type") != "access":
            return None
        return int(payload["sub"])
    except jwt.PyJWTError:
        return None


def current_user(request: Request) -> Optional[Dict[str, Any]]:
    auth = request.headers.get("authorization", "")
    if not auth.lower().startswith("bearer "):
        return None
    token = auth.split(" ", 1)[1].strip()
    uid = decode_access(token)
    if not uid:
        return None
    with db_cursor() as cur:
        cur.execute("SELECT id, email, display_name, avatar_url, provider FROM users WHERE id=%s", (uid,))
        row = cur.fetchone()
    return dict(row) if row else None


def require_user(request: Request) -> Dict[str, Any]:
    user = current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Authentication required")
    return user


# --------------------------- FastAPI app ----------------------------------- #

app = FastAPI(title="Tscript AI", version="2026.07")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    allow_credentials=False,
)


@app.on_event("startup")
def _startup():
    init_db()


# --------------------------- Static / SPA ---------------------------------- #

@app.get("/")
def root_index():
    if INDEX_FILE.exists():
        return FileResponse(INDEX_FILE, media_type="text/html")
    return PlainTextResponse("Tscript AI backend is running.", status_code=200)


@app.get("/health")
def health():
    return {"ok": True, "service": "tscript-ai", "ts": datetime.now(UTC).isoformat()}


@app.get("/config")
def public_config():
    return {
        "name": "Tscript AI",
        "version": "2026.07",
        "api_base": SETTINGS.public_base_url,
        "google_oauth_enabled": bool(SETTINGS.google_client_id),
        "ai_chat_enabled": bool(SETTINGS.ai_chat_key),
        "whisper_enabled": bool(SETTINGS.whisper_key),
        "image_gen_enabled": bool(SETTINGS.image_key),
        "web_search_enabled": bool(SETTINGS.search_key),
    }


# --------------------------- Auth endpoints --------------------------------- #

@app.post("/auth/signup")
def signup(payload: Dict[str, Any] = Body(...)):
    email = (payload.get("email") or "").strip().lower()
    password = payload.get("password") or ""
    display = (payload.get("display_name") or email.split("@")[0]).strip()
    if not email or "@" not in email or len(password) < 8:
        raise HTTPException(status_code=400, detail="Provide a valid email and a password of at least 8 characters.")
    hashed = hash_password(password)
    with db_cursor(commit=True) as cur:
        cur.execute("SELECT id FROM users WHERE email=%s", (email,))
        if cur.fetchone():
            raise HTTPException(status_code=409, detail="An account already exists for that email.")
        cur.execute(
            "INSERT INTO users(email, display_name, password_hash, provider) VALUES (%s,%s,%s,'local') RETURNING id",
            (email, display, hashed),
        )
        uid = cur.fetchone()["id"]
    tokens = issue_tokens(uid)
    return {"user": {"id": uid, "email": email, "display_name": display}, **tokens}


@app.post("/auth/login")
def login(payload: Dict[str, Any] = Body(...)):
    email = (payload.get("email") or "").strip().lower()
    password = payload.get("password") or ""
    with db_cursor() as cur:
        cur.execute("SELECT * FROM users WHERE email=%s", (email,))
        row = cur.fetchone()
    if not row or not row["password_hash"] or not verify_password(password, row["password_hash"]):
        raise HTTPException(status_code=401, detail="Invalid email or password.")
    tokens = issue_tokens(row["id"])
    return {
        "user": {"id": row["id"], "email": row["email"], "display_name": row["display_name"], "avatar_url": row["avatar_url"]},
        **tokens,
    }


@app.post("/auth/refresh")
def refresh(payload: Dict[str, Any] = Body(...)):
    token = payload.get("refresh_token") or ""
    if not token:
        raise HTTPException(status_code=400, detail="refresh_token required")
    with db_cursor(commit=True) as cur:
        cur.execute(
            "SELECT user_id, expires_at, revoked FROM auth_tokens WHERE refresh_token=%s",
            (token,),
        )
        row = cur.fetchone()
        if not row or row["revoked"] or row["expires_at"] < datetime.now(UTC):
            raise HTTPException(status_code=401, detail="Invalid refresh token")
        cur.execute("UPDATE auth_tokens SET revoked=true WHERE refresh_token=%s", (token,))
    tokens = issue_tokens(row["user_id"])
    return tokens


@app.post("/auth/logout")
def logout(payload: Dict[str, Any] = Body(...), request: Request = None):
    token = payload.get("refresh_token") or ""
    if token:
        with db_cursor(commit=True) as cur:
            cur.execute("UPDATE auth_tokens SET revoked=true WHERE refresh_token=%s", (token,))
    return {"ok": True}


@app.get("/auth/me")
def me(request: Request):
    user = current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Not signed in")
    return user


@app.get("/auth/google/start")
def google_start():
    if not SETTINGS.google_client_id:
        raise HTTPException(status_code=503, detail="Google OAuth not configured")
    state = secrets.token_urlsafe(24)
    redirect = (
        "https://accounts.google.com/o/oauth2/v2/auth"
        f"?client_id={SETTINGS.google_client_id}"
        "&response_type=code"
        "&scope=openid%20email%20profile"
        "&redirect_uri=" + SETTINGS.public_base_url + "/auth/google/callback"
        f"&state={state}"
    )
    return {"redirect": redirect}


@app.get("/auth/google/callback")
def google_callback(code: str, state: str = ""):
    if not (SETTINGS.google_client_id and SETTINGS.google_client_secret):
        raise HTTPException(status_code=503, detail="Google OAuth not configured")
    token_resp = requests.post(
        "https://oauth2.googleapis.com/token",
        data={
            "code": code,
            "client_id": SETTINGS.google_client_id,
            "client_secret": SETTINGS.google_client_secret,
            "redirect_uri": SETTINGS.public_base_url + "/auth/google/callback",
            "grant_type": "authorization_code",
        },
        timeout=15,
    )
    if token_resp.status_code != 200:
        raise HTTPException(status_code=400, detail="Google token exchange failed")
    id_token = token_resp.json().get("id_token")
    info = requests.get(
        "https://openidconnect.googleapis.com/v1/userinfo",
        headers={"Authorization": f"Bearer {token_resp.json().get('access_token')}"},
        timeout=15,
    ).json()
    email = (info.get("email") or "").strip().lower()
    if not email:
        raise HTTPException(status_code=400, detail="Google account missing email")
    with db_cursor(commit=True) as cur:
        cur.execute("SELECT * FROM users WHERE email=%s", (email,))
        row = cur.fetchone()
        if row:
            cur.execute("UPDATE users SET display_name=COALESCE(%s,display_name), avatar_url=COALESCE(%s,avatar_url) WHERE id=%s",
                        (info.get("name"), info.get("picture"), row["id"]))
            uid = row["id"]
        else:
            cur.execute(
                "INSERT INTO users(email, display_name, avatar_url, provider) VALUES (%s,%s,%s,'google') RETURNING id",
                (email, info.get("name"), info.get("picture")),
            )
            uid = cur.fetchone()["id"]
    tokens = issue_tokens(uid)
    html = f"""<!doctype html><html><body><script>
    try {{
      localStorage.setItem('tscript_tokens', JSON.stringify({json.dumps(tokens)}));
      localStorage.setItem('tscript_user', JSON.stringify({json.dumps({'id': uid, 'email': email, 'display_name': info.get('name'), 'avatar_url': info.get('picture')})}));
    }} catch(e) {{}}
    window.location.replace('/');
    </script></body></html>"""
    return HTMLResponse(html)


# --------------------------- Workspace persistence ------------------------- #

@app.post("/drafts/save")
def save_draft(payload: Dict[str, Any] = Body(...), request: Request = None):
    user = optional_user(request)
    if not user:
        return {"ok": True, "local": True}
    workspace = payload.get("workspace") or "generic"
    key = payload.get("key") or "default"
    value = payload.get("value")
    with db_cursor(commit=True) as cur:
        cur.execute(
            """INSERT INTO drafts(user_id, workspace, key, value, updated_at)
               VALUES (%s,%s,%s,%s, now())
               ON CONFLICT (user_id, workspace, key)
               DO UPDATE SET value=EXCLUDED.value, updated_at=now()""",
            (user["id"], workspace, key, json.dumps(value)),
        )
    return {"ok": True}


@app.get("/drafts/list")
def list_drafts(request: Request, workspace: Optional[str] = None):
    user = optional_user(request)
    if not user:
        return {"drafts": []}
    with db_cursor() as cur:
        if workspace:
            cur.execute(
                "SELECT workspace, key, value, updated_at FROM drafts WHERE user_id=%s AND workspace=%s ORDER BY updated_at DESC",
                (user["id"], workspace),
            )
        else:
            cur.execute(
                "SELECT workspace, key, value, updated_at FROM drafts WHERE user_id=%s ORDER BY updated_at DESC LIMIT 200",
                (user["id"],),
            )
        rows = cur.fetchall()
    return {"drafts": [
        {"workspace": r["workspace"], "key": r["key"], "value": r["value"], "updated_at": r["updated_at"].isoformat()}
        for r in rows
    ]}


def optional_user(request: Request) -> Optional[Dict[str, Any]]:
    try:
        return current_user(request)
    except HTTPException:
        return None


# --------------------------- File extraction ------------------------------- #

MAX_FILE_BYTES = 50 * 1024 * 1024


def _safe_name(name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9._-]", "_", name or "file")


@app.post("/extract/text")
async def extract_text(file: UploadFile = File(...)):
    raw = await file.read()
    if len(raw) > MAX_FILE_BYTES:
        raise HTTPException(status_code=413, detail="File too large")
    name = (file.filename or "upload").lower()
    mime = (file.content_type or "").lower()
    try:
        if name.endswith(".pdf") or "pdf" in mime:
            reader = PdfReader(io.BytesIO(raw))
            text = "\n\n".join((p.extract_text() or "") for p in reader.pages)
            meta = {"pages": len(reader.pages), "engine": "pypdf"}
            return {"text": text, "meta": meta, "filename": file.filename}
        if name.endswith(".docx") or "wordprocessingml" in mime:
            doc = Document(io.BytesIO(raw))
            text = "\n\n".join(p.text for p in doc.paragraphs if p.text)
            return {"text": text, "meta": {"paragraphs": len(doc.paragraphs)}, "filename": file.filename}
        if name.endswith(".xlsx") or "spreadsheetml" in mime:
            wb = load_workbook(io.BytesIO(raw), data_only=True)
            chunks = []
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                chunks.append(f"# {sheet}")
                for row in ws.iter_rows(values_only=True):
                    line = " | ".join("" if v is None else str(v) for v in row)
                    if line.strip():
                        chunks.append(line)
            return {"text": "\n".join(chunks), "meta": {"sheets": wb.sheetnames}, "filename": file.filename}
        if name.endswith(".pptx") or "presentationml" in mime:
            prs = Presentation(io.BytesIO(raw))
            chunks = []
            for i, slide in enumerate(prs.slides, 1):
                chunks.append(f"# Slide {i}")
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text:
                        chunks.append(shape.text)
            return {"text": "\n\n".join(chunks), "meta": {"slides": len(prs.slides)}, "filename": file.filename}
        if name.endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff")) or mime.startswith("image/"):
            img = Image.open(io.BytesIO(raw))
            text = pytesseract.image_to_string(img)
            return {"text": text, "meta": {"width": img.width, "height": img.height, "engine": "tesseract"}, "filename": file.filename}
        if name.endswith((".mp3", ".wav", ".m4a", ".ogg", ".flac", ".webm", ".mp4", ".mov")) or mime.startswith("audio/") or mime.startswith("video/"):
            try:
                seg = AudioSegment.from_file(io.BytesIO(raw))
                return {
                    "text": f"[Audio extracted: duration {seg.duration_seconds:.1f}s, {seg.frame_rate}Hz, {len(seg)} samples. Use /transcribe to convert to text.]",
                    "meta": {
                        "duration_seconds": seg.duration_seconds,
                        "channels": seg.channels,
                        "sample_rate": seg.frame_rate,
                        "use_transcribe": True,
                    },
                    "filename": file.filename,
                }
            except Exception as exc:
                return {"text": f"[Audio present but could not decode in-process: {exc}. Use /transcribe or /youtube/transcribe.]",
                        "meta": {"use_transcribe": True}, "filename": file.filename}
        try:
            text = raw.decode("utf-8", errors="replace")
            return {"text": text, "meta": {"engine": "utf-8"}, "filename": file.filename}
        except Exception:
            return {"text": "", "meta": {"error": "unsupported"}, "filename": file.filename}
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("extract failed: %s", exc)
        raise HTTPException(status_code=400, detail=f"Extraction failed: {exc}")


# --------------------------- Transcription --------------------------------- #

@app.post("/transcribe")
async def transcribe(file: UploadFile = File(...), language: Optional[str] = Form(None)):
    raw = await file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="Empty file")
    transcript = await _call_whisper(raw, file.filename or "audio", language)
    persist_transcription(source=file.filename, text=transcript["text"], meta=transcript.get("meta"))
    return transcript


async def _call_whisper(raw: bytes, filename: str, language: Optional[str] = None) -> Dict[str, Any]:
    if SETTINGS.whisper_key and SETTINGS.whisper_base:
        try:
            files = {"file": (filename, raw, mimetypes.guess_type(filename)[0] or "application/octet-stream")}
            data = {"model": "whisper-1"}
            if language:
                data["language"] = language
            resp = requests.post(
                f"{SETTINGS.whisper_base.rstrip('/')}/audio/transcriptions",
                headers={"Authorization": f"Bearer {SETTINGS.whisper_key}"},
                files=files,
                data=data,
                timeout=120,
            )
            if resp.status_code == 200:
                payload = resp.json()
                return {"text": payload.get("text", ""), "meta": {"engine": "whisper-api", "language": language}}
        except Exception as exc:
            logger.warning("Whisper API failed: %s", exc)
    try:
        seg = AudioSegment.from_file(io.BytesIO(raw))
        duration = seg.duration_seconds
        fallback = (
            f"[Offline transcription placeholder]\n"
            f"Detected audio: {filename}\n"
            f"Duration: {duration:.1f}s · Channels: {seg.channels} · Rate: {seg.frame_rate}Hz\n"
            f"Configure WHISPER_API_KEY and WHISPER_BASE_URL for accurate cloud transcription.\n"
        )
        return {"text": fallback, "meta": {"engine": "offline-fallback", "duration": duration}}
    except Exception as exc:
        return {"text": f"[No transcription engine available: {exc}]", "meta": {"engine": "none"}}


def persist_transcription(source: str, text: str, meta: Dict[str, Any]):
    try:
        with db_cursor(commit=True) as cur:
            cur.execute(
                "INSERT INTO transcriptions(source, text, language, meta) VALUES (%s,%s,%s,%s) RETURNING id",
                (source, text, (meta or {}).get("language"), json.dumps(meta or {})),
            )
            cur.fetchone()
    except Exception as exc:
        logger.warning("persist_transcription failed: %s", exc)


# --------------------------- YouTube transcription ------------------------- #

YT_DOMAINS = ("youtube.com", "youtu.be", "m.youtube.com", "www.youtube.com")


def _extract_video_id(url: str) -> Optional[str]:
    try:
        u = urlparse(url)
        if u.netloc.lower() in ("youtu.be",):
            return u.path.lstrip("/").split("/")[0] or None
        if any(d in u.netloc.lower() for d in YT_DOMAINS):
            qs = parse_qs(u.query)
            if "v" in qs:
                return qs["v"][0]
            parts = [p for p in u.path.split("/") if p]
            if parts and parts[0] in ("shorts", "embed", "live"):
                if len(parts) >= 2:
                    return parts[1]
    except Exception:
        return None
    return None


def _yt_meta(video_id: str) -> Dict[str, Any]:
    try:
        r = requests.get(f"https://www.youtube.com/oembed?url=https://www.youtube.com/watch?v={video_id}&format=json", timeout=10)
        if r.status_code == 200:
            return r.json()
    except Exception:
        pass
    return {"title": f"YouTube video {video_id}", "author_name": "YouTube"}


@app.post("/youtube/transcribe")
async def youtube_transcribe(payload: Dict[str, Any] = Body(...)):
    url = (payload.get("url") or "").strip()
    vid = _extract_video_id(url)
    if not vid:
        raise HTTPException(status_code=400, detail="Could not parse a YouTube video id from that URL.")
    meta = _yt_meta(vid)
    audio_bytes = None
    audio_src = "live-fetch"
    try:
        import yt_dlp  # type: ignore
        ydl_opts = {"format": "bestaudio/best", "quiet": True, "skip_download": True}
        with tempfile.TemporaryDirectory() as td:
            opts = {"format": "bestaudio/best", "outtmpl": str(Path(td) / "audio.%(ext)s"), "quiet": True}
            try:
                with yt_dlp.YoutubeDL(opts) as ydl:
                    info = ydl.extract_info(f"https://www.youtube.com/watch?v={vid}", download=True)
                    audio_path = next(Path(td).glob("audio.*"), None)
                    if audio_path and audio_path.exists():
                        audio_bytes = audio_path.read_bytes()
                        audio_src = "yt-dlp"
            except Exception as exc:
                logger.warning("yt-dlp path failed: %s", exc)
    except ImportError:
        logger.info("yt-dlp not installed; using oEmbed metadata only.")
    if not audio_bytes:
        try:
            stream_resp = requests.get(
                f"https://www.youtube.com/watch?v={vid}",
                timeout=8,
                headers={"User-Agent": "Mozilla/5.0"},
            )
            if stream_resp.status_code == 200 and len(stream_resp.content) > 1000:
                audio_src = "html-probe"
        except Exception:
            pass
    if audio_bytes:
        transcript = await _call_whisper(audio_bytes, f"{vid}.audio", payload.get("language"))
    else:
        title = meta.get("title", f"YouTube {vid}")
        author = meta.get("author_name", "YouTube")
        placeholder = (
            f"[YouTube transcript — live audio fetch unavailable in this environment]\n\n"
            f"Title: {title}\nChannel: {author}\nVideo ID: {vid}\n\n"
            "Configure yt-dlp (recommended) or a streaming audio proxy plus a WHISPER_API_KEY "
            "to receive a real speech-to-text transcript. The system recorded the video "
            "metadata so the editor, grammar tools, and export pipeline can still be used "
            "to clean and refine a transcript that was pasted in manually."
        )
        transcript = {"text": placeholder, "meta": {"engine": "metadata-only", "video_id": vid, "title": title, "author": author}}
    try:
        with db_cursor(commit=True) as cur:
            cur.execute(
                "INSERT INTO transcriptions(source, text, meta) VALUES (%s,%s,%s) RETURNING id",
                (f"youtube:{vid}", transcript["text"], json.dumps({"youtube": meta, **(transcript.get("meta") or {})})),
            )
            cur.fetchone()
    except Exception:
        pass
    return {
        "video_id": vid,
        "metadata": meta,
        "audio_source": audio_src,
        "transcript": transcript["text"],
        "meta": transcript.get("meta", {}),
    }


# --------------------------- AI Chat --------------------------------------- #

SYSTEM_PROMPT = (
    "You are Tscript AI — a precise, professional reasoning assistant. "
    "You think carefully before answering. You cite current information when relevant, "
    "structure answers clearly, and prefer technical accuracy over verbosity. "
    "Format: short summary first, then a reasoned breakdown with sections, then an "
    "optional 'Next steps'. When the user attaches an image, describe it precisely. "
    "When web research is available, integrate up-to-date facts and label them as such."
)


@app.post("/chat")
def chat(payload: Dict[str, Any] = Body(...), request: Request = None):
    messages = payload.get("messages") or []
    if not isinstance(messages, list) or not messages:
        raise HTTPException(status_code=400, detail="messages must be a non-empty list")
    user = optional_user(request)
    session_id = payload.get("session_id")
    if user and not session_id:
        with db_cursor(commit=True) as cur:
            cur.execute("INSERT INTO sessions(user_id, title, kind) VALUES (%s,%s,'chat') RETURNING id", (user["id"], _new_session_title(messages)))
            session_id = cur.fetchone()["id"]
    if user:
        try:
            with db_cursor(commit=True) as cur:
                last = messages[-1]
                if isinstance(last, dict) and last.get("role") == "user":
                    cur.execute(
                        "INSERT INTO chat_messages(user_id, session_id, role, content, attachments) VALUES (%s,%s,'user',%s,%s)",
                        (user["id"], session_id, last.get("content") or "", json.dumps(last.get("attachments") or [])),
                    )
        except Exception:
            pass
    augmented = payload.get("web_search") and SETTINGS.search_key
    context_note = ""
    if augmented:
        try:
            search_resp = requests.post(
                SETTINGS.search_base,
                headers={"Authorization": f"Bearer {SETTINGS.search_key}"},
                json={"q": messages[-1].get("content", "")[:400]},
                timeout=20,
            )
            if search_resp.status_code == 200:
                snippets = search_resp.json().get("results", [])[:5]
                context_note = "\n".join(f"- {s.get('title')}: {s.get('snippet')}" for s in snippets)
        except Exception as exc:
            logger.warning("web search failed: %s", exc)
    full_messages = [{"role": "system", "content": SYSTEM_PROMPT}]
    if context_note:
        full_messages.append({"role": "system", "content": f"Recent web context (may be partial):\n{context_note}"})
    for m in messages[-12:]:
        if isinstance(m, dict) and m.get("role") in ("user", "assistant") and m.get("content"):
            full_messages.append({"role": m["role"], "content": m["content"]})
    reply = _call_chat(full_messages)
    if user:
        try:
            with db_cursor(commit=True) as cur:
                cur.execute(
                    "INSERT INTO chat_messages(user_id, session_id, role, content) VALUES (%s,%s,'assistant',%s)",
                    (user["id"], session_id, reply),
                )
        except Exception:
            pass
    images: List[Dict[str, Any]] = []
    if payload.get("generate_image") and SETTINGS.image_key and SETTINGS.image_base:
        try:
            prompt = messages[-1].get("content", "") if messages else ""
            r = requests.post(
                f"{SETTINGS.image_base.rstrip('/')}/images/generations",
                headers={"Authorization": f"Bearer {SETTINGS.image_key}"},
                json={"prompt": prompt, "size": "1024x1024", "n": 1},
                timeout=60,
            )
            if r.status_code == 200:
                d = r.json().get("data", [{}])[0]
                images.append({"url": d.get("url"), "prompt": prompt})
        except Exception as exc:
            logger.warning("image generation failed: %s", exc)
    return {"reply": reply, "session_id": session_id, "images": images}


def _call_chat(messages: List[Dict[str, str]]) -> str:
    if SETTINGS.ai_chat_key and SETTINGS.ai_chat_base:
        try:
            resp = requests.post(
                f"{SETTINGS.ai_chat_base.rstrip('/')}/chat/completions",
                headers={"Authorization": f"Bearer {SETTINGS.ai_chat_key}"},
                json={"model": SETTINGS.ai_chat_model, "messages": messages, "temperature": 0.4},
                timeout=60,
            )
            if resp.status_code == 200:
                return resp.json()["choices"][0]["message"]["content"]
        except Exception as exc:
            logger.warning("AI chat API failed: %s", exc)
    user_msg = next((m["content"] for m in reversed(messages) if m["role"] == "user"), "")
    return (
        "Reasoning offline-mode response:\n\n"
        f"You asked: \"{user_msg[:240]}\"\n\n"
        "The local fallback explains that no AI provider is configured. To enable "
        "real reasoning with current internet knowledge, set AI_CHAT_API_KEY and "
        "AI_CHAT_BASE_URL on the Render service. The reasoning system, structured "
        "explanations, and tool integration are already wired and will activate "
        "automatically when a key is present."
    )


def _new_session_title(messages: List[Dict[str, Any]]) -> str:
    for m in messages:
        if isinstance(m, dict) and m.get("role") == "user" and m.get("content"):
            t = re.sub(r"\s+", " ", m["content"]).strip()
            return t[:60] + ("…" if len(t) > 60 else "")
    return "New chat"


# --------------------------- Grammar & style helpers ----------------------- #

COMMON_FIXES = [
    (re.compile(r"\bi\b"), "I"),
    (re.compile(r"\s+"), " "),
    (re.compile(r"\s+([,.!?;:])"), r"\1"),
    (re.compile(r"([,.!?;:])([A-Za-z])"), lambda m: m.group(1) + " " + m.group(2).upper() if m.group(2).isalpha() else m.group(0)),
]


def _apply_local_fixes(text: str) -> str:
    out = text
    for pattern, repl in COMMON_FIXES:
        out = pattern.sub(repl, out)
    sentences = re.split(r"([.!?]\s+)", out)
    rebuilt = []
    for chunk in sentences:
        if chunk and chunk[0].isalpha():
            chunk = chunk[0].upper() + chunk[1:]
        rebuilt.append(chunk)
    out = "".join(rebuilt)
    return out.strip()


@app.post("/text/clean")
def text_clean(payload: Dict[str, Any] = Body(...)):
    text = payload.get("text") or ""
    cleaned = text
    if SETTINGS.ai_chat_key and SETTINGS.ai_chat_base:
        try:
            resp = requests.post(
                f"{SETTINGS.ai_chat_base.rstrip('/')}/chat/completions",
                headers={"Authorization": f"Bearer {SETTINGS.ai_chat_key}"},
                json={
                    "model": SETTINGS.ai_chat_model,
                    "messages": [
                        {"role": "system", "content": "You revise transcripts for grammar, punctuation, and readability. Return only the improved text."},
                        {"role": "user", "content": text[:8000]},
                    ],
                    "temperature": 0.2,
                },
                timeout=60,
            )
            if resp.status_code == 200:
                cleaned = resp.json()["choices"][0]["message"]["content"].strip()
                return {"text": cleaned, "engine": "ai"}
        except Exception as exc:
            logger.warning("AI cleanup failed: %s", exc)
    return {"text": _apply_local_fixes(text), "engine": "local-rules"}


@app.post("/text/diff")
def text_diff(payload: Dict[str, Any] = Body(...)):
    a = (payload.get("original") or "").splitlines()
    b = (payload.get("revised") or "").splitlines()
    import difflib
    diff = list(difflib.unified_diff(a, b, lineterm="", n=2))
    return {"diff": "\n".join(diff)}


# --------------------------- Web search proxy ------------------------------ #

@app.get("/search")
def search(q: str, request: Request):
    if not q:
        raise HTTPException(status_code=400, detail="q required")
    if SETTINGS.search_key and SETTINGS.search_base:
        try:
            r = requests.get(SETTINGS.search_base, headers={"Authorization": f"Bearer {SETTINGS.search_key}"}, params={"q": q}, timeout=15)
            if r.status_code == 200:
                return {"provider": "configured", "results": r.json().get("results", [])[:10]}
        except Exception as exc:
            logger.warning("search provider failed: %s", exc)
    try:
        r = requests.get(f"https://duckduckgo.com/html/?q={requests.utils.quote(q)}", headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        soup = BeautifulSoup(r.text, "html.parser")
        results = []
        for el in soup.select("a.result__a")[:10]:
            results.append({"title": el.get_text(strip=True), "url": el.get("href"), "snippet": ""})
        return {"provider": "duckduckgo-fallback", "results": results}
    except Exception as exc:
        return {"provider": "none", "results": [], "warning": str(exc)}


@app.get("/images/search")
def images_search(q: str):
    if not q:
        raise HTTPException(status_code=400, detail="q required")
    try:
        url = f"https://commons.wikimedia.org/w/api.php?action=query&format=json&list=search&srsearch={requests.utils.quote(q)}&srnamespace=6&srlimit=10"
        r = requests.get(url, timeout=10)
        data = r.json()
        items = []
        for hit in data.get("query", {}).get("search", []):
            title = hit.get("title", "")
            info_url = f"https://commons.wikimedia.org/w/api.php?action=query&format=json&prop=imageinfo&iiprop=url&iiurlwidth=800&titles={requests.utils.quote(title)}"
            try:
                ir = requests.get(info_url, timeout=8).json()
                pages = ir.get("query", {}).get("pages", {})
                for _, p in pages.items():
                    ii = (p.get("imageinfo") or [{}])[0]
                    items.append({"title": title, "thumb": ii.get("thumburl"), "url": ii.get("url"), "source": "wikimedia-commons"})
            except Exception:
                items.append({"title": title, "url": f"https://commons.wikimedia.org/wiki/{title.replace(' ', '_')}", "source": "wikimedia-commons"})
        return {"provider": "wikimedia-commons", "results": items}
    except Exception as exc:
        return {"provider": "none", "results": [], "warning": str(exc)}


# --------------------------- Misc helpers ---------------------------------- #

HTMLResponse = type("_H", (), {})  # placeholder, replaced below by FastAPI import
from fastapi.responses import HTMLResponse as _HTML
HTMLResponse = _HTML


# Patch the placeholder code in /auth/google/callback that used the alias above.
# (already defined above before the function; left here for clarity)
