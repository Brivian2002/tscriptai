import io
import json
import logging
import mimetypes
import os
import re
import secrets
import tempfile
import zipfile
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

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

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("tscript-ai")

UTC = timezone.utc
APP_DIR = Path(__file__).resolve().parent
INDEX_FILE = APP_DIR / "index.html"
DATA_DIR = APP_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

MEDIA_EXTENSIONS = {".wav", ".mp3", ".m4a", ".flac", ".ogg", ".aac", ".wma", ".opus", ".mp4", ".mov", ".mkv", ".avi", ".webm", ".flv", ".wmv", ".m4v"}
TEXT_EXTENSIONS = {".txt", ".md", ".json", ".csv", ".tsv", ".log", ".py", ".js", ".ts", ".html", ".htm", ".css", ".xml", ".sql", ".yaml", ".yml", ".env"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff"}
DOC_EXTENSIONS = {".pdf", ".docx", ".pptx", ".xlsx"}

MODELS = {
    "standard": {"temperature": 0.4, "system": "You are Tscript AI. Be practical, precise, and concise."},
    "fast": {"temperature": 0.2, "system": "You are Tscript AI in fast mode. Answer directly with minimal fluff."},
    "think_deep": {"temperature": 0.25, "system": "You are Tscript AI in deep analysis mode. Reason carefully, structure your answer, and cite important evidence from the supplied context."},
    "advance": {"temperature": 0.35, "system": "You are Tscript AI in advanced mode. Provide a detailed, polished, production-quality answer with clear sections."},
    "annotation_expert": {"temperature": 0.1, "system": "You are Tscript AI. Focus on precise, standards-based annotation and labeling guidance."},
}


@dataclass
class Settings:
    app_name: str = os.getenv("APP_NAME", "Tscript AI")
    app_env: str = os.getenv("APP_ENV", "development")
    api_prefix: str = os.getenv("API_PREFIX", "/api/v1").rstrip("/") or ""
    frontend_url: str = os.getenv("FRONTEND_URL", "https://tscript-ai.vercel.app")
    cors_origins_raw: str = os.getenv("CORS_ORIGINS", "https://tscript-ai.vercel.app,http://localhost:3000,http://127.0.0.1:3000")
    app_secret: str = os.getenv("APP_SECRET", "change-me-in-production")
    neon_database_url: str = os.getenv("NEON_DATABASE_URL", "").strip()
    supabase_db_url: str = os.getenv("SUPABASE_DB_URL", "").strip()
    supabase_url: str = os.getenv("SUPABASE_URL", "").strip().rstrip("/")
    supabase_anon_key: str = os.getenv("SUPABASE_ANON_KEY", "").strip()
    supabase_jwt_audience: str = os.getenv("SUPABASE_JWT_AUDIENCE", "authenticated").strip() or "authenticated"
    groq_chat_api_key: str = os.getenv("GROQ_CHAT_API_KEY", "").strip()
    groq_chat_model: str = os.getenv("GROQ_CHAT_MODEL", "openai/gpt-oss-120b").strip()
    groq_transcription_api_key: str = os.getenv("GROQ_TRANSCRIPTION_API_KEY", "").strip()
    groq_transcription_model: str = os.getenv("GROQ_TRANSCRIPTION_MODEL", "whisper-large-v3-turbo").strip()
    google_api_key: str = os.getenv("GOOGLE_API_KEY", "").strip()
    google_maps_api_key: str = os.getenv("GOOGLE_MAPS_API_KEY", "").strip()
    max_upload_size_mb: int = int(os.getenv("MAX_UPLOAD_SIZE_MB", "200"))
    presence_ttl_seconds: int = int(os.getenv("PRESENCE_TTL_SECONDS", "120"))
    rate_limit_requests: int = int(os.getenv("RATE_LIMIT_REQUESTS", "30"))
    rate_limit_window_seconds: int = int(os.getenv("RATE_LIMIT_WINDOW_SECONDS", "60"))
    session_days: int = int(os.getenv("SESSION_TTL_DAYS", "14"))
    chunk_length_ms: int = int(os.getenv("TRANSCRIPTION_CHUNK_MS", str(8 * 60 * 1000)))

    @property
    def cors_origins(self) -> List[str]:
        items = [i.strip() for i in self.cors_origins_raw.split(",") if i.strip()]
        defaults = [self.frontend_url, "http://localhost:3000", "http://127.0.0.1:3000", "http://localhost:8000"]
        merged: List[str] = []
        for item in [*items, *defaults]:
            if item and item not in merged:
                merged.append(item)
        return merged


settings = Settings()

app = FastAPI(title=settings.app_name)
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

RATE_BUCKETS: Dict[str, List[float]] = {}
MEMORY_DB: Dict[str, List[Dict[str, Any]]] = {}
SESSION_FALLBACK: Dict[str, Dict[str, Any]] = {}


def now_utc() -> datetime:
    return datetime.now(UTC)


def iso_now() -> str:
    return now_utc().isoformat()


def trim_text(value: str, limit: int) -> str:
    value = (value or "").strip()
    return value if len(value) <= limit else value[: limit - 1].rstrip() + "…"


def chunked(items: List[Any], size: int) -> Iterable[List[Any]]:
    for idx in range(0, len(items), size):
        yield items[idx : idx + size]


def format_ts(seconds: float) -> str:
    seconds = max(0.0, float(seconds or 0.0))
    minutes = int(seconds // 60)
    sec = seconds - (minutes * 60)
    return f"{minutes:02d}:{sec:05.2f}"


def parse_json_response(resp: requests.Response) -> Dict[str, Any]:
    try:
        return resp.json()
    except Exception:
        return {"raw": resp.text}


def http_error(status: int, detail: str) -> HTTPException:
    return HTTPException(status_code=status, detail=detail)


def get_cookie_domain(request: Request) -> Optional[str]:
    host = request.url.hostname or ""
    if host.endswith("onrender.com"):
        return host
    return None


def set_security_headers(response: Response) -> None:
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Permissions-Policy"] = "microphone=(self), camera=()"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"


def session_cookie_kwargs(request: Request) -> Dict[str, Any]:
    return {
        "httponly": True,
        "secure": True,
        "samesite": "none",
        "max_age": settings.session_days * 24 * 60 * 60,
        "path": "/",
        "domain": get_cookie_domain(request),
    }


def public_actor_cookie_kwargs(request: Request) -> Dict[str, Any]:
    return {
        "httponly": True,
        "secure": True,
        "samesite": "none",
        "max_age": 365 * 24 * 60 * 60,
        "path": "/",
        "domain": get_cookie_domain(request),
    }


@contextmanager
def pg_connection(url: str):
    if not url:
        yield None
        return
    conn = None
    try:
        conn = psycopg2.connect(url)
        conn.autocommit = True
        yield conn
    except Exception as exc:
        logger.warning("Database connection failed: %s", exc)
        yield None
    finally:
        if conn:
            try:
                conn.close()
            except Exception:
                pass


def fetch_one(conn, query: str, params: Tuple[Any, ...] = ()) -> Optional[Dict[str, Any]]:
    if conn is None:
        return None
    with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
        cur.execute(query, params)
        row = cur.fetchone()
        return dict(row) if row else None


def fetch_all(conn, query: str, params: Tuple[Any, ...] = ()) -> List[Dict[str, Any]]:
    if conn is None:
        return []
    with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
        cur.execute(query, params)
        rows = cur.fetchall() or []
        return [dict(row) for row in rows]


def execute(conn, query: str, params: Tuple[Any, ...] = ()) -> None:
    if conn is None:
        return
    with conn.cursor() as cur:
        cur.execute(query, params)


def ensure_guest_id(request: Request, response: Optional[Response] = None) -> str:
    guest_id = request.cookies.get("tscript_guest_id")
    if guest_id:
        return guest_id
    guest_id = f"guest_{secrets.token_urlsafe(18)}"
    if response is not None:
        response.set_cookie("tscript_guest_id", guest_id, **public_actor_cookie_kwargs(request))
    return guest_id

def actor_from_request(request: Request, response: Optional[Response] = None, require_auth: bool = False) -> Dict[str, Any]:
    user = get_current_user(request)
    if user:
        return {"type": "user", "id": user["id"], "label": user.get("display_name") or user.get("email") or "User", "user": user}
    if require_auth:
        raise http_error(401, "Authentication required")
    guest_id = ensure_guest_id(request, response)
    return {"type": "guest", "id": guest_id, "label": "Guest", "user": None}


def rate_limit_key(request: Request, suffix: str) -> str:
    forwarded = request.headers.get("x-forwarded-for", "").split(",")[0].strip()
    host = forwarded or (request.client.host if request.client else "unknown")
    return f"{host}:{suffix}"


def enforce_rate_limit(request: Request, suffix: str) -> None:
    now = now_utc().timestamp()
    key = rate_limit_key(request, suffix)
    window = settings.rate_limit_window_seconds
    bucket = [ts for ts in RATE_BUCKETS.get(key, []) if now - ts < window]
    if len(bucket) >= settings.rate_limit_requests:
        raise http_error(429, "Too many requests. Please wait and try again.")
    bucket.append(now)
    RATE_BUCKETS[key] = bucket


def init_neon_schema() -> None:
    sql = """
    create extension if not exists pgcrypto;
    create schema if not exists app_core;
    create schema if not exists analytics;

    create table if not exists app_core.app_sessions (
      session_id text primary key,
      user_id text not null,
      email text,
      display_name text,
      access_token text not null,
      refresh_token text,
      expires_at timestamptz,
      created_at timestamptz not null default now(),
      updated_at timestamptz not null default now()
    );

    create table if not exists app_core.memories (
      id uuid primary key default gen_random_uuid(),
      actor_type text not null,
      actor_id text not null,
      memory text not null,
      memory_type text not null default 'general',
      enabled boolean not null default true,
      created_at timestamptz not null default now()
    );
    create index if not exists idx_memories_actor on app_core.memories(actor_type, actor_id, created_at desc);

    create table if not exists app_core.actor_settings (
      actor_type text not null,
      actor_id text not null,
      memory_enabled boolean not null default true,
      theme text,
      updated_at timestamptz not null default now(),
      primary key (actor_type, actor_id)
    );

    create table if not exists app_core.transcripts (
      id uuid primary key default gen_random_uuid(),
      actor_type text not null,
      actor_id text not null,
      user_id text,
      source_filename text,
      language text,
      summary text,
      paragraph_text text,
      clean_script text,
      raw_text text,
      utterances jsonb not null default '[]'::jsonb,
      created_at timestamptz not null default now(),
      updated_at timestamptz not null default now()
    );
    create index if not exists idx_transcripts_actor on app_core.transcripts(actor_type, actor_id, created_at desc);

    create table if not exists analytics.knowledge_chunks (
      id uuid primary key default gen_random_uuid(),
      transcript_id uuid not null references app_core.transcripts(id) on delete cascade,
      actor_type text not null,
      actor_id text not null,
      source_filename text,
      language text,
      chunk_index int not null,
      chunk_text text not null,
      created_at timestamptz not null default now()
    );
    create index if not exists idx_knowledge_actor on analytics.knowledge_chunks(actor_type, actor_id, created_at desc);

    create table if not exists app_core.chat_messages (
      id uuid primary key default gen_random_uuid(),
      actor_type text not null,
      actor_id text not null,
      role text not null,
      message text not null,
      attachment_name text,
      created_at timestamptz not null default now()
    );
    create index if not exists idx_chat_actor on app_core.chat_messages(actor_type, actor_id, created_at desc);
    """
    with pg_connection(settings.neon_database_url) as conn:
        if conn is None:
            return
        for statement in [s.strip() for s in sql.split(";") if s.strip()]:
            execute(conn, statement)


def init_supabase_schema() -> None:
    sql = """
    create extension if not exists pgcrypto;
    create table if not exists public.user_profiles (
      user_id text primary key,
      email text not null,
      display_name text,
      theme text,
      created_at timestamptz not null default now(),
      updated_at timestamptz not null default now()
    );

    create table if not exists public.user_presence (
      user_id text primary key references public.user_profiles(user_id) on delete cascade,
      status text not null default 'offline',
      current_page text,
      last_seen timestamptz not null default now(),
      created_at timestamptz not null default now(),
      updated_at timestamptz not null default now()
    );
    create index if not exists idx_user_presence_last_seen on public.user_presence(last_seen desc);

    create table if not exists public.public_posts (
      id uuid primary key default gen_random_uuid(),
      user_id text not null references public.user_profiles(user_id) on delete cascade,
      title text not null,
      body text not null,
      created_at timestamptz not null default now(),
      updated_at timestamptz not null default now()
    );

    create table if not exists public.public_post_comments (
      id uuid primary key default gen_random_uuid(),
      post_id uuid not null references public.public_posts(id) on delete cascade,
      user_id text not null references public.user_profiles(user_id) on delete cascade,
      body text not null,
      created_at timestamptz not null default now()
    );

    create table if not exists public.public_chat_messages (
      id uuid primary key default gen_random_uuid(),
      user_id text not null references public.user_profiles(user_id) on delete cascade,
      body text not null,
      created_at timestamptz not null default now()
    );
    """
    with pg_connection(settings.supabase_db_url) as conn:
        if conn is None:
            return
        for statement in [s.strip() for s in sql.split(";") if s.strip()]:
            execute(conn, statement)


@app.on_event("startup")
def startup() -> None:
    init_neon_schema()
    init_supabase_schema()


def supabase_headers() -> Dict[str, str]:
    return {
        "apikey": settings.supabase_anon_key,
        "Authorization": f"Bearer {settings.supabase_anon_key}",
        "Content-Type": "application/json",
    }


def upsert_user_profile(user_id: str, email: str, display_name: str = "", theme: str = "") -> None:
    with pg_connection(settings.supabase_db_url) as conn:
        execute(
            conn,
            """
            insert into public.user_profiles (user_id, email, display_name, theme, created_at, updated_at)
            values (%s, %s, %s, nullif(%s,''), now(), now())
            on conflict (user_id)
            do update set email = excluded.email,
                          display_name = coalesce(nullif(excluded.display_name,''), public.user_profiles.display_name),
                          theme = coalesce(nullif(excluded.theme,''), public.user_profiles.theme),
                          updated_at = now()
            """,
            (user_id, email, display_name, theme),
        )


def supabase_auth_request(path: str, payload: Dict[str, Any], query: str = "") -> Dict[str, Any]:
    if not settings.supabase_url or not settings.supabase_anon_key:
        raise http_error(500, "Supabase authentication is not configured")
    url = f"{settings.supabase_url}/auth/v1/{path}{query}"
    resp = requests.post(url, headers=supabase_headers(), json=payload, timeout=60)
    data = parse_json_response(resp)
    if resp.status_code >= 400:
        message = data.get("msg") or data.get("error_description") or data.get("error") or "Authentication request failed"
        raise http_error(resp.status_code, message)
    return data


def store_session(record: Dict[str, Any]) -> str:
    token = secrets.token_urlsafe(32)
    session_data = {
        "session_id": token,
        "user_id": record["user_id"],
        "email": record.get("email"),
        "display_name": record.get("display_name"),
        "access_token": record["access_token"],
        "refresh_token": record.get("refresh_token"),
        "expires_at": record.get("expires_at"),
        "updated_at": iso_now(),
    }
    with pg_connection(settings.neon_database_url) as conn:
        if conn is not None:
            execute(
                conn,
                """
                insert into app_core.app_sessions (session_id, user_id, email, display_name, access_token, refresh_token, expires_at, created_at, updated_at)
                values (%s,%s,%s,%s,%s,%s,%s,now(),now())
                on conflict (session_id)
                do update set user_id=excluded.user_id, email=excluded.email, display_name=excluded.display_name,
                              access_token=excluded.access_token, refresh_token=excluded.refresh_token,
                              expires_at=excluded.expires_at, updated_at=now()
                """,
                (
                    token,
                    session_data["user_id"],
                    session_data["email"],
                    session_data["display_name"],
                    session_data["access_token"],
                    session_data["refresh_token"],
                    session_data["expires_at"],
                ),
            )
        else:
            SESSION_FALLBACK[token] = session_data
    return token


def delete_session(session_id: str) -> None:
    SESSION_FALLBACK.pop(session_id, None)
    with pg_connection(settings.neon_database_url) as conn:
        execute(conn, "delete from app_core.app_sessions where session_id = %s", (session_id,))


def load_session(session_id: str) -> Optional[Dict[str, Any]]:
    if not session_id:
        return None
    with pg_connection(settings.neon_database_url) as conn:
        row = fetch_one(conn, "select * from app_core.app_sessions where session_id = %s", (session_id,))
        if row:
            return row
    return SESSION_FALLBACK.get(session_id)

def refresh_session_if_needed(session: Dict[str, Any]) -> Dict[str, Any]:
    expires_at = session.get("expires_at")
    if not expires_at:
        return session
    if isinstance(expires_at, str):
        try:
            expires_dt = datetime.fromisoformat(expires_at.replace("Z", "+00:00"))
        except Exception:
            return session
    else:
        expires_dt = expires_at
    if expires_dt - now_utc() > timedelta(minutes=3):
        return session
    refresh_token = session.get("refresh_token")
    if not refresh_token:
        return session
    data = supabase_auth_request("token", {"refresh_token": refresh_token}, "?grant_type=refresh_token")
    user = data.get("user") or {}
    refreshed = {
        "user_id": user.get("id") or session.get("user_id"),
        "email": user.get("email") or session.get("email"),
        "display_name": (user.get("user_metadata") or {}).get("display_name") or session.get("display_name"),
        "access_token": data.get("access_token") or session.get("access_token"),
        "refresh_token": data.get("refresh_token") or refresh_token,
        "expires_at": now_utc() + timedelta(seconds=int(data.get("expires_in") or 3600)),
    }
    token = session.get("session_id")
    with pg_connection(settings.neon_database_url) as conn:
        if conn is not None and token:
            execute(
                conn,
                "update app_core.app_sessions set access_token=%s, refresh_token=%s, expires_at=%s, email=%s, display_name=%s, updated_at=now() where session_id=%s",
                (
                    refreshed["access_token"],
                    refreshed["refresh_token"],
                    refreshed["expires_at"],
                    refreshed["email"],
                    refreshed["display_name"],
                    token,
                ),
            )
        elif token:
            SESSION_FALLBACK[token] = {**session, **refreshed}
    return {**session, **refreshed}


def get_current_user(request: Request) -> Optional[Dict[str, Any]]:
    session_id = request.cookies.get("tscript_session")
    session = load_session(session_id) if session_id else None
    if not session:
        return None
    try:
        session = refresh_session_if_needed({**session, "session_id": session_id})
    except HTTPException:
        delete_session(session_id)
        return None
    profile = None
    with pg_connection(settings.supabase_db_url) as conn:
        profile = fetch_one(conn, "select * from public.user_profiles where user_id = %s", (session.get("user_id"),))
    return {
        "id": session.get("user_id"),
        "email": session.get("email"),
        "display_name": (profile or {}).get("display_name") or session.get("display_name") or session.get("email"),
        "theme": (profile or {}).get("theme") or "dark",
        "access_token": session.get("access_token"),
    }


def require_user(request: Request) -> Dict[str, Any]:
    user = get_current_user(request)
    if not user:
        raise http_error(401, "Please sign in to continue")
    return user


def groq_headers(api_key: str) -> Dict[str, str]:
    return {"Authorization": f"Bearer {api_key}"}


def groq_chat(messages: List[Dict[str, str]], mode: str = "standard", max_tokens: int = 1600, temperature: Optional[float] = None) -> str:
    if not settings.groq_chat_api_key:
        raise http_error(500, "Groq chat is not configured")
    config = MODELS.get(mode, MODELS["standard"])
    payload = {
        "model": settings.groq_chat_model,
        "messages": [{"role": "system", "content": config["system"]}, *messages],
        "temperature": temperature if temperature is not None else config["temperature"],
        "max_tokens": max_tokens,
    }
    resp = requests.post(
        "https://api.groq.com/openai/v1/chat/completions",
        headers={**groq_headers(settings.groq_chat_api_key), "Content-Type": "application/json"},
        json=payload,
        timeout=180,
    )
    data = parse_json_response(resp)
    if resp.status_code >= 400:
        raise http_error(resp.status_code, data.get("error", {}).get("message") or "Chat request failed")
    return (((data.get("choices") or [{}])[0].get("message") or {}).get("content") or "").strip()


def transcribe_with_groq(file_path: Path, language_hint: str = "") -> Dict[str, Any]:
    if not settings.groq_transcription_api_key:
        raise http_error(500, "Groq transcription is not configured")
    with open(file_path, "rb") as fh:
        files = {"file": (file_path.name, fh, mimetypes.guess_type(file_path.name)[0] or "application/octet-stream")}
        data = {"model": settings.groq_transcription_model, "response_format": "verbose_json"}
        if language_hint and language_hint != "auto":
            data["language"] = language_hint
        resp = requests.post(
            "https://api.groq.com/openai/v1/audio/transcriptions",
            headers=groq_headers(settings.groq_transcription_api_key),
            files=files,
            data=data,
            timeout=300,
        )
    payload = parse_json_response(resp)
    if resp.status_code >= 400:
        raise http_error(resp.status_code, payload.get("error", {}).get("message") or payload.get("error") or "Transcription request failed")
    return payload


def load_audiosegment(path: Path) -> AudioSegment:
    return AudioSegment.from_file(path)


def local_paragraph_from_utterances(utterances: List[Dict[str, Any]]) -> str:
    lines = []
    current = []
    last_speaker = None
    for utterance in utterances:
        speaker = (utterance.get("speaker_label") or "Speaker").strip()
        text = (utterance.get("transcription") or "").strip()
        if not text:
            continue
        if current and (speaker != last_speaker or len(" ".join(current)) > 700):
            lines.append(" ".join(current).strip())
            current = []
        prefix = f"{speaker}: " if not current else ""
        current.append(prefix + text)
        last_speaker = speaker
    if current:
        lines.append(" ".join(current).strip())
    return "\n\n".join(lines)


def clean_script_text(utterances: List[Dict[str, Any]]) -> str:
    parts = []
    for item in utterances:
        text = re.sub(r"\s+", " ", item.get("transcription") or "").strip()
        if text:
            parts.append(text)
    return " ".join(parts)


def normalise_utterances(segments: List[Dict[str, Any]], offset: float = 0.0) -> List[Dict[str, Any]]:
    utterances = []
    for idx, seg in enumerate(segments or [], start=1):
        start = float(seg.get("start") or 0.0) + offset
        end = float(seg.get("end") or start) + offset
        utterances.append(
            {
                "id": f"utt_{idx}_{secrets.token_hex(3)}",
                "index": idx,
                "speaker_label": "Speaker A",
                "speaker_role": "Speaker A",
                "speaker_name": "",
                "role_tag": "Unknown",
                "speaker_callsign": "",
                "transcription": (seg.get("text") or seg.get("transcription") or "").strip(),
                "notes": "",
                "time": {
                    "start": start,
                    "end": end,
                    "start_str": format_ts(start),
                    "end_str": format_ts(end),
                },
            }
        )
    return utterances


def save_transcript(actor: Dict[str, Any], filename: str, language: str, utterances: List[Dict[str, Any]], summary: str = "") -> Optional[str]:
    paragraph = local_paragraph_from_utterances(utterances)
    clean = clean_script_text(utterances)
    raw = "\n".join([(u.get("transcription") or "").strip() for u in utterances if (u.get("transcription") or "").strip()])
    with pg_connection(settings.neon_database_url) as conn:
        if conn is None:
            return None
        row = fetch_one(
            conn,
            """
            insert into app_core.transcripts (actor_type, actor_id, user_id, source_filename, language, summary, paragraph_text, clean_script, raw_text, utterances, created_at, updated_at)
            values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s::jsonb,now(),now())
            returning id::text
            """,
            (
                actor["type"],
                actor["id"],
                actor["user"]["id"] if actor.get("user") else None,
                filename,
                language,
                summary,
                paragraph,
                clean,
                raw,
                json.dumps(utterances),
            ),
        )
        transcript_id = row["id"] if row else None
        if transcript_id:
            execute(conn, "delete from analytics.knowledge_chunks where transcript_id = %s::uuid", (transcript_id,))
            chunks = [c for c in chunked([u.get("transcription") or "" for u in utterances], 8)]
            for idx, items in enumerate(chunks, start=1):
                text = "\n".join([item.strip() for item in items if item.strip()])
                if text:
                    execute(
                        conn,
                        "insert into analytics.knowledge_chunks (transcript_id, actor_type, actor_id, source_filename, language, chunk_index, chunk_text, created_at) values (%s::uuid,%s,%s,%s,%s,%s,%s,now())",
                        (transcript_id, actor["type"], actor["id"], filename, language, idx, text),
                    )
        return transcript_id

def extract_zip_media(path: Path) -> Path:
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            lowered = name.lower()
            if any(lowered.endswith(ext) for ext in MEDIA_EXTENSIONS):
                out_path = path.parent / Path(name).name
                with archive.open(name) as src, open(out_path, "wb") as dst:
                    dst.write(src.read())
                return out_path
    raise http_error(400, "No supported audio or video file was found inside the ZIP archive")


def save_upload_to_path(upload: UploadFile) -> Path:
    suffix = Path(upload.filename or "upload.bin").suffix or ".bin"
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    data = upload.file.read()
    if len(data) > settings.max_upload_size_mb * 1024 * 1024:
        raise http_error(413, f"File exceeds {settings.max_upload_size_mb} MB limit")
    temp.write(data)
    temp.flush()
    temp.close()
    return Path(temp.name)


def extract_text_from_path(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in TEXT_EXTENSIONS:
        text = path.read_text(encoding="utf-8", errors="ignore")
        if ext in {".html", ".htm"}:
            return BeautifulSoup(text, "html.parser").get_text("\n")
        return text
    if ext == ".pdf":
        reader = PdfReader(str(path))
        return "\n".join([(page.extract_text() or "") for page in reader.pages])
    if ext == ".docx":
        doc = Document(str(path))
        return "\n".join([p.text for p in doc.paragraphs])
    if ext == ".pptx":
        prs = Presentation(str(path))
        parts: List[str] = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text:
                    parts.append(shape.text)
        return "\n".join(parts)
    if ext == ".xlsx":
        wb = load_workbook(str(path), read_only=True, data_only=True)
        lines: List[str] = []
        for sheet in wb.worksheets:
            lines.append(f"# Sheet: {sheet.title}")
            for row in sheet.iter_rows(values_only=True):
                vals = [str(v) for v in row if v is not None and str(v).strip()]
                if vals:
                    lines.append(" | ".join(vals))
        return "\n".join(lines)
    if ext in IMAGE_EXTENSIONS:
        try:
            return pytesseract.image_to_string(Image.open(str(path)))
        except Exception:
            return ""
    return ""


def build_transcript_from_file(path: Path, language_hint: str = "") -> Tuple[str, List[Dict[str, Any]]]:
    original = path
    if path.suffix.lower() == ".zip":
        path = extract_zip_media(path)
    audio = load_audiosegment(path)
    utterances: List[Dict[str, Any]] = []
    if len(audio) <= settings.chunk_length_ms:
        payload = transcribe_with_groq(path, language_hint)
        utterances = normalise_utterances(payload.get("segments") or [{"start": 0, "end": audio.duration_seconds, "text": payload.get("text") or ""}])
        return (payload.get("language") or language_hint or "auto"), utterances
    with tempfile.TemporaryDirectory() as tempdir:
        offset_sec = 0.0
        index = 0
        detected_language = language_hint or "auto"
        for start in range(0, len(audio), settings.chunk_length_ms):
            index += 1
            clip = audio[start : start + settings.chunk_length_ms]
            clip_path = Path(tempdir) / f"chunk_{index}.mp3"
            clip.export(str(clip_path), format="mp3", bitrate="128k")
            payload = transcribe_with_groq(clip_path, language_hint)
            detected_language = payload.get("language") or detected_language
            segments = payload.get("segments") or [{"start": 0, "end": clip.duration_seconds, "text": payload.get("text") or ""}]
            utterances.extend(normalise_utterances(segments, offset_sec))
            offset_sec += clip.duration_seconds
    for idx, item in enumerate(utterances, start=1):
        item["index"] = idx
        item["id"] = item.get("id") or f"utt_{idx}_{secrets.token_hex(3)}"
    return (detected_language or "auto"), utterances


def transcript_summary_fallback(utterances: List[Dict[str, Any]]) -> str:
    text = clean_script_text(utterances)
    return trim_text(text, 420) or "Transcript ready for review."


def enrich_transcript_data(utterances: List[Dict[str, Any]], target_language: str) -> Dict[str, Any]:
    paragraph = local_paragraph_from_utterances(utterances)
    clean = clean_script_text(utterances)
    speakers = sorted({u.get("speaker_label") or "Speaker A" for u in utterances})
    prompt = (
        "Return strict JSON with keys: summary, translated_paragraph, highlights, speakers. "
        "highlights must be an array of short bullet strings. "
        "speakers must be an array of objects with speaker_label, speaker_name, role_tag. "
        f"Target translation language: {target_language}.\n\nTranscript:\n{paragraph[:24000]}"
    )
    try:
        raw = groq_chat([{"role": "user", "content": prompt}], mode="advance", max_tokens=1400, temperature=0.2)
        match = re.search(r"\{.*\}", raw, re.S)
        parsed = json.loads(match.group(0) if match else raw)
    except Exception:
        parsed = {}
    speaker_cards = parsed.get("speakers") or [{"speaker_label": label, "speaker_name": "", "role_tag": "Unknown"} for label in speakers]
    highlights = [trim_text(str(item), 180) for item in (parsed.get("highlights") or []) if str(item).strip()][:8]
    translated = trim_text(str(parsed.get("translated_paragraph") or paragraph), 40000)
    summary = trim_text(str(parsed.get("summary") or transcript_summary_fallback(utterances)), 3000)
    return {
        "summary": summary,
        "paragraph_text": paragraph,
        "clean_script": clean,
        "translated_paragraph": translated,
        "highlights": highlights,
        "speakers": speaker_cards,
    }


def recent_chat_history(actor: Dict[str, Any], limit: int = 12) -> List[Dict[str, str]]:
    with pg_connection(settings.neon_database_url) as conn:
        rows = fetch_all(
            conn,
            "select role, message from app_core.chat_messages where actor_type=%s and actor_id=%s order by created_at desc limit %s",
            (actor["type"], actor["id"], limit),
        )
    rows.reverse()
    return [{"role": row["role"], "content": row["message"]} for row in rows]


def persist_chat_message(actor: Dict[str, Any], role: str, message: str, attachment_name: str = "") -> None:
    with pg_connection(settings.neon_database_url) as conn:
        execute(
            conn,
            "insert into app_core.chat_messages (actor_type, actor_id, role, message, attachment_name, created_at) values (%s,%s,%s,%s,%s,now())",
            (actor["type"], actor["id"], role, message, attachment_name),
        )


def update_presence(user: Dict[str, Any], page: str = "", status: str = "online") -> None:
    with pg_connection(settings.supabase_db_url) as conn:
        execute(
            conn,
            """
            insert into public.user_presence (user_id, status, current_page, last_seen, created_at, updated_at)
            values (%s,%s,%s,now(),now(),now())
            on conflict (user_id)
            do update set status=excluded.status, current_page=excluded.current_page, last_seen=now(), updated_at=now()
            """,
            (user["id"], status, page),
        )


def serialise_user(user: Optional[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    if not user:
        return None
    return {"id": user["id"], "email": user["email"], "display_name": user.get("display_name") or user["email"], "theme": user.get("theme") or "dark"}


def json_ok(payload: Dict[str, Any], response: Optional[Response] = None) -> JSONResponse:
    result = JSONResponse(payload)
    set_security_headers(result)
    if response:
        for key, value in response.headers.items():
            result.headers[key] = value
    return result


def route(path: str) -> str:
    return f"{settings.api_prefix}{path}" if settings.api_prefix else path


@app.get(route("/health"))
def health() -> JSONResponse:
    return json_ok({"ok": True, "app": settings.app_name, "env": settings.app_env, "api_prefix": settings.api_prefix or "/"})


@app.get("/health")
def health_root() -> JSONResponse:
    return health()


@app.get(route("/public/config"))
def public_config() -> JSONResponse:
    return json_ok({
        "app_name": settings.app_name,
        "api_prefix": settings.api_prefix or "/",
        "frontend_url": settings.frontend_url,
        "google_maps_enabled": bool(settings.google_maps_api_key),
    })


@app.get(route("/google/status"))
def google_status() -> JSONResponse:
    return json_ok({
        "google_api_enabled": bool(settings.google_api_key),
        "google_maps_enabled": bool(settings.google_maps_api_key),
        "recommended_usage": ["maps", "geolocation"],
    })


@app.post(route("/auth/register"))
def auth_register(request: Request, response: Response, payload: Dict[str, Any] = Body(...)) -> JSONResponse:
    enforce_rate_limit(request, "auth")
    email = trim_text(str(payload.get("email") or ""), 240).lower()
    password = str(payload.get("password") or "")
    display_name = trim_text(str(payload.get("display_name") or ""), 120)
    if not email or "@" not in email:
        raise http_error(400, "A valid email address is required")
    if len(password) < 8:
        raise http_error(400, "Password must be at least 8 characters long")
    data = supabase_auth_request("signup", {"email": email, "password": password, "data": {"display_name": display_name}})
    user = data.get("user") or {}
    if user.get("id"):
        upsert_user_profile(user["id"], email, display_name)
    session_payload = data.get("session") or {}
    if session_payload.get("access_token") and user.get("id"):
        session_id = store_session({
            "user_id": user["id"],
            "email": email,
            "display_name": display_name,
            "access_token": session_payload.get("access_token"),
            "refresh_token": session_payload.get("refresh_token"),
            "expires_at": now_utc() + timedelta(seconds=int(session_payload.get("expires_in") or 3600)),
        })
        response.set_cookie("tscript_session", session_id, **session_cookie_kwargs(request))
    return json_ok({"ok": True, "message": "Account created successfully.", "email_confirmation_required": not bool(data.get("session"))}, response)


@app.post(route("/auth/login"))
def auth_login(request: Request, response: Response, payload: Dict[str, Any] = Body(...)) -> JSONResponse:
    enforce_rate_limit(request, "auth")
    email = trim_text(str(payload.get("email") or ""), 240).lower()
    password = str(payload.get("password") or "")
    data = supabase_auth_request("token", {"email": email, "password": password}, "?grant_type=password")
    user = data.get("user") or {}
    display_name = ((user.get("user_metadata") or {}).get("display_name") or "").strip()
    if user.get("id"):
        upsert_user_profile(user["id"], user.get("email") or email, display_name)
        update_presence({"id": user["id"]}, "login", "online")
    session_id = store_session({
        "user_id": user.get("id"),
        "email": user.get("email") or email,
        "display_name": display_name,
        "access_token": data.get("access_token"),
        "refresh_token": data.get("refresh_token"),
        "expires_at": now_utc() + timedelta(seconds=int(data.get("expires_in") or 3600)),
    })
    response.set_cookie("tscript_session", session_id, **session_cookie_kwargs(request))
    return json_ok({"ok": True, "user": {"id": user.get("id"), "email": user.get("email") or email, "display_name": display_name or user.get("email") or email}}, response)


@app.post(route("/auth/logout"))
def auth_logout(request: Request, response: Response) -> JSONResponse:
    session_id = request.cookies.get("tscript_session")
    user = get_current_user(request)
    if user:
        update_presence(user, "logout", "offline")
    if session_id:
        delete_session(session_id)
    response.delete_cookie("tscript_session", path="/", domain=get_cookie_domain(request), secure=True, samesite="none")
    return json_ok({"ok": True}, response)


@app.get(route("/auth/me"))
def auth_me(request: Request) -> JSONResponse:
    return json_ok({"authenticated": bool(get_current_user(request)), "user": serialise_user(get_current_user(request))})


@app.post(route("/auth/profile"))
def auth_profile(request: Request, payload: Dict[str, Any] = Body(...)) -> JSONResponse:
    user = require_user(request)
    display_name = trim_text(str(payload.get("display_name") or ""), 120)
    theme = trim_text(str(payload.get("theme") or ""), 20)
    with pg_connection(settings.supabase_db_url) as conn:
        execute(conn, "update public.user_profiles set display_name=%s, theme=nullif(%s,''), updated_at=now() where user_id=%s", (display_name or user["display_name"], theme, user["id"]))
    return json_ok({"ok": True})


@app.post(route("/presence/heartbeat"))
def presence_heartbeat(request: Request, payload: Dict[str, Any] = Body(...)) -> JSONResponse:
    user = require_user(request)
    update_presence(user, trim_text(str(payload.get("page") or ""), 60), "online")
    return json_ok({"ok": True})


@app.get(route("/presence/online"))
def presence_online() -> JSONResponse:
    cutoff = now_utc() - timedelta(seconds=settings.presence_ttl_seconds)
    with pg_connection(settings.supabase_db_url) as conn:
        rows = fetch_all(
            conn,
            """
            select p.user_id, p.status, p.current_page, p.last_seen, u.display_name, u.email
            from public.user_presence p
            join public.user_profiles u on u.user_id = p.user_id
            order by p.last_seen desc
            limit 40
            """,
        )
    users = []
    for row in rows:
        last_seen = row.get("last_seen")
        active = bool(last_seen and last_seen >= cutoff)
        users.append({
            "user_id": row.get("user_id"),
            "display_name": row.get("display_name") or row.get("email") or "User",
            "status": "online" if active else "offline",
            "current_page": row.get("current_page") or "",
            "last_seen": last_seen.isoformat() if hasattr(last_seen, "isoformat") else str(last_seen or ""),
        })
    return json_ok({"users": users})


@app.get(route("/public/posts"))
def public_posts() -> JSONResponse:
    with pg_connection(settings.supabase_db_url) as conn:
        posts = fetch_all(
            conn,
            """
            select p.id::text, p.title, p.body, p.created_at,
                   u.display_name, u.email,
                   (select count(*) from public.public_post_comments c where c.post_id = p.id) as comment_count
            from public.public_posts p
            join public.user_profiles u on u.user_id = p.user_id
            order by p.created_at desc
            limit 50
            """,
        )
    return json_ok({"items": posts})


@app.post(route("/public/posts"))
def create_public_post(request: Request, payload: Dict[str, Any] = Body(...)) -> JSONResponse:
    user = require_user(request)
    title = trim_text(str(payload.get("title") or ""), 160)
    body = trim_text(str(payload.get("body") or ""), 8000)
    if not title or not body:
        raise http_error(400, "Title and body are required")
    with pg_connection(settings.supabase_db_url) as conn:
        execute(conn, "insert into public.public_posts (user_id, title, body, created_at, updated_at) values (%s,%s,%s,now(),now())", (user["id"], title, body))
    return json_ok({"ok": True})


@app.get(route("/public/posts/{post_id}/comments"))
def public_post_comments(post_id: str) -> JSONResponse:
    with pg_connection(settings.supabase_db_url) as conn:
        items = fetch_all(
            conn,
            """
            select c.id::text, c.body, c.created_at, u.display_name, u.email
            from public.public_post_comments c
            join public.user_profiles u on u.user_id = c.user_id
            where c.post_id = %s::uuid
            order by c.created_at asc
            """,
            (post_id,),
        )
    return json_ok({"items": items})


@app.post(route("/public/posts/{post_id}/comments"))
def create_public_post_comment(post_id: str, request: Request, payload: Dict[str, Any] = Body(...)) -> JSONResponse:
    user = require_user(request)
    body = trim_text(str(payload.get("body") or ""), 3000)
    if not body:
        raise http_error(400, "Comment body is required")
    with pg_connection(settings.supabase_db_url) as conn:
        execute(conn, "insert into public.public_post_comments (post_id, user_id, body, created_at) values (%s::uuid,%s,%s,now())", (post_id, user["id"], body))
    return json_ok({"ok": True})


@app.get(route("/public/chat/messages"))
def public_chat_messages() -> JSONResponse:
    with pg_connection(settings.supabase_db_url) as conn:
        items = fetch_all(
            conn,
            """
            select m.id::text, m.body, m.created_at, u.display_name, u.email
            from public.public_chat_messages m
            join public.user_profiles u on u.user_id = m.user_id
            order by m.created_at desc
            limit 60
            """,
        )
    items.reverse()
    return json_ok({"items": items})


@app.post(route("/public/chat/messages"))
def create_public_chat_message(request: Request, payload: Dict[str, Any] = Body(...)) -> JSONResponse:
    user = require_user(request)
    body = trim_text(str(payload.get("body") or ""), 1200)
    if not body:
        raise http_error(400, "Message body is required")
    with pg_connection(settings.supabase_db_url) as conn:
        execute(conn, "insert into public.public_chat_messages (user_id, body, created_at) values (%s,%s,now())", (user["id"], body))
    return json_ok({"ok": True})


@app.get(route("/memory/list"))
def memory_list(request: Request, response: Response) -> JSONResponse:
    actor = actor_from_request(request, response)
    with pg_connection(settings.neon_database_url) as conn:
        memories = fetch_all(conn, "select id::text, memory, memory_type, created_at from app_core.memories where actor_type=%s and actor_id=%s and enabled=true order by created_at desc limit 100", (actor["type"], actor["id"]))
        settings_row = fetch_one(conn, "select memory_enabled from app_core.actor_settings where actor_type=%s and actor_id=%s", (actor["type"], actor["id"])) if conn else None
    if not memories:
        memories = MEMORY_DB.get(f"{actor['type']}:{actor['id']}", [])
    return json_ok({"memories": memories, "memory_enabled": (settings_row or {}).get("memory_enabled", True)}, response)


@app.post(route("/memory/update"))
def memory_update(request: Request, response: Response, payload: Dict[str, Any] = Body(...)) -> JSONResponse:
    actor = actor_from_request(request, response)
    with pg_connection(settings.neon_database_url) as conn:
        if "enabled" in payload:
            execute(conn, "insert into app_core.actor_settings (actor_type, actor_id, memory_enabled, updated_at) values (%s,%s,%s,now()) on conflict (actor_type, actor_id) do update set memory_enabled=excluded.memory_enabled, updated_at=now()", (actor["type"], actor["id"], bool(payload.get("enabled"))))
        note = trim_text(str(payload.get("note") or ""), 1000)
        if note:
            if conn is not None:
                execute(conn, "insert into app_core.memories (actor_type, actor_id, memory, memory_type, enabled, created_at) values (%s,%s,%s,'general',true,now())", (actor["type"], actor["id"], note))
            else:
                MEMORY_DB.setdefault(f"{actor['type']}:{actor['id']}", []).insert(0, {"id": secrets.token_hex(8), "memory": note, "memory_type": "general", "created_at": iso_now()})
    return json_ok({"ok": True}, response)


@app.post(route("/memory/delete"))
def memory_delete(request: Request, response: Response, payload: Dict[str, Any] = Body(...)) -> JSONResponse:
    actor = actor_from_request(request, response)
    mem_id = str(payload.get("memory_id") or "")
    with pg_connection(settings.neon_database_url) as conn:
        execute(conn, "delete from app_core.memories where actor_type=%s and actor_id=%s and id=%s::uuid", (actor["type"], actor["id"], mem_id))
    MEMORY_DB[f"{actor['type']}:{actor['id']}"] = [m for m in MEMORY_DB.get(f"{actor['type']}:{actor['id']}", []) if m.get("id") != mem_id]
    return json_ok({"ok": True}, response)


@app.delete(route("/memory/clear"))
def memory_clear(request: Request, response: Response) -> JSONResponse:
    actor = actor_from_request(request, response)
    with pg_connection(settings.neon_database_url) as conn:
        execute(conn, "delete from app_core.memories where actor_type=%s and actor_id=%s", (actor["type"], actor["id"]))
    MEMORY_DB[f"{actor['type']}:{actor['id']}"] = []
    return json_ok({"ok": True}, response)


@app.post(route("/dictate"))
def dictate(request: Request, response: Response, file: UploadFile = File(...), language_hint: str = Form("")) -> JSONResponse:
    actor_from_request(request, response)
    file_path = save_upload_to_path(file)
    try:
        payload = transcribe_with_groq(file_path, language_hint)
        return json_ok({"text": trim_text(str(payload.get("text") or ""), 20000)}, response)
    finally:
        file_path.unlink(missing_ok=True)


@app.post(route("/transcribe"))
def transcribe(request: Request, response: Response, file: UploadFile = File(...), language_hint: str = Form("")) -> JSONResponse:
    enforce_rate_limit(request, "transcribe")
    actor = actor_from_request(request, response)
    file_path = save_upload_to_path(file)
    try:
        language, utterances = build_transcript_from_file(file_path, language_hint)
        transcript_id = save_transcript(actor, file.filename or file_path.name, language, utterances)
        return json_ok({
            "transcript_id": transcript_id or "",
            "language": language,
            "utterances": utterances,
            "paragraph_text": local_paragraph_from_utterances(utterances),
            "clean_script": clean_script_text(utterances),
            "speakers": [{"speaker_label": "Speaker A", "speaker_name": "", "role_tag": "Unknown"}],
        }, response)
    finally:
        file_path.unlink(missing_ok=True)


@app.post(route("/transcript/enrich"))
def transcript_enrich(payload: Dict[str, Any] = Body(...)) -> JSONResponse:
    utterances = payload.get("utterances") or []
    if not utterances:
        raise http_error(400, "No utterances supplied")
    target_language = trim_text(str(payload.get("target_language") or "English"), 40)
    data = enrich_transcript_data(utterances, target_language)
    return json_ok({"utterances": utterances, **data, "language": target_language})


@app.post(route("/translate-text"))
def translate_text(payload: Dict[str, Any] = Body(...)) -> JSONResponse:
    text = str(payload.get("text") or "").strip()
    target_language = trim_text(str(payload.get("target_language") or "English"), 40)
    if not text:
        raise http_error(400, "No text supplied")
    prompt = f"Translate the following text into {target_language}. Return only the translated text and preserve separators exactly, including --- markers.\n\n{text}"
    translated = groq_chat([{"role": "user", "content": prompt}], mode="standard", max_tokens=2000, temperature=0.1)
    return json_ok({"translated": translated})


@app.get(route("/knowledge/list"))
def knowledge_list(request: Request, response: Response) -> JSONResponse:
    actor = actor_from_request(request, response)
    with pg_connection(settings.neon_database_url) as conn:
        items = fetch_all(conn, "select id::text, source_filename, language, summary, created_at from app_core.transcripts where actor_type=%s and actor_id=%s order by created_at desc limit 20", (actor["type"], actor["id"]))
    return json_ok({"items": items}, response)


@app.get(route("/knowledge/search"))
def knowledge_search(request: Request, response: Response, q: str = "") -> JSONResponse:
    actor = actor_from_request(request, response)
    query = trim_text(q, 240)
    if not query:
        return json_ok({"items": []}, response)
    like = f"%{query}%"
    with pg_connection(settings.neon_database_url) as conn:
        items = fetch_all(conn, "select source_filename, language, chunk_text as snippet, chunk_index as score from analytics.knowledge_chunks where actor_type=%s and actor_id=%s and chunk_text ilike %s order by created_at desc limit 12", (actor["type"], actor["id"], like))
    return json_ok({"items": items}, response)


@app.post(route("/knowledge/ask"))
def knowledge_ask(request: Request, response: Response, payload: Dict[str, Any] = Body(...)) -> JSONResponse:
    actor = actor_from_request(request, response)
    question = trim_text(str(payload.get("question") or ""), 600)
    if not question:
        raise http_error(400, "Question is required")
    with pg_connection(settings.neon_database_url) as conn:
        rows = fetch_all(conn, "select source_filename, language, chunk_text from analytics.knowledge_chunks where actor_type=%s and actor_id=%s order by created_at desc limit 18", (actor["type"], actor["id"]))
    citations = rows[:6]
    context = "\n\n".join([f"[{idx+1}] {row.get('source_filename') or 'Transcript'}\n{row.get('chunk_text') or ''}" for idx, row in enumerate(citations)])
    answer = groq_chat([{"role": "user", "content": f"Answer the question using only the transcript memory below when possible.\n\nQuestion: {question}\n\nTranscript memory:\n{context}"}], mode="advance", max_tokens=1200, temperature=0.2)
    citation_rows = []
    for row in citations:
        citation_rows.append({"speaker": row.get("source_filename") or "Transcript", "start_str": "", "end_str": "", "role_tag": row.get("language") or "", "text": row.get("chunk_text") or ""})
    return json_ok({"answer": answer, "citations": citation_rows}, response)


@app.post(route("/chat"))
def chat(request: Request, response: Response, message: str = Form(""), mode: str = Form("standard"), file: Optional[UploadFile] = File(None)) -> JSONResponse:
    enforce_rate_limit(request, "chat")
    actor = actor_from_request(request, response)
    prompt = trim_text(message, 4000)
    attachment_context = ""
    attachment_name = ""
    if file is not None and file.filename:
        attachment_name = file.filename
        temp_path = save_upload_to_path(file)
        try:
            ext = temp_path.suffix.lower()
            if ext in MEDIA_EXTENSIONS or ext == ".zip":
                language, utterances = build_transcript_from_file(temp_path)
                transcript_text = local_paragraph_from_utterances(utterances)
                attachment_context = f"\n\nAttached media transcript ({language}):\n{trim_text(transcript_text, 24000)}"
                save_transcript(actor, file.filename, language, utterances, "")
            else:
                extracted = trim_text(extract_text_from_path(temp_path), 24000)
                attachment_context = f"\n\nAttached file content ({file.filename}):\n{extracted}" if extracted else f"\n\nAttached file name: {file.filename}"
        finally:
            temp_path.unlink(missing_ok=True)
    history = recent_chat_history(actor)
    user_message = prompt or ("Please analyze the attached file." if attachment_context else "Hello")
    messages = [*history, {"role": "user", "content": f"{user_message}{attachment_context}"}]
    reply = groq_chat(messages, mode=mode if mode in MODELS else "standard", max_tokens=1800)
    persist_chat_message(actor, "user", user_message, attachment_name)
    persist_chat_message(actor, "assistant", reply)
    return json_ok({"reply": reply}, response)


@app.get("/")
def root() -> Response:
    return FileResponse(INDEX_FILE) if INDEX_FILE.exists() else PlainTextResponse("Tscript AI backend is running")


@app.get("/{full_path:path}")
def spa_fallback(full_path: str) -> Response:
    if full_path.startswith(settings.api_prefix.lstrip("/")):
        raise http_error(404, "Not found")
    return FileResponse(INDEX_FILE) if INDEX_FILE.exists() else PlainTextResponse("Tscript AI backend is running")
