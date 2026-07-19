# Tscript AI — Fixes Applied

## 1. 404 errors on AI Chat / Transcription (root cause)
Frontend (Vercel) and backend (Render) are separate origins. The frontend defaulted
to relative API URLs, which only works if the backend serves the page itself.
Fixed: `index.html` now defaults to `https://atc-transcriber.onrender.com` for all
API calls whenever the page isn't served directly by that backend (still
overridable via `?api=`, a meta tag, or `localStorage['tscript-api-url']`).

## 2. Session cookie wasn't reaching cross-origin requests
Once #1 made every request cross-origin, several endpoints
(`/transcribe`, `/dictate`, `/translate-text`, `/knowledge/*`, `/transcript/enrich`)
were calling `fetch()` without `credentials: 'include'`, so the session cookie
never made it across origins — those actions were being treated as anonymous even
when signed in. Added `credentials: 'include'` to every one of those calls.

## 3. Google/Firebase profile photo didn't persist
Added a `picture_url` column to `users` (safe migration for existing databases,
SQLite and Postgres). Saved on every Google/Firebase sign-in and returned from
`/auth/me`, so the real photo now survives reloads instead of falling back to
initials.

## 4. Noisy "ENV WARNING" log lines
SERPER/TAVILY/OCR/YouTube-key-missing messages were logged at `warning` level.
Downgraded to `info` — only a missing `GROQ_API_KEY` is still a real `error`.

## 5. CORS
Added `https://tscript-ai.vercel.app` and `https://atc-transcriber.onrender.com`
explicitly to the allowed-origins list, on top of the existing regex fallback.

## 6. Visual polish
Layered professional refinements onto the existing 4-theme system — smoother
shadows, hover lift on primary buttons, consistent focus rings, refined
scrollbars — without touching element IDs, layout structure, or JS bindings.

## 7. Language selection for transcription (new this round)
Added a "Spoken language" dropdown on the Transcription screen (auto-detect by
default) that is now wired into `/transcribe` and `/dictate`:
- **Auto-detect** (default) — unchanged, Whisper picks the language itself.
- **Explicit hint, officially supported languages** (English, French, Spanish,
  Arabic, Chinese, Hindi, Swahili, Yoruba, Hausa, Amharic, Afrikaans, Shona,
  Somali, Lingala, Malagasy, and 20+ more) — selecting one skips language
  *detection* and decodes directly in that language, which measurably improves
  accuracy for the language you already know is spoken.
- **Experimental (best-effort): Akan/Twi, Ewe, Ga, Igbo, Wolof** — these are
  **not in Whisper's training data**, so there is no `language=` code for them
  and no code change can add real coverage. Selecting one instead biases
  decoding with a short same-script prompt and still auto-detects — this may
  help marginally in mixed-language audio but will not turn Whisper into a
  fluent Ewe/Akan/Ga transcriber. The UI now shows this caveat inline when you
  pick one of these, instead of silently failing or pretending it's fully
  supported.
- Backend also now returns the languages Whisper actually detected per file, so
  you can see what it thought it heard.

**Bottom line on Ewe/Akan/Ga specifically:** genuinely accurate recognition for
these would require a different, specialized ASR model — that's a separate,
larger project (data + hosting), not a setting to flip. What's shipped here is
the most honest, useful version of "support" possible on top of Groq Whisper.

## Session persistence — verified correct
`SESSION_SAMESITE=none` + `secure=True` + `httponly=True`, 14-day TTL, stored
server-side — already supports "stay logged in across refresh/restart, only
Logout ends it" now that fix #2 ensures the cookie actually gets sent.

## Verified before packaging
- `api/main.py` parses with no Python syntax errors.
- Both real inline `<script>` blocks in `index.html` pass a JavaScript syntax
  check (the third `<script type="module">` block is Firebase init and is
  expected to only run as an ES module).

## Not yet done (from your "My Memory" / cloud-sync spec)
Rename History → My Memory as a unified inbox for chats + transcriptions,
guaranteed auto-save with no manual step, per-item delete + delete-all with
confirmation, and a dedicated YouTube Transcription section with an audio
fallback when captions are unavailable. Some pieces already exist (a "My Memory"
panel and `/history`, `/memory/*` endpoints; YouTube is already used inside AI
chat as an analysis tool) but not to the full spec, and there are duplicate
`/memory/*` route definitions to clean up first. This is a real feature build —
say the word and I'll scope and build it next.

## Redeploying
- **Render (backend)**: push `api/` as-is — no new environment variables required.
- **Vercel (frontend)**: push `index.html` / `vercel.json` / `.vercelignore` as-is.
