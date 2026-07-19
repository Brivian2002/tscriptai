# Tscript AI Workspace Refresh

## What changed
- Preserved the existing Tscript AI visual language while upgrading the workspace behavior.
- Grouped **About** and **Support Us** under **Support & Legal** in the sidebar.
- Added new sidebar destinations for **Vibe Coding** and **Artifacts**.
- Reworked the transcription assistant area into a compact tab system:
  - Transcript
  - Summary
  - AI Workspace
  - Notes
- Added collapsible cards with smaller headers, icons, and collapse arrows.
- Added a focused **Voice Recorder** card that keeps only:
  - Record
  - Timer / waveform
  - Transcribe
- Kept the live voice agent inside **AI Chat** behind the headset control.

## AI Chat upgrades
- Added a mode selector beside **Tscript AI · Standard**:
  - Standard
  - Deep Research
  - Structured Code Output
  - Analyze Images
  - URL Analyze
  - Web Scraping
- Added browser-side chat retention with automatic cleanup after **1 hour**.
- Added read-aloud controls for AI replies.
- Added pinned messages, inline source pills, and a usage indicator.
- Added browser-local chat persistence plus a visible privacy/storage note.
- Added keyboard shortcuts:
  - `Ctrl/Cmd + Enter`
  - `Ctrl/Cmd + K`
- Added onboarding guidance and a quick theme toggle.

## File and web support
The chat workspace now supports:
- Images
- PDF
- Word documents
- Excel spreadsheets
- Text and code files
- Video and audio uploads
- Direct URLs
- YouTube links
- Website analysis
- Current web search when requested
- Recent information when available

## Transcription upgrades
- Preserved speaker labeling workflows.
- Added confidence-aware review highlighting from transcription metadata.
- Kept timestamp-based playback and export support.
- Exports still include:
  - TXT
  - PDF
  - DOCX
  - SRT
  - VTT
  - JSON / CSV / XML

## Artifacts and Vibe Coding
### Vibe Coding
- Opens the AI Chat in structured code mode.
- Supports file-oriented answers, code previews, editing, and ZIP export through Canvas.

### Artifacts
- Accepts Word, Excel, PDF, text, ZIP, and image-style document inputs.
- Routes selected files into AI Chat for analysis, editing instructions, and updated output generation.

## Technical notes
### Existing endpoints still used
- `/transcribe`
- `/transcript/enrich`
- `/dictate`
- `/chat`
- `/live/chat`
- `/live/reset`
- `/knowledge/list`
- `/knowledge/search`
- `/knowledge/ask`

### Additional behavior now supported
- `/chat` accepts chat mode and browser-provided history context.
- URL and search context can be injected into AI Chat responses.
- Spreadsheet extraction is supported for `.xlsx` / `.xls` files.
- Temporary live voice sessions are cleaned up automatically server-side.

## Deployment notes
- `api/main.py` now resolves `index.html` correctly from the project root during single-service deployments.
- Docker runs the FastAPI app from `api.main:app`.
- Added helper dependencies for spreadsheets, HTML parsing, and YouTube transcript ingestion.
