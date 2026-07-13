# SALIC AI Insights

Ask about SALIC group financials (voice or text) → ElevenLabs Speech-to-Text →
Claude answers grounded in `data/dashboard_data.json` → optional chart/table →
LiveAvatar speaks the answer → download a SALIC-branded PowerPoint report.

Managed with **[uv](https://docs.astral.sh/uv/)** and served by **FastAPI**.

---

## Run it (any OS)

1. Install [uv](https://docs.astral.sh/uv/getting-started/installation/).
2. From this folder:

```bash
uv sync
uv run uvicorn app:app --host 127.0.0.1 --port 8000
```

3. Open **http://localhost:3000** (Next.js frontend). Legacy **http://localhost:8000** redirects to the frontend.

For the split frontend repo (`salic-frontend`), run both services — see that repo's README.

### Environment

| Variable | Required | Purpose |
|----------|----------|---------|
| `ANTHROPIC_API_KEY` | Yes (chat) | Claude Q&A |
| `ELEVENLABS_API_KEY` | Yes (mic) | [ElevenLabs Speech-to-Text](https://elevenlabs.io/docs/eleven-api/guides/cookbooks/speech-to-text) (`scribe_v2`) |
| `LIVEAVATAR_API_KEY` | Optional | LiveAvatar speech |
| `FRONTEND_URL` | Yes (split deploy) | Next.js origin for CORS + `/` redirect (default `http://localhost:3000`) |
| `ELEVENLABS_STT_MODEL` | No | Default `scribe_v2` |
| `ELEVENLABS_STT_LANGUAGE` | No | ISO language hint (e.g. `eng`); omit for auto-detect |
| `PORT` | No | Default `8000` |

Keys can also live in a local `.env` file (never commit it).

---

## API (FastAPI)

Interactive docs: **http://localhost:8000/docs**

| Method | Path | Body | Role |
|--------|------|------|------|
| `GET` | `/`, `/report` | — | Redirect to Next.js frontend (`FRONTEND_URL`) |
| `POST` | `/transcribe` | `multipart/form-data` (`audio`) | Mic → ElevenLabs STT transcript |
| `POST` | `/chat` | `{ question, history[] }` | Grounded answer + chart/table (one shot) |
| `POST` | `/chat/stream` | same | SSE: `delta` text chunks, then `done` with chart/table |
| `POST` | `/heygen_token` | — | LiveAvatar session token |
| `POST` | `/export_report` | `{ company, period, blocks[] }` | PPTX download |

Pydantic models live in `schemas.py`.

---

## What each file does

| File | Role |
|------|------|
| `app.py` | FastAPI app and routers |
| `elevenlabs_stt.py` | ElevenLabs Speech-to-Text (`/transcribe`) |
| `schemas.py` | Request/response models |
| `chat.py` | Claude Q&A over `dashboard_data.json` |
| `heygen.py` | LiveAvatar session token minting |
| `pptx_builder.py` | SALIC-template PowerPoint assembly |
| `planner.py` | Shared chart-spec validation (+ legacy Groq planner) |
| `static/report.html` | **Retired** — reference only; UI is `salic-frontend` |
| `data/dashboard_data.json` | YTD Dec-2025 Power BI snapshot |
| `pyproject.toml` / `uv.lock` | Dependencies (uv) |

---

## Deploy on Render

**Use Docker** (recommended). The app no longer needs local Whisper/ffmpeg — mic transcription calls ElevenLabs in the cloud.

### Render dashboard settings

1. **New → Web Service** → connect repo, branch **`salikavatar`**
2. **Environment → Docker**
3. **Dockerfile Path:** `./Dockerfile`
4. **Environment variables:**

| Key | Value |
|-----|--------|
| `ANTHROPIC_API_KEY` | your key |
| `ELEVENLABS_API_KEY` | your [ElevenLabs API key](https://elevenlabs.io/app/settings/api-keys) |
| `LIVEAVATAR_API_KEY` | your key (optional) |

5. Deploy → open `https://<your-service>.onrender.com`

### Native Python (also works now)

| Setting | Value |
|---------|--------|
| Runtime | Python |
| Python Version | `3.11.14` |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `uvicorn app:app --host 0.0.0.0 --port $PORT` |

No `faster-whisper`, no ffmpeg, no `uv sync` required on Render.

---

## Moving to production (Azure)

1. **Transcription** — swap local Whisper for Azure Whisper in your tenant.
2. **Numbers** — pull verified KPIs from Microsoft Fabric / Power BI (read-only);
   speech/AI only writes the narrative.
3. **Hosting** — private networking, Entra ID, audit logs.

*Prototype for demonstration. Figures in the demo dataset are illustrative.*
