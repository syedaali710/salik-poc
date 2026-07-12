# SALIC AI Insights

Ask about SALIC group financials (voice or text) → local Whisper transcription →
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

3. Open **http://localhost:8000**. Click the mic (or type), ask a question, then
   download the report from the sidebar.

### macOS one-click

Double-click **`START_HERE.command`** (right-click → Open the first time if macOS blocks it).

### Environment

| Variable | Required | Purpose |
|----------|----------|---------|
| `ANTHROPIC_API_KEY` | Yes (chat) | Claude Q&A |
| `LIVEAVATAR_API_KEY` | Optional | LiveAvatar speech |
| `WHISPER_MODEL` | No | `tiny` (default) / `base` / `small` / `medium` |
| `PORT` | No | Default `8000` |

Keys can also live in a local `.env` file (never commit it).

---

## API (FastAPI)

Interactive docs: **http://localhost:8000/docs**

| Method | Path | Body | Role |
|--------|------|------|------|
| `GET` | `/`, `/report` | — | Serves the chat UI |
| `POST` | `/transcribe` | `multipart/form-data` (`audio`) | Mic → Whisper transcript |
| `POST` | `/chat` | `{ question, history[] }` | Grounded answer + chart/table (one shot) |
| `POST` | `/chat/stream` | same | SSE: `delta` text chunks, then `done` with chart/table |
| `POST` | `/heygen_token` | — | LiveAvatar session token |
| `POST` | `/export_report` | `{ company, period, blocks[] }` | PPTX download |

Pydantic models live in `schemas.py`.

---

## What each file does

| File | Role |
|------|------|
| `app.py` | FastAPI app, lifespan Whisper preload, routers |
| `schemas.py` | Request/response models |
| `chat.py` | Claude Q&A over `dashboard_data.json` |
| `heygen.py` | LiveAvatar session token minting |
| `pptx_builder.py` | SALIC-template PowerPoint assembly |
| `planner.py` | Shared chart-spec validation (+ legacy Groq planner) |
| `static/report.html` | Chat + avatar + report UI |
| `data/dashboard_data.json` | YTD Dec-2025 Power BI snapshot |
| `pyproject.toml` / `uv.lock` | Dependencies (uv) |

---

## Deploy on Render

**Use Docker** (recommended). Native Python + `uv sync` on Render defaults to **Python 3.14**, which has no prebuilt `av` wheel for `faster-whisper` and the build fails without FFmpeg dev libraries.

### Render dashboard settings

1. **New → Web Service** → connect `syedaali710/salik-poc`, branch **`salikavatar`**
2. **Environment → Docker** (not Python)
3. **Dockerfile Path:** `./Dockerfile`
4. **Instance:** Free (or paid for more RAM — Whisper is heavy)
5. **Environment variables:**

| Key | Value |
|-----|--------|
| `ANTHROPIC_API_KEY` | your key |
| `LIVEAVATAR_API_KEY` | your key (optional) |
| `WHISPER_MODEL` | `tiny` |

6. **Do not** set a custom Build Command — Docker uses the Dockerfile
7. Deploy → open `https://<your-service>.onrender.com`

### If you must use Native Python (not recommended)

| Setting | Value |
|---------|--------|
| Runtime | Python |
| **Python Version** | **`3.11.14`** (required — not 3.14) |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `uvicorn app:app --host 0.0.0.0 --port $PORT` |
| Env `PYTHON_VERSION` | `3.11.14` |

Do **not** use `uv sync --frozen` on native Python unless you also install FFmpeg system packages (not available on Render native builds).

### Local Docker test

```bash
docker build -t salic-ai-insights .
docker run --rm -p 8000:8000 -e ANTHROPIC_API_KEY=… -e LIVEAVATAR_API_KEY=… salic-ai-insights
```

---

## Moving to production (Azure)

1. **Transcription** — swap local Whisper for Azure Whisper in your tenant.
2. **Numbers** — pull verified KPIs from Microsoft Fabric / Power BI (read-only);
   speech/AI only writes the narrative.
3. **Hosting** — private networking, Entra ID, audit logs.

*Prototype for demonstration. Figures in the demo dataset are illustrative.*
