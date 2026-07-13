# Daily Work Report — SALIC AI Insights POC

| Field | Detail |
|-------|--------|
| **Date** | Monday, 13 July 2026 |
| **Project** | SALIC AI Insights — Proof of Concept (POC) |
| **Repositories** | Backend: `salik-poc` · Frontend: `salic-frontend` |
| **Live demo** | Frontend: [https://salic-ai.vercel.app/](https://salic-ai.vercel.app/) |
| **Report type** | End-of-day progress summary (Frontend + Backend) |

---

## 1. Executive Summary

Today's work advanced the SALIC AI Insights POC from a single-page HTML prototype to a **production-style split architecture**: a **Next.js frontend** deployed on Vercel and a **FastAPI backend** prepared for Render. The main deliverable was aligning generated PowerPoint reports with the **client Board of Directors (BoD) financial deck** standard, including a second AI pass for professional report composition, layout fixes, deployment wiring, and operational hardening.

---

## 2. Backend Work (`salik-poc`)

### 2.1 Board-format PowerPoint export

- Added **`report_writer.py`** — a second Claude LLM pass that reads the full `dashboard_data.json` dataset plus user focus questions and produces a structured multi-slide deck (executive summary, KPI overview, KPI table, waterfall bridges, drivers/risks, matters for the Board).
- Rebuilt **`pptx_builder.py`** to render reports in the **client BoD format**:
  - Standard **16:9 slide size** (13.33 × 7.5 in), matching the client reference deck.
  - Navy masthead with SALIC / PIF logos, compact KPI tiles, dense tables, side-by-side bridges, and wins/watch-point panels.
  - **4-page board layout** instead of 8+ sparse oversized slides.
- Added board assets: `assets/SALIC_board_template.pptx`, `assets/board_logo_left.png`, `assets/board_logo_right.png`.
- Extended **`/export_report`** in `app.py` with `generate: true` and `questions[]` to trigger the professional report path; falls back to chat blocks if generation fails.
- Updated **`schemas.py`** and **`planner.py`** to support waterfall chart variants and export request fields.

### 2.2 Report quality & layout fixes

- Removed template slide-number artefacts that caused duplicate page numbers.
- Fixed truncated insight text on bridge slides (taller text boxes, adaptive font sizing).
- Improved waterfall chart readability (white labels on navy total bars, client colour palette).
- Added **focused report mode** for single analytical questions (e.g. “What drove the EBITDA bridge by geography?”) to produce shorter, topic-relevant decks.

### 2.3 API reliability & operations

- Improved **Anthropic API error handling** in `chat.py` — surfaces insufficient-credit/billing errors clearly instead of generic failure messages.
- Documented split-deploy environment variables (`FRONTEND_URL` on backend; `BACKEND_URL` / `NEXT_PUBLIC_API_URL` on frontend).
- Updated **`.gitignore`** to exclude generated reports, client reference PDFs/PPTX at repo root, and Office lock files.
- Committed backend changes (commit `1ed89d7`): *Add board-format PPTX export with a second LLM report pass.*

### 2.4 Documentation

- Maintained **`docs/nextjs-frontend-plan.md`** — architecture, API contract, deployment options, and migration phases for the Next.js frontend.

---

## 3. Frontend Work (`salic-frontend`)

### 3.1 Next.js application (replaces legacy `static/report.html`)

Built a full **App Router** frontend with:

| Area | Components / modules |
|------|----------------------|
| **Chat** | SSE streaming (`useChatStream`, `ChatPanel`, `StreamingAnswer`, `MiniChart`, `MiniTable`) |
| **Voice** | Mic recording → ElevenLabs STT → auto-ask (`MicButton`, `useTranscribe`) |
| **Avatar** | LiveAvatar integration with preconnect, greet, reconnect, chroma-key (`LiveAvatarPanel`, `AvatarProvider`) |
| **Report builder** | Sidebar accumulates Q&A blocks; download triggers board-grade PPTX (`ReportSidebar`, `useReportStore`) |

### 3.2 Production API integration

- **`app/api/chat/stream/route.ts`** — server-side SSE proxy to FastAPI (avoids dev rewrite timeouts).
- **`app/api/export_report/route.ts`** — long-running report proxy (`maxDuration: 300`) for the second LLM + PPTX build (~30–50 s).
- **`lib/api.ts`** — configurable `NEXT_PUBLIC_API_URL` for direct backend calls (transcribe, avatar token).
- **`next.config.ts`** — dev-only rewrites to `localhost:8000`; production uses env-based backend URL.

### 3.3 Report download flow

- Extended types and store to pass **`generate: true`** and collected **`questions[]`** from successful chat answers to `/export_report`.
- **Download Report** button enabled when report items exist; shows composing status during generation.

### 3.4 Deployment

- **Deployed to Vercel:** [https://salic-ai.vercel.app/](https://salic-ai.vercel.app/)
- Documented production env setup in frontend `README.md` and `.env.example`.

---

## 4. End-to-end POC flow (current state)

```
User (Vercel frontend)
    → Chat: POST /api/chat/stream → FastAPI /chat/stream (Claude + dashboard data)
    → Mic:  POST /transcribe → ElevenLabs STT
    → Avatar: POST /heygen_token → LiveAvatar session
    → Report: POST /api/export_report → report_writer (2nd LLM) → pptx_builder (BoD deck) → .pptx download
```

---

## 5. Issues identified & mitigations

| Issue | Impact | Action taken / required |
|-------|--------|-------------------------|
| Anthropic API credit exhaustion | Chat/stream returns HTTP 400 | Clear user-facing error message; **billing top-up required** on Anthropic account |
| Split deploy CORS | Browser blocked cross-origin calls | Backend `FRONTEND_URL=https://salic-ai.vercel.app`; frontend `NEXT_PUBLIC_API_URL` = backend URL |
| Legacy 26.67×15 in template | Reports looked unprofessional vs client BoD deck | New 16:9 board builder aligned to client reference |
| Vercel serverless timeout on report | Long LLM pass may fail on Hobby plan | Route handler `maxDuration: 300`; Pro plan recommended for report download |

---

## 6. Deliverables completed today

1. Client-aligned **4-page board PPTX** generation pipeline (backend).
2. **Second LLM report composition** pass (`report_writer.py`).
3. **Next.js frontend** with chat, mic, avatar, and report builder.
4. **Vercel production deployment** of the frontend.
5. **Split-deploy documentation** and environment variable mapping.
6. **Git hygiene** — `.gitignore` for generated artefacts; backend commit on branch `change/frontent-ui`.

---

## 7. Pending / next steps

- Ensure **backend is deployed** to a public URL (e.g. Render) and Vercel env vars (`BACKEND_URL`, `NEXT_PUBLIC_API_URL`) point to it.
- Set **`FRONTEND_URL=https://salic-ai.vercel.app`** on the backend service.
- **Restore Anthropic API credits** for chat and report generation in production.
- Optional: add Vercel route handlers for `/transcribe` and `/heygen_token` so all traffic proxies through Next.js.
- Future: enrich dataset (full P&L, portfolio table, reported→recurring bridge) to match full client BoD deck content.

---

## 8. Technology stack

| Layer | Technologies |
|-------|----------------|
| **Backend** | Python 3.11+, FastAPI, uv, Claude (Anthropic), python-pptx, ElevenLabs STT, LiveAvatar |
| **Frontend** | Next.js (App Router), TypeScript, React, Zustand, Tailwind CSS, Vercel |
| **Data** | `data/dashboard_data.json` (YTD Dec-2025 financial snapshot) |
| **Deploy targets** | Vercel (frontend), Render/Docker (backend) |

---

## 9. Time allocation (estimate)

| Activity | Approx. share |
|----------|----------------|
| Backend report engine & BoD layout | 45% |
| Frontend integration & API proxies | 25% |
| Deployment & environment configuration | 15% |
| Testing, bug fixes, client deck comparison | 15% |

---

*This report documents POC development work on 13 July 2026. Figures in the demo dataset are illustrative; the prototype is for demonstration purposes.*
