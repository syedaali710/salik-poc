# Next.js Frontend Plan — SALIC AI Insights

Plan to replace `static/report.html` (~694 lines) with **Next.js 15 App Router**, while keeping the existing **FastAPI** backend and porting current avatar behavior (preconnect, auto-greet, reconnect, `keepAlive`).

---

## 1. What You're Migrating

| Feature | Today (`report.html`) | Backend |
|---------|----------------------|---------|
| Page shell | Single HTML file | FastAPI serves at `/` |
| Chat + SSE | `ask()` + `readSse()` | `POST /chat/stream` |
| Mic | `MediaRecorder` → transcribe | `POST /transcribe` (ElevenLabs) |
| Avatar | `@heygen/liveavatar-web-sdk` via **esm.sh** | `POST /heygen_token` |
| Report | In-memory blocks → PPTX | `POST /export_report` |
| Styling | Inline CSS, SALIC tokens | — |

**Keep in FastAPI:** `chat.py`, `elevenlabs_stt.py`, `heygen.py`, `pptx_builder.py`, `data/dashboard_data.json`, API keys.

**Retire after cutover:** `static/report.html`. Legacy `static/index.html` (Voice→Slide) stays out of scope unless requested.

---

## 2. Target Architecture

```
┌─────────────────────────────────────────────────────────────┐
│  Next.js 15 (frontend/)                                      │
│  ┌─────────────┐  ┌──────────────────┐  ┌─────────────────┐ │
│  │ ChatPanel   │  │ LiveAvatarPlayer │  │ ReportSidebar   │ │
│  └──────┬──────┘  └────────┬─────────┘  └────────┬────────┘ │
└─────────┼──────────────────┼─────────────────────┼──────────┘
          │ SSE              │ npm SDK + token     │ JSON blob
          ▼                  ▼                   ▼
┌─────────────────────────────────────────────────────────────┐
│  FastAPI (:8000)                                             │
│  /transcribe  /chat/stream  /heygen_token  /export_report   │
└─────────────────────────────────────────────────────────────┘
          │                  │                   │
          ▼                  ▼                   ▼
     ElevenLabs          LiveAvatar           pptx_builder
     Claude + data        api.liveavatar.com
```

**Monorepo layout:**

```
salik-poc/
├── app.py, chat.py, heygen.py, ...   # backend (unchanged for now)
├── data/
├── assets/
├── docs/
│   └── nextjs-frontend-plan.md       # this file
├── frontend/                          # new Next.js app
│   ├── app/
│   ├── components/
│   ├── hooks/
│   ├── lib/
│   └── types/
├── static/report.html                 # retired after cutover
└── docker-compose.yml                 # optional: next + fastapi
```

---

## 3. Tech Stack

| Layer | Choice | Rationale |
|-------|--------|-----------|
| Framework | **Next.js 15** App Router | SSR shell + client islands |
| Language | **TypeScript** | Mirrors `schemas.py` |
| Styling | **Tailwind CSS** | Port SALIC CSS variables |
| State | **Zustand** | Chat history, report blocks, mic/avatar status |
| Avatar SDK | **`npm install @heygen/liveavatar-web-sdk`** | Bundled by Next — no `esm.sh` |
| API | `fetch` + SSE reader | Same contracts as today |
| Charts | Custom mini-bars (port from HTML) | Match current previews |

Optional: **shadcn/ui** for inputs/buttons — not required for POC parity.

---

## 4. Routes & Layout

| Route | Type | Purpose |
|-------|------|---------|
| `/` | Server + client | Main insights page (replaces `/report`) |
| `app/icon.png` | Static | Fix favicon 404 |

**No Next API routes for chat/avatar** — browser talks directly to FastAPI (simpler SSE). Optional rewrites in dev only.

---

## 5. Component Map

```
frontend/
├── app/
│   ├── layout.tsx              # fonts, SALIC header, globals.css tokens
│   ├── page.tsx                # 2-column grid (chat | report)
│   └── icon.png
├── components/
│   ├── layout/
│   │   └── AppHeader.tsx
│   ├── chat/
│   │   ├── AskPanel.tsx        # mic + input + Ask button
│   │   ├── ChatTranscript.tsx
│   │   ├── StreamingAnswer.tsx
│   │   ├── SuggestionChips.tsx
│   │   ├── MiniChart.tsx
│   │   └── MiniTable.tsx
│   ├── avatar/
│   │   └── LiveAvatarPlayer.tsx   # 'use client' — video, canvas, chroma-key
│   ├── report/
│   │   ├── ReportSidebar.tsx
│   │   ├── ReportBlockCard.tsx
│   │   └── DownloadReportButton.tsx
│   └── mic/
│       └── MicButton.tsx
├── hooks/
│   ├── useChatStream.ts        # SSE + history (last 6 turns)
│   ├── useTranscribe.ts        # MediaRecorder → /transcribe → auto-ask
│   └── useLiveAvatar.ts        # port ALL current avatar logic
├── lib/
│   ├── api.ts                  # base URL, fetch helpers
│   ├── sse.ts                  # readSse port
│   └── types.ts                # from schemas.py
└── store/
    └── useReportStore.ts       # blocks, company, period
```

---

## 6. `useLiveAvatar` — Port Existing Behavior

Move logic from `static/report.html` into one hook. Do not re-simplify.

| Behavior | Port from current HTML |
|----------|------------------------|
| Token prefetch on load | `prefetchToken()` |
| Auto-connect on mount | `ensureSession()` |
| Greeting after `start()` | `playGreetingOnce()` — **not** on `SESSION_STREAM_READY` |
| `keepAlive()` every 60s | `startKeepAlive()` / `stopKeepAlive()` |
| Reconnect on disconnect | `scheduleReconnect()` + token prefetch |
| Speak queue | `speakChain` sequential `repeat()` |
| Chroma-key canvas loop | `startChromaKeyLoop()` |
| Audio unlock fallback | `enableAudioUnlock()` |
| Dead session retry | `isSessionDeadError` + `forceNew` |

**Loading the SDK in Next.js:**

```tsx
// components/avatar/LiveAvatarPlayer.tsx
'use client';

import { LiveAvatarSession, SessionEvent } from '@heygen/liveavatar-web-sdk';
```

**Lazy-load (optional)** to shrink initial bundle:

```tsx
const LiveAvatarPlayer = dynamic(
  () => import('./LiveAvatarPlayer'),
  { ssr: false, loading: () => <AvatarSkeleton /> }
);
```

**Important:** Avatar code must be `'use client'` only — never import SDK in Server Components.

---

## 7. API Contract (Unchanged)

Types aligned with `schemas.py`:

```typescript
// lib/types.ts
export type ChatRequest = { question: string; history: HistoryTurn[] };

export type ChatDoneEvent = {
  type: 'done';
  answer: string;
  chart: ChartSpec | null;
  table: TableSpec | null;
  ok: boolean;
};

export type ReportBlock =
  | { type: 'narrative'; title?: string; text: string }
  | { type: 'chart'; spec: ChartSpec }
  | { type: 'table'; spec: TableSpec };
```

| Endpoint | Client usage |
|----------|--------------|
| `POST /transcribe` | Mic `FormData` field `audio` |
| `POST /chat/stream` | SSE: `delta` events → `done` |
| `POST /heygen_token` | Prefetch on load + on disconnect |
| `POST /export_report` | Blob download |

**Environment variables:**

```env
# frontend/.env.local
NEXT_PUBLIC_API_URL=http://localhost:8000

# FastAPI / Render
FRONTEND_URL=http://localhost:3000
ANTHROPIC_API_KEY=...
LIVEAVATAR_API_KEY=...
ELEVENLABS_API_KEY=...
```

**FastAPI change (required when origins differ):**

```python
from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=[os.environ.get("FRONTEND_URL", "http://localhost:3000")],
    allow_methods=["*"],
    allow_headers=["*"],
)
```

---

## 8. Chat Flow

```
User types or uses mic
        │
        ├─ mic ──► POST /transcribe ──► fill input ──► auto-ask
        │
        └─ ask ──► POST /chat/stream (question + history)
                        │
                        ├─ SSE delta (streaming text)
                        └─ SSE done (answer, chart, table, ok)
                                │
                                ├─ add to transcript + report store
                                └─ avatar.speak(answer) if isUsableAnswer()
```

**Guards to preserve:**

- `isUsableAnswer()` — don't speak errors or add bad answers to report
- Don't speak via avatar on failed/stream errors
- History capped at last 6 turns

---

## 9. Deployment Options

| Option | Setup | Best for |
|--------|-------|----------|
| **A. Two Render services** | Next on `:3000`, FastAPI on `:8000` + CORS | Production POC |
| **B. Next rewrites (dev)** | Proxy `/api/*` → FastAPI | Local dev, single origin |
| **C. FastAPI redirect** | `/` → Next URL in prod | Clean cutover |

**`next.config.ts` (local dev):**

```typescript
async rewrites() {
  return process.env.NODE_ENV === 'development'
    ? [{ source: '/api/:path*', destination: 'http://localhost:8000/:path*' }]
    : [];
}
```

Then set `NEXT_PUBLIC_API_URL=""` in dev and call `/api/chat/stream`, etc.

---

## 10. Migration Phases

### Phase 0 — Backend prep (0.5 day)

- [ ] Add CORS to `app.py`
- [ ] Document `FRONTEND_URL`, `NEXT_PUBLIC_API_URL` in README
- [ ] Add favicon (`app/icon.png` in Next)

### Phase 1 — Scaffold (1 day)

- [ ] `npx create-next-app@latest frontend --ts --tailwind --app --eslint`
- [ ] `npm install @heygen/liveavatar-web-sdk zustand`
- [ ] Port SALIC tokens to `globals.css`
- [ ] `AppHeader` + 2-column layout
- [ ] `lib/api.ts`, `lib/types.ts`

### Phase 2 — Chat (2–3 days)

- [ ] `useChatStream` + `readSse`
- [ ] `ChatTranscript`, streaming cursor
- [ ] `SuggestionChips` (4 default prompts)
- [ ] `MiniChart` / `MiniTable`
- [ ] Report auto-add on successful answer

### Phase 3 — Mic (1 day)

- [ ] `MicButton` + `useTranscribe`
- [ ] Auto-ask after transcript
- [ ] Error states (401, empty transcript, network)

### Phase 4 — Avatar (2–3 days) — highest risk

- [ ] `useLiveAvatar` — full port from `report.html`
- [ ] `LiveAvatarPlayer` with chroma-key
- [ ] npm SDK (no esm.sh)
- [ ] `speak(answer)` wired from chat `done` event
- [ ] Test disconnect → reconnect while user asks

### Phase 5 — Report + deploy (1–2 days)

- [ ] `ReportSidebar`, company/period fields, PPTX download
- [ ] Render: two services or Docker Compose
- [ ] Remove FastAPI `report_page()` or redirect to Next
- [ ] Retire `static/report.html`

**Total estimate:** ~8–12 dev days for full parity.

**Recommended build order:** Phase 0 → 1 → 2 → 4 → 3 → 5 (chat before avatar validates API; mic is independent).

---

## 11. Next.js vs Current `report.html`

| Area | `report.html` + esm.sh | Next.js |
|------|------------------------|---------|
| SDK load | Third-party CDN on cold cache | Bundled, cached `_next/static/chunks` |
| Code structure | 694-line monolith | Components + hooks |
| Types | None | TypeScript from `schemas.py` |
| Deploy | HTML served by Python | Dedicated frontend service |
| Avatar connect/TTS latency | LiveAvatar + WebRTC | **Same** — not fixed by framework |

Next.js improves maintainability, SDK caching, and DX. It does **not** remove LiveAvatar session or TTS delay.

---

## 12. Out of Scope (Unless Requested)

- Moving Claude/STT/PPTX logic to Next API routes
- Auth (Entra ID) — add via middleware later
- Legacy Voice→Slide (`static/index.html`, `parsing.py`, `planner.py`)
- LITE mode / `repeatAudio()` with custom TTS pipeline

---

## 13. Success Criteria

- [ ] Feature parity: stream chat, mic, avatar (preconnect, greet, reconnect, keepAlive), report export
- [ ] SALIC visual match (header, colors, 2-column layout, mobile single column)
- [ ] SDK via `npm`, not `esm.sh`
- [ ] FastAPI serves JSON only; Next serves UI
- [ ] No regression: spelling tolerance, error messages, avatar reconnect
- [ ] Documented env vars for both apps

---

## 14. Getting Started

```bash
cd salik-poc
npx create-next-app@latest frontend --typescript --tailwind --eslint --app --no-src-dir
cd frontend
npm install @heygen/liveavatar-web-sdk zustand
```

Copy `.env.local`:

```env
NEXT_PUBLIC_API_URL=http://localhost:8000
```

Run locally (two terminals):

```bash
# Terminal 1 — backend
uv run uvicorn app:app --reload --port 8000

# Terminal 2 — frontend
cd frontend && npm run dev
```

---

## 15. Reference — Current Backend Routes

| Method | Path | Purpose |
|--------|------|---------|
| `GET` | `/`, `/report` | Serves `report.html` today (retire after cutover) |
| `POST` | `/transcribe` | ElevenLabs STT |
| `POST` | `/chat` | Non-streaming chat |
| `POST` | `/chat/stream` | SSE streaming chat |
| `POST` | `/heygen_token` | LiveAvatar session token |
| `POST` | `/export_report` | SALIC PPTX download |
