"""
SALIC AI Insights — chat over the dashboard data, generate charts/tables/
narrative, export a SALIC-branded .pptx report.

Flow: browser chat -> type a question, or mic -> /transcribe (ElevenLabs STT)
      fills the question box -> /chat (Claude, grounded in
      data/dashboard_data.json) -> answer + optional chart/table -> "Report"
      panel -> /export_report -> SALIC-template .pptx download.
"""
from __future__ import annotations

import os
from typing import Any

from fastapi import APIRouter, FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import RedirectResponse, StreamingResponse

from chat import answer_question, stream_answer_question
from elevenlabs_stt import transcribe_audio_bytes
from heygen import create_session_token
from pptx_builder import build_report_pptx
from report_writer import generate_report
from schemas import (
    ChatRequest,
    ChatResponse,
    ExportReportRequest,
    TranscribeResponse,
)

HERE = os.path.dirname(os.path.abspath(__file__))

app = FastAPI(
    title="SALIC AI Insights",
    description="Grounded financial Q&A, LiveAvatar speech, and SALIC-branded PPTX export.",
    version="0.1.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[os.environ.get("FRONTEND_URL", "http://localhost:3000")],
    allow_methods=["*"],
    allow_headers=["*"],
)

api = APIRouter(tags=["api"])
pages = APIRouter(tags=["pages"])


FRONTEND_URL = os.environ.get("FRONTEND_URL", "http://localhost:3000").rstrip("/")


@pages.get("/")
@pages.get("/report")
def report_page():
    """Redirect legacy HTML UI to the Next.js frontend."""
    return RedirectResponse(url=FRONTEND_URL, status_code=307)


@api.post("/transcribe", response_model=TranscribeResponse)
async def transcribe(audio: UploadFile = File(...)) -> TranscribeResponse:
    """Mic audio → ElevenLabs Speech-to-Text transcript."""
    data = await audio.read()
    filename = audio.filename or "mic.webm"
    text = transcribe_audio_bytes(data, filename)
    return TranscribeResponse(transcript=text)


@api.post("/chat", response_model=ChatResponse)
async def chat_endpoint(payload: ChatRequest) -> ChatResponse:
    """Question (+ history) → narrative answer + optional chart/table (non-streaming)."""
    history = [turn.model_dump() for turn in payload.history]
    result = answer_question(payload.question, history)
    return ChatResponse.model_validate(result)


@api.post("/chat/stream")
async def chat_stream(payload: ChatRequest) -> StreamingResponse:
    """SSE stream: delta text chunks, then a final done event with chart/table."""
    import json as _json

    history = [turn.model_dump() for turn in payload.history]

    def event_gen():
        try:
            for event in stream_answer_question(payload.question, history):
                etype = event.get("type") or "message"
                yield f"event: {etype}\ndata: {_json.dumps(event, ensure_ascii=False)}\n\n"
        except Exception as e:
            print(f"[chat] stream endpoint failed: {e}")
            fail = {"type": "done", "answer": "Sorry, I couldn't reach the AI backend just now. Please try again.",
                    "chart": None, "table": None, "ok": False}
            yield f"event: done\ndata: {_json.dumps(fail, ensure_ascii=False)}\n\n"

    return StreamingResponse(
        event_gen(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache, no-transform",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


@api.post("/heygen_token")
async def heygen_token() -> dict[str, Any]:
    """Mint a short-lived LiveAvatar session token for the browser SDK."""
    return create_session_token()


@api.post("/export_report")
def export_report(payload: ExportReportRequest) -> StreamingResponse:
    """SALIC-template .pptx report download.

    When `generate` is set, a second LLM pass composes a professional board-style
    deck from the full dataset (focused on the user's questions); otherwise the
    chat-accumulated `blocks` are used. Falls back to `blocks` if generation fails.

    Declared sync (not async) on purpose: the report path makes a blocking LLM
    call that can take ~30-50s, so FastAPI runs it in a threadpool and keeps the
    event loop free (an async def here would freeze the server and reset proxies).
    """
    company = payload.company or "SALIC"
    period = payload.period or ""
    blocks = [b.model_dump(exclude_none=True) for b in payload.blocks]

    if payload.generate:
        report = generate_report(payload.questions)
        if report and report.get("blocks"):
            blocks = report["blocks"]
            if report.get("period"):
                period = report["period"]

    pptx_buf = build_report_pptx(blocks, company, period)
    fname = f"{company.strip().replace(' ', '_') or 'SALIC'}_AI_Insights_Report.pptx"
    return StreamingResponse(
        pptx_buf,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


app.include_router(pages)
app.include_router(api)


def main() -> None:
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))


if __name__ == "__main__":
    main()
