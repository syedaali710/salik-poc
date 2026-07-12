"""
SALIC AI Insights — chat over the dashboard data, generate charts/tables/
narrative, export a SALIC-branded .pptx report.

Flow: browser chat -> type a question, or mic -> /transcribe (local Whisper)
      fills the question box -> /chat (Claude, grounded in
      data/dashboard_data.json) -> answer + optional chart/table -> "Report"
      panel -> /export_report -> SALIC-template .pptx download.
"""
from __future__ import annotations

import os
import tempfile
from contextlib import asynccontextmanager
from typing import Any

from fastapi import APIRouter, FastAPI, File, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse

from chat import answer_question, stream_answer_question
from heygen import create_session_token
from pptx_builder import build_report_pptx
from schemas import (
    ChatRequest,
    ChatResponse,
    ExportReportRequest,
    TranscribeResponse,
)

HERE = os.path.dirname(os.path.abspath(__file__))

_MODEL = None


def get_model():
    global _MODEL
    if _MODEL is None:
        from faster_whisper import WhisperModel
        size = os.environ.get("WHISPER_MODEL", "tiny")  # tiny/base/small/medium
        print(f"[whisper] loading '{size}' model (first run downloads it)…")
        _MODEL = WhisperModel(size, device="cpu", compute_type="int8")
        print("[whisper] ready.")
    return _MODEL


def transcribe_audio(path: str) -> str:
    model = get_model()
    segments, _info = model.transcribe(path, language="en", vad_filter=True)
    return " ".join(s.text.strip() for s in segments).strip()


@asynccontextmanager
async def lifespan(_app: FastAPI):
    # Preload Whisper so the first mic click isn't stuck on model init/download.
    get_model()
    yield


app = FastAPI(
    title="SALIC AI Insights",
    description="Grounded financial Q&A, LiveAvatar speech, and SALIC-branded PPTX export.",
    version="0.1.0",
    lifespan=lifespan,
)

api = APIRouter(tags=["api"])
pages = APIRouter(tags=["pages"])


@pages.get("/", response_class=HTMLResponse)
@pages.get("/report", response_class=HTMLResponse)
def report_page() -> HTMLResponse:
    with open(os.path.join(HERE, "static", "report.html"), encoding="utf-8") as f:
        return HTMLResponse(f.read(), headers={"Cache-Control": "no-store"})


@api.post("/transcribe", response_model=TranscribeResponse)
async def transcribe(audio: UploadFile = File(...)) -> TranscribeResponse:
    """Mic audio → local Whisper transcript."""
    data = await audio.read()
    suffix = os.path.splitext(audio.filename or "")[1] or ".webm"
    tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    try:
        tmp.write(data)
        tmp.close()
        text = transcribe_audio(tmp.name)
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass
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
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


@api.post("/heygen_token")
async def heygen_token() -> dict[str, Any]:
    """Mint a short-lived LiveAvatar session token for the browser SDK."""
    return create_session_token()


@api.post("/export_report")
async def export_report(payload: ExportReportRequest) -> StreamingResponse:
    """Accumulated chat blocks → SALIC-template .pptx report download."""
    company = payload.company or "SALIC"
    period = payload.period or ""
    blocks = [b.model_dump(exclude_none=True) for b in payload.blocks]
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
