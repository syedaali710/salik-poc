"""
SALIC Voice -> Slide  |  local backend
-------------------------------------------------
Flow:  browser mic  ->  /transcribe (local Whisper)  ->  parse  ->  slide fields
       edit fields  ->  /generate  ->  SALIC-branded .pptx

Everything runs on THIS machine. No cloud, no API keys, no data leaves your computer.
To move to production you swap ONE function (transcribe_audio) to call Azure Whisper,
and feed the KPI numbers from Microsoft Fabric / Power BI instead of the transcript.
"""
import os
import tempfile

from fastapi import FastAPI, UploadFile, File, Body
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse

from parsing import parse
from planner import plan_charts
from pptx_builder import build_pptx

HERE = os.path.dirname(os.path.abspath(__file__))
app = FastAPI(title="SALIC Voice-to-Slide POC")

# Whisper model is loaded lazily on first transcription (first call downloads it, ~1 min).
_MODEL = None


def get_model():
    global _MODEL
    if _MODEL is None:
        from faster_whisper import WhisperModel
        size = os.environ.get("WHISPER_MODEL", "small")   # tiny/base/small/medium
        print(f"[whisper] loading '{size}' model (first run downloads it)…")
        _MODEL = WhisperModel(size, device="cpu", compute_type="int8")
        print("[whisper] ready.")
    return _MODEL


def transcribe_audio(path: str) -> str:
    """
    Local Whisper transcription.
    --- PRODUCTION SWAP -------------------------------------------------------
    Replace the body of this function with a call to Azure Whisper, e.g.:
        from openai import AzureOpenAI
        client = AzureOpenAI(azure_endpoint=..., api_key=..., api_version=...)
        with open(path, "rb") as f:
            r = client.audio.transcriptions.create(model="whisper", file=f)
        return r.text
    Nothing else in this app changes.
    ---------------------------------------------------------------------------
    """
    model = get_model()
    segments, _info = model.transcribe(path, language="en", vad_filter=True)
    return " ".join(s.text.strip() for s in segments).strip()


@app.get("/", response_class=HTMLResponse)
def index():
    with open(os.path.join(HERE, "static", "index.html"), encoding="utf-8") as f:
        return HTMLResponse(f.read(), headers={"Cache-Control": "no-store"})


@app.post("/transcribe")
async def transcribe(audio: UploadFile = File(...)):
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
    return JSONResponse({"transcript": text, "data": parse(text)})


@app.post("/parse")
async def parse_endpoint(payload: dict = Body(...)):
    return JSONResponse({"data": parse(payload.get("text", ""))})


@app.post("/plan")
async def plan_endpoint(payload: dict = Body(...)):
    """Transcript text -> AI chart plan (list of chart specs, possibly empty)."""
    return JSONResponse({"charts": plan_charts(payload.get("text", ""))})


@app.post("/generate")
async def generate(payload: dict = Body(...)):
    buf = build_pptx(payload)
    company = (payload.get("company") or "SALIC").strip().replace(" ", "_") or "SALIC"
    fname = f"{company}_Performance_Summary.pptx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
