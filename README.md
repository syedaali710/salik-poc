# SALIC Voice → Slide (local POC)

Speak a finance update → it is **transcribed on your own machine** by Whisper →
mapped onto the **SALIC template** as an editable **PowerPoint** slide with a KPI table.
No cloud, no API keys, no data leaves your computer.

---

## Run it (macOS — simplest path)

1. Make sure **Python 3** is installed. (Check: open Terminal, type `python3 --version`.
   If missing, install from https://www.python.org/downloads/ or run `xcode-select --install`.)
2. Double-click **`START_HERE.command`**.
   - macOS may say *"unidentified developer"* the first time → right-click the file →
     **Open** → **Open**.
3. The first run installs things and downloads the Whisper model (a few minutes, one time).
4. Your browser opens at **http://localhost:8000**. Click the mic, allow the microphone,
   speak your update, press stop. Edit any field, then **Download PowerPoint**.
5. To stop: close the black Terminal window (or press `Ctrl+C` in it).

## Run it (any OS — manual)

```bash
cd salic-voice-poc
python3 -m venv .venv
source .venv/bin/activate            # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python -m uvicorn app:app --port 8000
# open http://localhost:8000
```

---

## What each file does

| File | Role |
|------|------|
| `app.py` | Backend server. Endpoints: `/transcribe` (mic→Whisper), `/parse`, `/generate` (→.pptx) |
| `parsing.py` | Turns the transcript into slide fields (company, period, KPIs, commentary) |
| `pptx_builder.py` | Builds the SALIC-branded slide with python-pptx |
| `static/index.html` | The web page: mic recorder, editable fields, live slide preview |
| `START_HERE.command` | One-click launcher for macOS |

## Tips for good results
- Speak numbers as **digits**: say "two fifty" as **"250"**, "one point one eight billion" as **"1,180 million"**.
- Mention the **company name** and the **period** ("NADEC", "Q1 2025").
- The parser is deliberately simple — **every field is editable** before you download.
- Change model size for speed/accuracy: set `WHISPER_MODEL=tiny` (fast) or `medium` (accurate)
  before launching, e.g. `WHISPER_MODEL=base python -m uvicorn app:app --port 8000`.

---

## Moving to production (Azure)
This POC is built so production is a small swap, not a rewrite:

1. **Transcription** — replace the body of `transcribe_audio()` in `app.py` with a call to
   **Azure Whisper** (Azure OpenAI) in your tenant. Instructions are in the code comment.
2. **Numbers** — instead of parsing figures from speech, pull the verified KPIs from
   **Microsoft Fabric / Power BI** (read-only). Speech/AI then only writes the narrative.
3. **Hosting & security** — run inside SALIC's Azure (private networking, Entra ID sign-in,
   audit logs). Nothing else in the app changes.

*Prototype for demonstration. Numbers spoken in the demo are illustrative, not real financials.*
