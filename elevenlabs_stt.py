"""
ElevenLabs Speech-to-Text for mic transcription.

Replaces local faster-whisper. The browser uploads recorded audio to
POST /transcribe; this module forwards it to ElevenLabs Scribe and returns
the transcript text.

API docs: https://elevenlabs.io/docs/eleven-api/guides/cookbooks/speech-to-text
"""
import os
import re

import requests

HERE = os.path.dirname(os.path.abspath(__file__))
STT_URL = "https://api.elevenlabs.io/v1/speech-to-text"
MODEL = os.environ.get("ELEVENLABS_STT_MODEL", "scribe_v2")

_MIME = {
    ".webm": "audio/webm",
    ".wav": "audio/wav",
    ".mp3": "audio/mpeg",
    ".m4a": "audio/mp4",
    ".ogg": "audio/ogg",
    ".flac": "audio/flac",
}


def _api_key():
    key = os.environ.get("ELEVENLABS_API_KEY", "").strip()
    if key:
        return key
    env_path = os.path.join(HERE, ".env")
    if os.path.exists(env_path):
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                m = re.match(r"\s*ELEVENLABS_API_KEY\s*=\s*(\S+)", line)
                if m:
                    return m.group(1).strip().strip('"').strip("'")
    return ""


def _guess_mime(filename: str) -> str:
    ext = os.path.splitext(filename or "")[1].lower()
    return _MIME.get(ext, "application/octet-stream")


def _extract_text(body: dict) -> str:
    """Pull transcript string from ElevenLabs STT JSON response."""
    if not isinstance(body, dict):
        return ""
    text = body.get("text")
    if isinstance(text, str) and text.strip():
        return text.strip()
    # Multichannel responses nest transcripts per channel
    transcripts = body.get("transcripts")
    if isinstance(transcripts, list):
        parts = []
        for item in transcripts:
            if isinstance(item, dict):
                t = item.get("text")
                if isinstance(t, str) and t.strip():
                    parts.append(t.strip())
        if parts:
            return " ".join(parts)
    return ""


def transcribe_audio_bytes(data: bytes, filename: str = "audio.webm", timeout=90) -> str:
    """Upload audio bytes to ElevenLabs STT → transcript string (or \"\" on failure)."""
    if not data:
        return ""
    key = _api_key()
    if not key:
        print("[stt] ELEVENLABS_API_KEY not configured")
        return ""

    form = {
        "model_id": MODEL,
        "tag_audio_events": "false",
        "diarize": "false",
    }
    lang = os.environ.get("ELEVENLABS_STT_LANGUAGE", "").strip()
    if lang:
        form["language_code"] = lang

    try:
        r = requests.post(
            STT_URL,
            headers={"xi-api-key": key},
            files={"file": (filename, data, _guess_mime(filename))},
            data=form,
            timeout=timeout,
        )
        r.raise_for_status()
        text = _extract_text(r.json())
        if not text:
            print(f"[stt] empty transcript from ElevenLabs ({len(data)} bytes)")
        return text
    except Exception as e:
        print(f"[stt] ElevenLabs request failed: {e}")
        return ""
