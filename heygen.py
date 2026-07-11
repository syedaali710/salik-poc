"""
Mints short-lived LiveAvatar session tokens for the browser.

HeyGen's old Streaming Avatar API (v1 streaming.create_token) is sunset; this
now talks to LiveAvatar (api.liveavatar.com), HeyGen's current real-time
avatar product. The raw LIVEAVATAR_API_KEY never reaches the frontend: the
browser calls our /heygen_token endpoint, which calls LiveAvatar server-side
and hands back only the resulting session token. The frontend SDK
(@heygen/liveavatar-web-sdk) takes that token and handles connecting/starting
the session itself.
"""
import os
import re

import requests

HERE = os.path.dirname(os.path.abspath(__file__))
SESSION_TOKEN_URL = "https://api.liveavatar.com/v1/sessions/token"

# Graham in Black Suit (the wider, non-portrait framing, which shows natural
# hand gestures while talking rather than the tight head-and-shoulders crop
# of the "(Portrait)" variant) + its native voice, from the account's
# LiveAvatar public catalog (avatar/voice IDs are account-specific, not secrets).
AVATAR_ID = "03f8332d-9046-42a1-bff3-3b2309f77b58"
VOICE_ID = "e04e9d57-853f-4d72-a8ff-8e3c768f4c9c"


def _api_key():
    key = os.environ.get("LIVEAVATAR_API_KEY", "").strip()
    if key:
        return key
    env_path = os.path.join(HERE, ".env")
    if os.path.exists(env_path):
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                m = re.match(r"\s*LIVEAVATAR_API_KEY\s*=\s*(\S+)", line)
                if m:
                    return m.group(1).strip().strip('"').strip("'")
    return ""


def create_session_token(timeout=15):
    """-> LiveAvatar's session-token JSON on success, or {"data": None, "error": "..."}
    on failure. Always returns a JSON-serializable dict so the frontend never has to
    handle a raw 500."""
    key = _api_key()
    if not key:
        return {"data": None, "error": "LiveAvatar backend is not configured (missing LIVEAVATAR_API_KEY)."}
    try:
        r = requests.post(
            SESSION_TOKEN_URL,
            headers={"X-API-KEY": key, "Content-Type": "application/json"},
            json={
                "mode": "FULL",
                "avatar_id": AVATAR_ID,
                "avatar_persona": {"voice_id": VOICE_ID, "language": "en"},
            },
            timeout=timeout,
        )
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"[heygen] session token request failed: {e}")
        return {"data": None, "error": "Could not reach LiveAvatar just now."}
