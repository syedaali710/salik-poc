"""
AI chart planner.

Sends the transcript to an LLM (Groq, OpenAI-compatible API) and gets back a
structured *chart plan*: a list of chart specs the generic renderer in
pptx_builder.py can draw. The LLM never draws anything and is instructed to
use only numbers that appear in the text, so every figure stays checkable in
the UI before download.

Returns [] on any failure (no key, network down, bad JSON) so the app always
degrades gracefully to the standard deck.
"""
import json
import os
import re

import requests

HERE = os.path.dirname(os.path.abspath(__file__))
GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"
MODEL = os.environ.get("GROQ_MODEL", "llama-3.3-70b-versatile")

ALLOWED_TYPES = {"column", "bar", "stacked_column", "line", "area", "pie", "donut"}

SYSTEM_PROMPT = """You are a data-visualization planner for executive finance decks.
Given a spoken/written finance update, extract chartable datasets and decide the best chart for each.

Rules — accuracy first:
- Use ONLY numbers explicitly present in the text. Never invent, extrapolate or total numbers.
- Pick the chart type that fits the data: "line" for values over time (quarters, months, years);
  "column" for comparing a few items or series side by side; "bar" for ranked comparisons or
  variances (horizontal); "stacked_column" for composition over categories; "pie" or "donut"
  ONLY for parts-of-a-whole with 2-6 slices that genuinely sum to a meaningful total.
- Do NOT chart a time trend as pie/donut. Do NOT duplicate the same dataset in two charts.
- Skip datasets already fully covered by a simpler chart. 0-3 charts total; return zero charts
  if the text has no chartable structure beyond single standalone figures.
- Keep series in the same unit within one chart. Note the unit (e.g. "PKR m", "SAR m", "%").
- "insight" is one short executive takeaway (max 12 words) derived from the numbers.

Respond with JSON only, exactly this shape:
{"charts":[{"type":"line","title":"...","insight":"...","unit":"PKR m",
  "categories":["Q2 2025","Q3 2025"],
  "series":[{"name":"Revenue","values":["24500","25100"]}]}]}
CRITICAL: every value in "values" is a STRING containing only digits, an optional
leading minus and optional decimal point (e.g. "24500", "-15", "10.4") — never
thousands separators, units or words. Every series must have exactly one value
per category (use "" if truly missing)."""


def _api_key():
    key = os.environ.get("GROQ_API_KEY", "").strip()
    if key:
        return key
    env_path = os.path.join(HERE, ".env")
    if os.path.exists(env_path):
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                m = re.match(r"\s*GROQ_API_KEY\s*=\s*(\S+)", line)
                if m:
                    return m.group(1).strip().strip('"').strip("'")
    return ""


def _trim_words(text, limit):
    """Trim to <= limit chars on a word boundary so insights never cut mid-word."""
    text = (text or "").strip()
    if len(text) <= limit:
        return text
    cut = text[:limit].rsplit(" ", 1)[0].rstrip(" ,;:-–—")
    return (cut or text[:limit]) + "…"


def _valid_spec(c):
    if not isinstance(c, dict):
        return None
    ctype = str(c.get("type", "")).lower().strip()
    if ctype not in ALLOWED_TYPES:
        return None
    cats = [str(x) for x in (c.get("categories") or []) if str(x).strip()]
    series = []
    for s in c.get("series") or []:
        if not isinstance(s, dict):
            continue
        vals = s.get("values") or []
        clean = []
        for v in vals:
            try:
                sv = str(v).replace(",", "").strip()
                clean.append(float(sv) if sv else None)
            except (TypeError, ValueError):
                clean.append(None)
        if len(clean) == len(cats) and any(v is not None for v in clean):
            series.append({"name": str(s.get("name") or "Series"), "values": clean})
    if not cats or not series:
        return None
    if ctype in ("pie", "donut"):
        series = series[:1]
        if len(cats) > 6 or any(v is None or v < 0 for v in series[0]["values"]):
            return None
    spec = {
        "type": ctype,
        "title": str(c.get("title") or "Chart")[:120],
        "insight": _trim_words(str(c.get("insight") or ""), 220),
        "unit": str(c.get("unit") or "")[:20],
        "categories": cats,
        "series": series[:5],
    }
    # Optional rendering hint: bridge/waterfall charts (start total → deltas →
    # end total) render as floating up/down bars in pptx_builder.
    if str(c.get("variant") or "").lower().strip() == "waterfall":
        spec["variant"] = "waterfall"
    return spec


def plan_charts(text, timeout=30):
    """transcript text -> list of validated chart specs (possibly empty)."""
    text = (text or "").strip()
    key = _api_key()
    if not text or not key:
        return []
    try:
        r = requests.post(
            GROQ_URL,
            headers={"Authorization": f"Bearer {key}"},
            json={
                "model": MODEL,
                "temperature": 0.2,
                "response_format": {"type": "json_object"},
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": text},
                ],
            },
            timeout=timeout,
        )
        r.raise_for_status()
        content = r.json()["choices"][0]["message"]["content"]
        plan = json.loads(content)
    except Exception as e:
        print(f"[planner] chart planning failed: {e}")
        return []
    charts = []
    for c in (plan.get("charts") or [])[:3]:
        spec = _valid_spec(c)
        if spec:
            charts.append(spec)
    return charts
