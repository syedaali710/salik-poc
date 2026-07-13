"""
Report-writer LLM pass — the second Claude call behind the "Download report" button.

The chat endpoints answer one question at a time for the UI. This module does
something different: it hands Claude the WHOLE dashboard dataset (plus the topics
the user explored in chat) and asks for a cohesive, board-grade deck — an
executive summary, a KPI overview, variance/bridge charts, a KPI table and a
drivers & risks section — modelled on the client's Board of Directors report.

Claude returns a list of "blocks" in the exact schema pptx_builder already
renders (narrative / chart / table), so no new rendering code is needed:
multi-KPI Actual/Budget/PY charts become KPI cards, and bridge specs render as
signed variance bars. If anything fails we return None so the caller can fall
back to the chat-accumulated blocks.
"""
from __future__ import annotations

import json

import requests

from chat import (
    ANTHROPIC_URL,
    ANTHROPIC_VERSION,
    MODEL,
    _api_key,
    _extract_json,
    _load_dataset,
    _valid_table,
)
from planner import _valid_spec

REPORT_SYSTEM_PROMPT = """You are the report architect for SALIC (Saudi Agricultural and Livestock
Investment Company). Produce a professional, Board-of-Directors-grade financial deck as
STRUCTURED JSON, using ONLY the JSON dataset provided below — never invent, extrapolate or
guess a number. If a figure isn't in the dataset, omit it rather than fabricate it.

GOAL: a cohesive multi-slide report (like a real board pack), NOT a single answer. Cover the
group's performance in a logical narrative arc. If the user explored specific topics (given as
FOCUS AREAS), weight the deck toward those, but always include the executive framing.

Build the deck as an ordered list of "blocks". Target a 4-page BOARD deck matching the
client's SALIC BoD Financial Performance layout (13.33×7.5 in, compact KPI tiles,
dense tables, side-by-side bridges). Produce exactly these blocks in order:
1. NARRATIVE — "Executive Summary": 4-5 crisp bullet sentences (board memo tone).
2. CHART — "KPI Overview": column chart with Actual/Budget/Prior Year → KPI cards.
   insight = one sentence (max 140 chars).
3. TABLE — "KPI Summary vs Budget and Prior Year": compact 7-row KPI table with
   variance % columns (signed, coloured).
4. CHART(s) — 1-3 waterfall bridges (variant:"waterfall"), most decision-relevant.
   One pointed insight per bridge (max 140 chars).
5. NARRATIVE — "Key Drivers & Risks": Wins:/Watch: prefixed bullets (3-4 each).
6. TABLE — "Matters for the Board": Issue/Recommendation/Owner, 2-4 rows.

For FOCUSED single questions, still include blocks 1-2 and the relevant bridge(s),
then 5-6 — skip unrelated bridges and optionally skip block 3.

NUMBER RULES for charts: every value in "values" is a STRING of digits only (optional leading
"-" and decimal point), no separators, no units. Convert raw dataset figures to millions
(divide by 1,000,000) and set "unit":"SAR m" — EXCEPT ratio KPIs (Net Debt/EBITDA "x", ROIC
"%") which keep native units and are NOT divided. Keep tables readable: format numbers with
thousands separators and units as needed in the STRING cells.

CHART SCHEMA (exact):
 {"type":"column|bar|stacked_column|line|area|pie|donut","title":"...","insight":"...",
  "unit":"SAR m","categories":[...],"series":[{"name":"...","values":["123","-45",...]}]}
 Bridge/waterfall charts additionally include "variant":"waterfall".
TABLE SCHEMA (exact): {"title":"...","columns":[...],"rows":[[...], ...]}
NARRATIVE SCHEMA: {"title":"...","text":"line1\\nline2\\n..."}

Respond with JSON ONLY (no markdown fences, no commentary), exactly this shape:
{"title":"...","period":"...","blocks":[
  {"type":"narrative","title":"...","text":"..."},
  {"type":"chart","spec":{...}},
  {"type":"table","spec":{...}},
  ...
]}

INSIGHT RULE: chart/table "insight" fields must be ONE crisp sentence, max 140 characters.
No trailing clauses that will wrap past the slide margin."""

FOCUSED_PROMPT = """
FOCUSED MODE — the user asked ONE specific analytical question (see FOCUS AREAS).
Produce a SHORT deck (4-7 blocks), not the full generic arc:
1. NARRATIVE — brief executive summary (3-4 bullets) framing the answer.
2. CHART — KPI Overview cards (keep this one slide for context).
3. CHART — the bridge(s) most relevant to the focus question (e.g. geography → EBITDA
   by Geography waterfall; portfolio → Net Income by Portfolio Company). Usually 1-2 bridges.
4. NARRATIVE — Key Drivers & Risks (Wins:/Watch: prefixes).
5. TABLE — Matters for the Board (2-3 rows tied to the focus topic).
Skip unrelated bridges, skip the full KPI Summary table unless the question is broad.
Keep every chart insight to one sentence (max 140 chars)."""


def _is_focused_question(questions):
    """Single narrow analytical question → shorter, topic-weighted deck."""
    if len(questions) != 1:
        return False
    q = questions[0].lower()
    markers = (
        "bridge", "drove", "what drove", "by geography", "by segment",
        "by portfolio", "waterfall", "variance", "driver", "drivers",
        "what caused", "why did", "explain the",
    )
    return any(m in q for m in markers)


def _trim_insight(text, limit=140):
    s = str(text or "").strip()
    if len(s) <= limit:
        return s
    cut = s[:limit].rsplit(" ", 1)[0]
    return (cut or s[:limit]).rstrip(",;:") + "…"


def _validate_blocks(raw_blocks):
    """Keep only well-formed blocks, coercing chart/table specs through the same
    validators the chat + pptx pipeline already trust."""
    blocks = []
    for b in raw_blocks or []:
        if not isinstance(b, dict):
            continue
        btype = b.get("type")
        if btype == "narrative":
            text = str(b.get("text") or "").strip()
            if text:
                blocks.append({"type": "narrative",
                               "title": str(b.get("title") or "Insight")[:120],
                               "text": text})
        elif btype == "chart":
            spec = _valid_spec(b.get("spec"))
            if spec:
                if spec.get("insight"):
                    spec["insight"] = _trim_insight(spec["insight"])
                blocks.append({"type": "chart", "spec": spec})
        elif btype == "table":
            spec = _valid_table(b.get("spec"))
            if spec:
                if spec.get("insight"):
                    spec["insight"] = _trim_insight(spec["insight"])
                blocks.append({"type": "table", "spec": spec})
    return blocks


def generate_report(focus_questions=None, timeout=120):
    """Full dataset (+ optional user focus areas) -> professional report.

    Returns {"title", "period", "blocks":[...]} on success, or None on any
    failure so the caller can fall back to the chat-accumulated blocks."""
    key = _api_key()
    if not key:
        print("[report] no ANTHROPIC_API_KEY; falling back to chat blocks")
        return None

    focus = [str(q).strip() for q in (focus_questions or []) if str(q).strip()][:12]
    system = REPORT_SYSTEM_PROMPT
    if _is_focused_question(focus):
        system += FOCUSED_PROMPT
    user_content = "Generate the full SALIC board report deck from the dataset."
    if focus:
        user_content += "\n\nFOCUS AREAS (topics the user explored — weight the deck toward these):\n"
        user_content += "\n".join(f"- {q}" for q in focus)

    payload = {
        "model": MODEL,
        "max_tokens": 8000,
        "system": system + "\n\nDATASET:\n" + json.dumps(_load_dataset()),
        "messages": [{"role": "user", "content": user_content}],
    }

    try:
        r = requests.post(
            ANTHROPIC_URL,
            headers={
                "x-api-key": key,
                "anthropic-version": ANTHROPIC_VERSION,
                "content-type": "application/json",
            },
            json=payload,
            timeout=timeout,
        )
        r.raise_for_status()
        body = r.json()
        if body.get("stop_reason") == "max_tokens":
            print("[report] response truncated (stop_reason=max_tokens)")
        content = body.get("content") or []
        text = next((b["text"] for b in content if b.get("type") == "text"), "")
    except Exception as e:
        print(f"[report] request failed: {e}")
        return None

    try:
        parsed = _extract_json(text)
    except (ValueError, json.JSONDecodeError) as e:
        print(f"[report] JSON parse failed ({e})")
        return None

    blocks = _validate_blocks(parsed.get("blocks"))
    if not blocks:
        print("[report] no valid blocks produced")
        return None

    return {
        "title": str(parsed.get("title") or "").strip(),
        "period": str(parsed.get("period") or "").strip(),
        "blocks": blocks,
    }
