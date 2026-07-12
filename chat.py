"""
Chat over the SALIC dashboard dataset (data/dashboard_data.json), powered by
Claude (Anthropic Messages API).

Answers user questions using ONLY the extracted dashboard data and proposes a
chart and/or a table alongside an analytical narrative answer. Reuses
planner.py's chart-spec validator so chat charts render with the exact same
pptx machinery as the AI chart planner.
"""
import json
import os
import re

import requests

from planner import _valid_spec

HERE = os.path.dirname(os.path.abspath(__file__))
DATASET_PATH = os.path.join(HERE, "data", "dashboard_data.json")

ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
ANTHROPIC_VERSION = "2023-06-01"
MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5")

SYSTEM_PROMPT = """You are a senior financial analyst writing for SALIC's (Saudi Agricultural
and Livestock Investment Company) executive team. Answer questions about SALIC's group
financial performance using ONLY the JSON dataset provided below — never invent, extrapolate
or guess a number. If the dataset doesn't contain what's needed, say so plainly.

ENTITY / SPELLING TOLERANCE:
- Treat misspellings and near-matches of the company name as SALIC without correcting the
  user or saying you lack data for that name. Examples that ALL mean SALIC: Salik, Salic,
  SALIK, salic, salik, "Saudi Ag", "Saudi Agricultural", and similar typos.
- Do NOT claim you have no data on "Salik" / "Salic" — those are the same company as SALIC.
- Portfolio-company names in questions may also be misspelled; map them to the closest name
  that appears in the dataset waterfalls when the intent is clear (e.g. "Nadec" → NADEC).
- Use the official spelling "SALIC" in your answer narrative.

PERIOD TOLERANCE:
- The dataset covers YTD / LTM ending Dec-2025 only — not a multi-year 2021–2025 series.
- If the user asks for another year range, still answer with the available Dec-2025 YTD
  (and LTM where relevant) figures from the dataset. One short clause can note the coverage
  ("based on YTD Dec-2025 dashboard data"), then deliver the numbers — do not refuse or
  lecture. Include a helpful chart and/or table when the question asks for performance.

LANGUAGE: Always answer in the same language the question was asked in. If the question is
in Arabic, write the "answer" field — and any "chart"/"table" title, insight, or column text
you choose to include — in Arabic, using standard financial Modern Standard Arabic
terminology consistent with the CFO-memo tone described below. If the question is in
English, answer in English. The JSON *keys* themselves ("answer", "chart", "table", "type",
"title", "categories", "series", "name", "values", etc.) must always stay literal English
keys — only the human-readable string values change language.

The dataset has:
- "kpis": headline Actual/Budget/Prior-Year figures (Revenue, Gross Profit, EBITDA, Net Income,
  Working Capital, Net Debt to EBITDA, ROIC) for the YTD period ending Dec-2025.
- "waterfalls": bridge/waterfall breakdowns of each KPI by Geography, Segment or Portfolio
  Company, for YTD and LTM periods. Each waterfall goes from a starting bar "B" (budget/
  baseline) through positive/negative steps to an ending bar "A" (actual).

WRITE THE NARRATIVE LIKE A CFO MEMO, NOT A CAPTION:
- 3-6 sentences. Open with the headline number and direction vs budget and/or prior year
  (with both absolute and % magnitude). Then explain the primary driver using the waterfall
  breakdown data (name the specific geography/segment/company and its contribution). Add a
  secondary driver or offsetting factor if the data shows one. Close with a one-line
  so-what/implication for the business, still grounded only in the given numbers.
- Vary sentence structure; don't just restate the same "X was Y, a Z% change" template every
  time. Write like an analyst who has actually looked at the bridge, not a template filler.

CHOOSE THE CHART DELIBERATELY — avoid defaulting to a plain two-bar comparison every time:
- Comparing Actual vs Budget vs Prior Year for ONE OR MORE KPIs → "column" chart, categories =
  KPI names, with 2-3 series ("Actual", "Budget", "Prior Year") so it reads as a grouped chart,
  not a single bar pair.
- A waterfall/bridge breakdown (by geography, segment, or portfolio company) → "bar" or
  "column" with categories = step labels (B, driver names in order, A) and one series of values
  — this visually reconstructs the bridge. Never render a bridge as a pie.
- Composition / "what share of X" questions where all contributing values are positive and
  sum to something meaningful (e.g. only the positive contributors to a KPI, or a set of
  peer companies/segments compared side by side) → "pie" or "donut" (2-6 slices only).
- Ranking/"who contributed most" across companies or segments → "bar" (horizontal), sorted by
  magnitude if you can express that via category order.
- Multiple KPIs' variance vs budget side by side → "stacked_column" only if the components
  genuinely stack into a meaningful total; otherwise use grouped "column".
Only propose ONE chart, and only if it truly clarifies the answer. Chart schema (exact):
   {"type":"column|bar|stacked_column|line|area|pie|donut","title":"...","insight":"...",
    "unit":"SAR m","categories":[...],"series":[{"name":"...","values":["123","-45",...]}]}
   Every value in "values" is a STRING of digits only (optional leading "-" and decimal
   point), no thousands separators, no units. Convert raw dataset numbers to millions for
   readability (divide by 1,000,000) and set unit to "SAR m", unless the question is about a
   ratio/percentage KPI (Net Debt/EBITDA, ROIC) — then use the dataset's native unit ("x" or
   "%") and don't divide. "insight" is a short (<12 words) pointed takeaway, not a repeat of
   the title.

TABLES: if a table communicates better than a chart (e.g. a multi-KPI summary with several
columns of context), propose ONE "table": {"title":"...","columns":[...],"rows":[[...], ...]}.
A question can have a chart, a table, both, or neither — never force one that doesn't fit.

META / OUT-OF-SCOPE (still return the same JSON shape — never plain prose):
- Capability / greeting questions alone ("what can you do?", "how are you?") with no
  finance ask → 2-4 sentences explaining you analyze SALIC group financials (YTD Dec-2025).
  chart and table null.
- Greeting + a finance ask in the same message → skip the small talk beyond one short
  clause and answer the finance ask with numbers from the dataset (chart/table as usual).
- Truly unknown entities (not SALIC aliases and not close to any portfolio name in the
  dataset) → say you don't have that entity; do not invent figures. chart and table null.
- Keep "answer" concise when there is no chart/table so the JSON always finishes.

Respond with JSON only (no markdown fences, no commentary outside the JSON), exactly this
shape: {"answer":"...", "chart": {...} | null, "table": {...} | null}"""


def _api_key():
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if key:
        return key
    env_path = os.path.join(HERE, ".env")
    if os.path.exists(env_path):
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                m = re.match(r"\s*ANTHROPIC_API_KEY\s*=\s*(\S+)", line)
                if m:
                    return m.group(1).strip().strip('"').strip("'")
    return ""


def _load_dataset():
    with open(DATASET_PATH, encoding="utf-8") as f:
        return json.load(f)


def _valid_table(t):
    if not isinstance(t, dict):
        return None
    cols = [str(c) for c in (t.get("columns") or []) if str(c).strip()]
    if not cols:
        return None
    rows = []
    for r in (t.get("rows") or [])[:30]:
        if not isinstance(r, list) or not r:
            continue
        row = [str(c) for c in r[:len(cols)]]
        row += [""] * (len(cols) - len(row))
        rows.append(row)
    if not rows:
        return None
    return {"title": str(t.get("title") or "Table")[:120], "columns": cols[:8], "rows": rows}


def _extract_json(text):
    """Claude sometimes wraps JSON in prose or a markdown fence despite instructions;
    pull out the first {...} block and parse that."""
    text = (text or "").strip()
    m = re.search(r"```(?:json)?\s*(\{.*\})\s*```", text, re.S)
    if m:
        text = m.group(1)
    else:
        start, end = text.find("{"), text.rfind("}")
        if start != -1 and end != -1 and end > start:
            text = text[start:end + 1]
    return json.loads(text)


def _unescape_partial(raw):
    """Decode a JSON string fragment that may be missing its closing quote."""
    raw = raw or ""
    if raw.endswith("\\") and not raw.endswith("\\\\"):
        raw = raw[:-1]
    try:
        return json.loads('"' + raw + '"')
    except (ValueError, json.JSONDecodeError):
        return (raw.replace('\\"', '"').replace("\\n", "\n").replace("\\t", "\t")
                    .replace("\\\\", "\\")).strip()


def _salvage_answer(text):
    """Pull "answer" out even when the rest of the JSON is truncated/malformed.
    max_tokens often cuts mid-string before the closing quote — handle that too."""
    text = text or ""
    m = re.search(r'"answer"\s*:\s*"((?:[^"\\]|\\.)*)"', text, re.S)
    if m:
        return _unescape_partial(m.group(1)) or None
    # Truncated: "answer": "....  with no closing quote before end of text
    m = re.search(r'"answer"\s*:\s*"(.*)$', text, re.S)
    if m:
        salvaged = _unescape_partial(m.group(1))
        return salvaged or None
    return None


def _fallback_from_text(text):
    """Last resort when Claude ignored JSON: show usable prose instead of a dead-end error."""
    text = (text or "").strip()
    if not text:
        return None
    salvaged = _salvage_answer(text)
    if salvaged:
        return salvaged
    # Strip markdown fences / leading junk, keep remaining prose
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text).strip()
    if text.startswith("{") and '"answer"' not in text:
        return None
    if text.startswith("{"):
        return None
    return text[:4000]


def _normalize_question(question: str) -> str:
    """Map common SALIC misspellings so the model doesn't treat them as another company."""
    q = question or ""
    # Salik / Salic / SALIK / salic / etc. → SALIC (whole-word only)
    return re.sub(r"\bsali[ck]\b", "SALIC", q, flags=re.IGNORECASE)


def _prepare(question, history):
    """Normalize question + build Anthropic messages, or return an early result dict."""
    question = _normalize_question((question or "").strip())
    if not question:
        return None, {"answer": "Ask me something about SALIC's financial performance.",
                      "chart": None, "table": None, "ok": True}
    key = _api_key()
    if not key:
        return None, {"answer": "AI backend is not configured (missing ANTHROPIC_API_KEY).",
                      "chart": None, "table": None, "ok": False}

    messages = []
    for turn in (history or [])[-6:]:
        role = "assistant" if turn.get("role") == "assistant" else "user"
        content = turn.get("content")
        if content:
            messages.append({"role": role, "content": str(content)[:2000]})
    messages.append({"role": "user", "content": question})
    payload = {
        "model": MODEL,
        "max_tokens": 6000,
        "system": SYSTEM_PROMPT + "\n\nDATASET:\n" + json.dumps(_load_dataset()),
        "messages": messages,
    }
    return {"key": key, "payload": payload}, None


def _finalize_text(text: str) -> dict:
    """Parse Claude's full text into {answer, chart, table, ok}."""
    text = text or ""
    if not text.strip():
        return {"answer": "Sorry, I got an empty reply from the AI. Please try again.",
                "chart": None, "table": None, "ok": False}
    try:
        parsed = _extract_json(text)
    except (ValueError, json.JSONDecodeError) as e:
        print(f"[chat] JSON parse failed ({e}); salvaging answer text")
        salvaged = _fallback_from_text(text)
        if salvaged:
            return {"answer": salvaged, "chart": None, "table": None, "ok": True}
        return {"answer": "Sorry, that answer got cut off — please try asking again.",
                "chart": None, "table": None, "ok": False}

    answer = str(parsed.get("answer") or "").strip() or "I couldn't find that in the dashboard data."
    chart = _valid_spec(parsed.get("chart")) if parsed.get("chart") else None
    table = _valid_table(parsed.get("table")) if parsed.get("table") else None
    return {"answer": answer, "chart": chart, "table": table, "ok": True}


class _AnswerDeltaExtractor:
    """Emit readable answer characters as Claude streams JSON (or plain prose)."""

    def __init__(self):
        self.buf = ""
        self.mode = "detect"  # detect | json_answer | prose | done
        self.i = 0
        self.escape = False
        self.prose_emitted = 0

    def feed(self, chunk: str) -> list[str]:
        if not chunk or self.mode == "done":
            return []
        self.buf += chunk
        out: list[str] = []

        if self.mode == "detect":
            stripped = self.buf.lstrip()
            if not stripped:
                return []
            if stripped.startswith("{") or stripped.startswith("`"):
                m = re.search(r'"answer"\s*:\s*"', self.buf)
                if not m:
                    return []
                self.mode = "json_answer"
                self.i = m.end()
            else:
                self.mode = "prose"

        if self.mode == "prose":
            fresh = self.buf[self.prose_emitted:]
            if fresh:
                out.append(fresh)
                self.prose_emitted = len(self.buf)
            return out

        if self.mode == "json_answer":
            while self.i < len(self.buf):
                c = self.buf[self.i]
                self.i += 1
                if self.escape:
                    out.append({"n": "\n", "t": "\t", '"': '"', "\\": "\\", "/": "/"}.get(c, c))
                    self.escape = False
                    continue
                if c == "\\":
                    self.escape = True
                    continue
                if c == '"':
                    self.mode = "done"
                    break
                out.append(c)
        return out


def answer_question(question, history=None, timeout=60):
    """question + conversation history -> {"answer", "chart", "table", "ok"}."""
    prepared, early = _prepare(question, history)
    if early:
        return early

    try:
        r = requests.post(
            ANTHROPIC_URL,
            headers={
                "x-api-key": prepared["key"],
                "anthropic-version": ANTHROPIC_VERSION,
                "content-type": "application/json",
            },
            json=prepared["payload"],
            timeout=timeout,
        )
        r.raise_for_status()
        body = r.json()
        stop = body.get("stop_reason")
        blocks = body.get("content") or []
        text = next((b["text"] for b in blocks if b.get("type") == "text"), "")
        if stop == "max_tokens":
            print(f"[chat] response truncated (stop_reason=max_tokens, chars={len(text)})")
    except Exception as e:
        print(f"[chat] request failed: {e}")
        return {"answer": "Sorry, I couldn't reach the AI backend just now. Please try again.",
                "chart": None, "table": None, "ok": False}

    return _finalize_text(text)


def stream_answer_question(question, history=None, timeout=120):
    """Yield SSE-oriented dicts: {"type":"delta","text":...} then {"type":"done", ...result}.

    Streams the narrative as Claude generates JSON so the chat UI can type it out live;
    chart/table arrive on the final done event.
    """
    prepared, early = _prepare(question, history)
    if early:
        if early.get("answer"):
            yield {"type": "delta", "text": early["answer"]}
        yield {"type": "done", **early}
        return

    extractor = _AnswerDeltaExtractor()
    full = []
    try:
        with requests.post(
            ANTHROPIC_URL,
            headers={
                "x-api-key": prepared["key"],
                "anthropic-version": ANTHROPIC_VERSION,
                "content-type": "application/json",
            },
            json={**prepared["payload"], "stream": True},
            stream=True,
            timeout=timeout,
        ) as r:
            r.raise_for_status()
            for raw in r.iter_lines(decode_unicode=True):
                if not raw:
                    continue
                if raw.startswith("event:"):
                    continue
                if not raw.startswith("data:"):
                    continue
                data = raw[5:].strip()
                if not data or data == "[DONE]":
                    continue
                try:
                    event = json.loads(data)
                except json.JSONDecodeError:
                    continue
                et = event.get("type")
                if et == "content_block_delta":
                    piece = (event.get("delta") or {}).get("text") or ""
                    if not piece:
                        continue
                    full.append(piece)
                    deltas = extractor.feed(piece)
                    if deltas:
                        yield {"type": "delta", "text": "".join(deltas)}
                elif et == "message_delta":
                    stop = (event.get("delta") or {}).get("stop_reason")
                    if stop == "max_tokens":
                        print("[chat] stream truncated (stop_reason=max_tokens)")
                elif et == "error":
                    msg = (event.get("error") or {}).get("message") or "stream error"
                    print(f"[chat] stream error: {msg}")
                    fail = {"answer": "Sorry, I couldn't reach the AI backend just now. Please try again.",
                            "chart": None, "table": None, "ok": False}
                    yield {"type": "done", **fail}
                    return
    except Exception as e:
        print(f"[chat] stream request failed: {e}")
        fail = {"answer": "Sorry, I couldn't reach the AI backend just now. Please try again.",
                "chart": None, "table": None, "ok": False}
        yield {"type": "done", **fail}
        return

    result = _finalize_text("".join(full))
    # If we never managed to stream deltas (e.g. odd formatting), push the final answer once.
    if extractor.mode in ("detect",) and result.get("answer"):
        yield {"type": "delta", "text": result["answer"]}
    yield {"type": "done", **result}
