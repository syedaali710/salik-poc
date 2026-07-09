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


def _salvage_answer(text):
    """Extended thinking eats an unpredictable chunk of the token budget, so on long
    answers the trailing chart/table JSON occasionally gets cut off by max_tokens.
    "answer" is always written first, so pull it out with a regex even if the rest
    of the JSON is truncated/malformed, rather than showing nothing at all."""
    m = re.search(r'"answer"\s*:\s*"((?:[^"\\]|\\.)*)"', text or "", re.S)
    if not m:
        return None
    try:
        return json.loads('"' + m.group(1) + '"')
    except (ValueError, json.JSONDecodeError):
        return m.group(1).replace('\\"', '"').replace("\\n", "\n")


def answer_question(question, history=None, timeout=60):
    """question + conversation history -> {"answer", "chart", "table"}.
    chart/table may be None. Always returns a usable answer, even on failure,
    so the chat UI never has nothing to show."""
    question = (question or "").strip()
    if not question:
        return {"answer": "Ask me something about SALIC's financial performance.",
                "chart": None, "table": None}
    key = _api_key()
    if not key:
        return {"answer": "AI backend is not configured (missing ANTHROPIC_API_KEY).",
                "chart": None, "table": None}

    dataset = _load_dataset()
    messages = []
    for turn in (history or [])[-6:]:
        role = "assistant" if turn.get("role") == "assistant" else "user"
        content = turn.get("content")
        if content:
            messages.append({"role": role, "content": str(content)[:2000]})
    messages.append({"role": "user", "content": question})

    try:
        r = requests.post(
            ANTHROPIC_URL,
            headers={
                "x-api-key": key,
                "anthropic-version": ANTHROPIC_VERSION,
                "content-type": "application/json",
            },
            json={
                "model": MODEL,
                "max_tokens": 6000,
                "system": SYSTEM_PROMPT + "\n\nDATASET:\n" + json.dumps(dataset),
                "messages": messages,
            },
            timeout=timeout,
        )
        r.raise_for_status()
        blocks = r.json()["content"]
        text = next(b["text"] for b in blocks if b.get("type") == "text")
    except Exception as e:
        print(f"[chat] request failed: {e}")
        return {"answer": "Sorry, I couldn't reach the AI backend just now. Please try again.",
                "chart": None, "table": None}

    try:
        parsed = _extract_json(text)
    except (ValueError, json.JSONDecodeError) as e:
        print(f"[chat] JSON parse failed ({e}); salvaging answer text")
        salvaged = _salvage_answer(text)
        if salvaged:
            return {"answer": salvaged, "chart": None, "table": None}
        return {"answer": "Sorry, that answer got cut off — please try asking again.",
                "chart": None, "table": None}

    answer = str(parsed.get("answer") or "").strip() or "I couldn't find that in the dashboard data."
    chart = _valid_spec(parsed.get("chart")) if parsed.get("chart") else None
    table = _valid_table(parsed.get("table")) if parsed.get("table") else None
    return {"answer": answer, "chart": chart, "table": table}
