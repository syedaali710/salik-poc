"""
Turns a spoken/typed finance update into structured slide fields.
Deliberately simple heuristics — in production this step is done by Azure OpenAI
and the numbers come from Microsoft Fabric / Power BI, not from the transcript.
Every field is editable in the UI, so imperfect parsing is fine.
"""
import re

COMPANIES = {
    "nadec": "NADEC", "cfg": "CFG", "continental": "CFG", "sabil": "SABIL",
    "naqua": "NAQUA", "al marai": "Almarai", "almarai": "Almarai",
    "minerva": "Minerva Foods", "brf": "BRF", "mhp": "MHP", "olam": "Olam Agri",
    "g3": "G3", "ufic": "UFIC", "national grain": "NGC", "ngc": "NGC",
    "lt foods": "LT Foods", "merredin": "MFPL", "mfpl": "MFPL", "salic": "SALIC",
}

NUMRE = re.compile(r'(-?\d[\d,]*\.?\d*)\s*(billion|million|bn|b|m)?', re.I)


def to_millions(numstr, unit):
    try:
        n = float(numstr.replace(',', ''))
    except ValueError:
        return None
    if unit and re.fullmatch(r'(bn|billion|b)', unit, re.I):
        n *= 1000
    return n


def extract_kpi(text, keys):
    lower = text.lower()
    idx, hit = -1, ""
    for k in keys:
        p = lower.find(k)
        if p >= 0 and (idx < 0 or p < idx):
            idx, hit = p, k
    if idx < 0:
        return None
    seg = text[idx:idx + 180]
    stops = [s for s in ["revenue", "ebitda", "net income", "net profit"] if s not in keys]
    cut = len(seg)
    for s in stops:
        p = seg.lower().find(s, len(hit))
        if p > 0:
            cut = min(cut, p)
    # also stop at the end of the sentence that mentions this KPI, so numbers
    # from the NEXT sentence (e.g. a year like "...in February 2025.") don't bleed in.
    mterm = re.search(r'[.!?](\s|$)', seg[len(hit):])
    if mterm:
        cut = min(cut, len(hit) + mterm.end())
    seg = seg[:cut]
    res = {"actual": None, "budget": None, "py": None}
    for m in NUMRE.finditer(seg):
        unit_raw = m.group(2)
        val = to_millions(m.group(1), unit_raw or "m")
        if val is None:
            continue
        before = seg[max(0, m.start() - 22):m.start()].lower()
        after = seg[m.end():m.end() + 22].lower()
        ctx = before + " " + after
        # ignore bare 4-digit years (2000-2099) used as dates, not financial values
        if (unit_raw is None and 2000 <= val <= 2099
                and re.search(r'(in|by|during|since|year|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+$', before)):
            continue
        signed = -abs(val) if re.search(r'negative|loss of|loss', before) else val
        if re.search(r'budget|target|plan', before) and res["budget"] is None:
            res["budget"] = signed
            continue
        if re.search(r'last year|prior year|previous year|year on year|2024', ctx) and res["py"] is None:
            res["py"] = signed
            continue
        if res["actual"] is None:
            res["actual"] = signed
            continue
        if res["budget"] is None:
            res["budget"] = signed
            continue
        if res["py"] is None:
            res["py"] = signed
            continue
    return res


def parse(text):
    text = (text or "").strip()
    d = {"company": "", "period": "", "unit": "SAR m", "kpis": {}, "comm": []}
    lower = text.lower()
    for k, v in COMPANIES.items():
        if k in lower:
            d["company"] = v
            break
    pm = re.search(r"q\s*([1-4])\s*'?\s*(20)?(\d{2})", text, re.I)
    if pm:
        d["period"] = "Q" + pm.group(1) + " 20" + pm.group(3)
    elif re.search(r'first quarter', text, re.I):
        d["period"] = "Q1 2025"
    else:
        y = re.search(r'\b20(2\d)\b', text)
        d["period"] = "FY 20" + y.group(1) if y else "Q1 2025"

    d["kpis"]["rev"] = extract_kpi(text, ["revenue"]) or _empty()
    d["kpis"]["ebitda"] = extract_kpi(text, ["ebitda"]) or _empty()
    d["kpis"]["ni"] = extract_kpi(text, ["net income", "net profit"]) or _empty()

    for s in re.split(r'(?<=[.!?])\s+', text):
        s = s.strip()
        if not s:
            continue
        sl = s.lower()
        num_heavy = (re.search(r'revenue|ebitda|net income|net profit', sl)
                     and re.search(r'\d', s)
                     and re.search(r'budget|last year|prior|versus|against|vs', sl))
        if num_heavy or re.search(r'performance summary', s, re.I):
            continue
        if len(s) > 8:
            d["comm"].append(re.sub(r'\s+', ' ', s))
    if not d["comm"]:
        d["comm"] = ["Add commentary points here…"]
    return d


def _empty():
    return {"actual": None, "budget": None, "py": None}
