"""
Builds the deck by editing the official SALIC template's own slides in place
(assets/SALIC_template.pptx) and adding chart slides on the template's
"Content Slide Light" layout, so every slide keeps the exact SALIC design:

  template slide 1  -> Cover        ("Insert Your Title Here")
  template slide 2  -> Agenda       (numbered "Insert Title Here" list)   [optional]
  template slide 3  -> Section      ("Section title here" + subtitle)
  template slide 4  -> Sub-section  ("Add your Subsection Title")
  template slide 8  -> Content      (Analysis 01..04 blocks -> KPIs + commentary)
  + chart slide     -> Actual vs Budget vs Prior Year (native pptx chart)  [optional]
  + chart slide     -> Variance vs Budget (native pptx chart)              [optional]
  template slide 10 -> Thank You    (kept as-is)                           [optional]

Template slides 5-7 and 9 (brand-guide/demo pages) are removed from the output.
Returns an in-memory .pptx ready to stream back to the browser.
"""
import io, os
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "assets", "SALIC_template.pptx")

# ---- SALIC brand palette (from ppt/theme/theme1.xml in the official template) ----
NAVY = RGBColor(0x0B, 0x4D, 0x89)
SKY = RGBColor(0x33, 0x9B, 0xD6)
GREEN = RGBColor(0x2A, 0x7C, 0x62)
LIGHTGREEN = RGBColor(0xA4, 0xD5, 0xAD)
DARKNAVY = RGBColor(0x08, 0x30, 0x5C)
RED = RGBColor(0xD9, 0x3F, 0x40)

# 0-based indices of the template slides we use
COVER, AGENDA, SECTION, SUBSECTION, CONTENT, THANKYOU = 0, 1, 2, 3, 7, 9

CONTENT_LAYOUT_NAME = "Content Slide Light"
TITLE_PLACEHOLDER_IDX = 11
EYEBROW_PLACEHOLDER_IDX = 10

KPIDEFS = [("rev", "Revenue"), ("ebitda", "EBITDA"), ("ni", "Net Income")]


def fmt(n):
    if n is None or n == "":
        return "—"
    try:
        n = float(n)
    except (ValueError, TypeError):
        return "—"
    neg = n < 0
    n = abs(n)
    s = f"{round(n):,}" if n >= 100 else (f"{n:,.1f}".rstrip("0").rstrip("."))
    return (f"({s})" if neg else s)


def _get(k, key):
    return (k or {}).get(key)


def _set_lines(text_frame, lines):
    """Replace a text frame's paragraphs with `lines`, preserving the template's
    formatting by reusing existing paragraphs/runs (cloning the last paragraph
    when more are needed)."""
    lines = [str(l) for l in lines if str(l).strip()] or [""]
    while len(text_frame.paragraphs) < len(lines):
        last_p = text_frame.paragraphs[-1]._p
        last_p.addnext(deepcopy(last_p))
    for i, para in enumerate(list(text_frame.paragraphs)):
        if i < len(lines):
            runs = para.runs
            if runs:
                runs[0].text = lines[i]
                for r in runs[1:]:
                    r._r.getparent().remove(r._r)
            else:
                para.add_run().text = lines[i]
        else:
            para._p.getparent().remove(para._p)


def _shape_with_text(slide, needle):
    for shape in slide.shapes:
        if shape.has_text_frame and needle.lower() in shape.text_frame.text.lower():
            return shape
    return None


def _set_title_placeholders(slide, title, eyebrow):
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == TITLE_PLACEHOLDER_IDX and title is not None:
            _set_lines(ph.text_frame, [title])
        elif idx == EYEBROW_PLACEHOLDER_IDX and eyebrow is not None:
            _set_lines(ph.text_frame, [eyebrow])


def _content_layout(prs):
    for layout in prs.slide_masters[0].slide_layouts:
        if layout.name == CONTENT_LAYOUT_NAME:
            return layout
    return prs.slide_layouts[0]


# ---------------------------------------------------------------- template slides

def _fill_cover(slide, title):
    ph = _shape_with_text(slide, "Insert Your Title")
    if ph is not None:
        _set_lines(ph.text_frame, [title])


def _fill_agenda(slide, items):
    ph = _shape_with_text(slide, "Insert Title Here")
    if ph is not None:
        _set_lines(ph.text_frame, items)


def _fill_section(slide, title, subtitle):
    t = _shape_with_text(slide, "Section title here")
    if t is not None:
        _set_lines(t.text_frame, [title])
    s = _shape_with_text(slide, "Subtitle here")
    if s is not None:
        _set_lines(s.text_frame, [subtitle])


def _fill_subsection(slide, title):
    ph = _shape_with_text(slide, "Subsection Title")
    if ph is not None:
        _set_lines(ph.text_frame, [title])


def _fill_content(slide, kpis, comm, unit, title, eyebrow):
    """Fill the four 'Analysis 0N' blocks: three KPIs + one commentary."""
    _set_title_placeholders(slide, title, eyebrow)

    titles, bodies = {}, {}
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if text.startswith("Analysis 0"):
            titles[text[:11]] = shape
        elif "Synth chartreuse" in text:
            # body boxes are unlabeled; key them by position (row, col)
            bodies[(shape.top, shape.left)] = shape

    # order bodies to match Analysis 01..04: top-left, top-right, bottom-left, bottom-right
    ordered_bodies = [bodies[k] for k in sorted(bodies)]
    ordered_titles = [titles.get(f"Analysis 0{i}") for i in (1, 2, 3, 4)]

    blocks = []
    for key, label in KPIDEFS:
        k = kpis.get(key) or {}
        a, b, py = _get(k, "actual"), _get(k, "budget"), _get(k, "py")
        lines = [f"Actual: {fmt(a)} {unit}"]
        sub = []
        if b is not None:
            sub.append(f"Budget: {fmt(b)}")
        if py is not None:
            sub.append(f"Prior Year: {fmt(py)}")
        if sub:
            lines.append("  ·  ".join(sub))
        if a is not None and b is not None:
            diff = a - b
            lines.append(("Above budget by " if diff >= 0 else "Below budget by ") + fmt(abs(diff)))
        blocks.append((label, lines))
    blocks.append(("Commentary", comm or ["—"]))

    for i, (label, lines) in enumerate(blocks[:4]):
        if i < len(ordered_titles) and ordered_titles[i] is not None:
            _set_lines(ordered_titles[i].text_frame, [label])
        if i < len(ordered_bodies):
            _set_lines(ordered_bodies[i].text_frame, lines)


# ---------------------------------------------------------------- chart slides

def _add_titled_slide(prs, title, eyebrow):
    slide = prs.slides.add_slide(_content_layout(prs))
    _set_title_placeholders(slide, title, eyebrow)
    return slide


def _style_chart(chart):
    chart.has_title = False
    chart.font.size = Pt(16)
    chart.font.color.rgb = DARKNAVY


def _kpi_chart_slide(prs, kpis, unit, eyebrow):
    cats, py, bud, act = [], [], [], []
    for key, label in KPIDEFS:
        k = kpis.get(key) or {}
        a, b, p = _get(k, "actual"), _get(k, "budget"), _get(k, "py")
        if a is None and b is None and p is None:
            continue
        cats.append(label)
        act.append(a if a is not None else 0)
        bud.append(b if b is not None else 0)
        py.append(p if p is not None else 0)
    if not cats:
        return None

    slide = _add_titled_slide(prs, f"KPI Overview — Actual vs Budget vs Prior Year ({unit})", eyebrow)
    data = CategoryChartData()
    data.categories = cats
    data.add_series("Prior Year", py)
    data.add_series("Budget", bud)
    data.add_series("Actual", act)
    gf = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1.4), Inches(3.2), Inches(23.8), Inches(10.4), data,
    )
    chart = gf.chart
    _style_chart(chart)
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    for series, color in zip(chart.plots[0].series, (LIGHTGREEN, SKY, NAVY)):
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = color
    chart.value_axis.tick_labels.number_format = "#,##0"
    chart.value_axis.tick_labels.number_format_is_linked = False
    return slide


def _variance_chart_slide(prs, kpis, unit, eyebrow):
    cats, vals = [], []
    for key, label in KPIDEFS:
        k = kpis.get(key) or {}
        a, b = _get(k, "actual"), _get(k, "budget")
        if a is None or b is None:
            continue
        cats.append(label)
        vals.append(a - b)
    if not cats:
        return None

    slide = _add_titled_slide(prs, f"Variance vs Budget ({unit})", eyebrow)
    data = CategoryChartData()
    data.categories = cats
    data.add_series("Variance vs Budget", vals)
    gf = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1.4), Inches(3.2), Inches(23.8), Inches(10.4), data,
    )
    chart = gf.chart
    _style_chart(chart)
    chart.has_legend = False
    series = chart.plots[0].series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = GREEN
    try:  # colour each bar green/red by sign
        for i, v in enumerate(vals):
            point = series.points[i]
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = GREEN if v >= 0 else RED
    except Exception:
        pass
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = "#,##0;(#,##0)"
    dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(18)
    dl.font.bold = True
    dl.font.color.rgb = DARKNAVY
    chart.value_axis.tick_labels.number_format = "#,##0;(#,##0)"
    chart.value_axis.tick_labels.number_format_is_linked = False
    return slide


# ---------------------------------------------------------------- assembly

def _arrange(prs, ordered_ids):
    """Rewrite the slide list to exactly `ordered_ids`, in order."""
    lst = prs.slides._sldIdLst
    elems = {int(e.get("id")): e for e in lst}
    for e in list(lst):
        lst.remove(e)
    for sid in ordered_ids:
        lst.append(elems[sid])


def build_pptx(d):
    d = d or {}
    company = (d.get("company") or "Company").strip()
    period = (d.get("period") or "").strip()
    unit = (d.get("unit") or "SAR m").strip()
    kpis = d.get("kpis") or {}
    comm = [str(c) for c in (d.get("comm") or []) if str(c).strip()][:5]
    include = d.get("include") or {}
    inc_agenda = include.get("agenda", True)
    inc_charts = include.get("charts", True)
    inc_thanks = include.get("thankyou", True)

    prs = Presentation(TEMPLATE)
    slides = list(prs.slides)

    title = f"{company} Performance Summary"
    if period:
        title += f" — {period}"
    eyebrow = " · ".join(x for x in (company, period) if x)

    agenda = ["Performance Overview", "Revenue", "EBITDA", "Net Income"]
    if inc_charts:
        agenda.append("KPI Charts")
    agenda.append("Commentary & Outlook")

    _fill_cover(slides[COVER], title)
    if inc_agenda:
        _fill_agenda(slides[AGENDA], agenda)
    _fill_section(
        slides[SECTION],
        (period + " Performance").strip() if period else "Performance",
        f"All figures in Million {unit.replace('m', '').strip() or 'SAR'} unless stated",
    )
    _fill_subsection(slides[SUBSECTION], f"{company} Financial Highlights")
    _fill_content(slides[CONTENT], kpis, comm, unit,
                  "Financial Performance", eyebrow)

    chart_slides = []
    if inc_charts:
        for maker in (_kpi_chart_slide, _variance_chart_slide):
            s = maker(prs, kpis, unit, eyebrow)
            if s is not None:
                chart_slides.append(s)

    ordered = [slides[COVER].slide_id]
    if inc_agenda:
        ordered.append(slides[AGENDA].slide_id)
    ordered += [slides[SECTION].slide_id, slides[SUBSECTION].slide_id,
                slides[CONTENT].slide_id]
    ordered += [s.slide_id for s in chart_slides]
    if inc_thanks:
        ordered.append(slides[THANKYOU].slide_id)

    _arrange(prs, ordered)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
