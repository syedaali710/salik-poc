"""
Builds the deck by editing the official SALIC template's own slides in place
(assets/SALIC_template.pptx) and adding chart slides on the template's
"Content Slide Light" layout, so every slide keeps the exact SALIC design:

  template slide 1  -> Cover        ("Insert Your Title Here")
  template slide 2  -> Agenda       (numbered "Insert Title Here" list)   [optional]
  template slide 3  -> Section      ("Section title here" + subtitle)
  template slide 4  -> Sub-section  ("Add your Subsection Title")
  template slide 8  -> Content      (Analysis 01..04 blocks -> KPIs + commentary)
  + chart slide     -> Actual vs Budget vs Prior Year (small multiples)   [optional]
  + chart slide     -> Variance vs Budget (native pptx chart)             [optional]
  template slide 10 -> Thank You    (kept as-is)                          [optional]

Template slides 5-7 and 9 (brand-guide/demo pages) are removed from the output.
Returns an in-memory .pptx ready to stream back to the browser.
"""
import io, os
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "assets", "SALIC_template.pptx")

# ---- SALIC brand palette (from ppt/theme/theme1.xml in the official template) ----
NAVY = RGBColor(0x0B, 0x4D, 0x89)
SKY = RGBColor(0x33, 0x9B, 0xD6)
GREEN = RGBColor(0x2A, 0x7C, 0x62)
LIGHTGREEN = RGBColor(0xA4, 0xD5, 0xAD)
DARKNAVY = RGBColor(0x08, 0x30, 0x5C)
RED = RGBColor(0xD9, 0x3F, 0x40)
GREY = RGBColor(0x6B, 0x6B, 0x6B)

# 0-based indices of the template slides we use. This template has no
# standalone Agenda slide (only a thumbnail example on its brand-guide page),
# so AGENDA is None and every AGENDA usage below is guarded accordingly.
COVER, SECTION, SUBSECTION, CONTENT, THANKYOU = 0, 1, 2, 6, 8
AGENDA = None

CONTENT_LAYOUT_NAME = "Content Slide Light"
TITLE_PLACEHOLDER_IDX = 2
EYEBROW_PLACEHOLDER_IDX = 1

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
    when more are needed).

    Each entry in `lines` is either a plain string, or a list of
    (text, overrides) run tuples where overrides may set size (Pt), bold,
    and color (RGBColor) on top of the template's run formatting."""
    norm = []
    for l in lines:
        if isinstance(l, list):
            if any(str(t).strip() for t, _ in l):
                norm.append(l)
        elif str(l).strip():
            norm.append([(str(l), None)])
    lines = norm or [[("", None)]]

    while len(text_frame.paragraphs) < len(lines):
        last_p = text_frame.paragraphs[-1]._p
        last_p.addnext(deepcopy(last_p))
    for i, para in enumerate(list(text_frame.paragraphs)):
        if i >= len(lines):
            para._p.getparent().remove(para._p)
            continue
        runs = para.runs
        if not runs:
            runs = [para.add_run()]
        # one template run kept as the style prototype; clone it per requested run
        proto = runs[0]
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
        needed = lines[i]
        while len(para.runs) < len(needed):
            proto._r.addnext(deepcopy(proto._r))
        for run, (text, ov) in zip(para.runs, needed):
            run.text = text
            if ov:
                if "size" in ov:
                    run.font.size = ov["size"]
                if "bold" in ov:
                    run.font.bold = ov["bold"]
                if "color" in ov:
                    run.font.color.rgb = ov["color"]


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
            # the layout's eyebrow box is narrow; widen it so an insight
            # sentence fits on one line
            ph.width = Inches(22.0)
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
        # the template centers these lines; a numbered list needs a
        # straight left edge
        for para in ph.text_frame.paragraphs:
            para.alignment = PP_ALIGN.LEFT


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


def _variance_pct(a, b):
    try:
        return (a - b) / abs(b) * 100.0
    except (TypeError, ZeroDivisionError):
        return None


def _fill_content(slide, kpis, comm, unit, title, eyebrow):
    """Fill the four 'Analysis 0N' blocks: three KPIs + one commentary.
    The template's own vector icons next to each block are kept as-is.
    Each KPI block gets a big headline number, a context line and a
    green/red variance line."""
    _set_title_placeholders(slide, title, eyebrow)

    titles, bodies = {}, {}
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if text.startswith("Analysis 0"):
            titles[text[:11]] = shape
        elif "Synth chartreuse" in text:
            bodies[(shape.top, shape.left)] = shape

    # order bodies to match Analysis 01..04: top-left, top-right, bottom-left, bottom-right
    ordered_bodies = [bodies[k] for k in sorted(bodies)]
    ordered_titles = [titles.get(f"Analysis 0{i}") for i in (1, 2, 3, 4)]

    blocks = []
    for key, label in KPIDEFS:
        k = kpis.get(key) or {}
        a, b, py = _get(k, "actual"), _get(k, "budget"), _get(k, "py")
        lines = [[(fmt(a), {"size": Pt(40), "bold": True, "color": NAVY}),
                  (f"  {unit}", {"size": Pt(20), "bold": False, "color": GREY})]]
        sub = []
        if b is not None:
            sub.append(f"Budget {fmt(b)}")
        if py is not None:
            sub.append(f"Prior Year {fmt(py)}")
        if sub:
            lines.append([("   ·   ".join(sub), {"size": Pt(16), "color": GREY})])
        if a is not None and b is not None:
            diff = a - b
            pct = _variance_pct(a, b)
            arrow = "▲" if diff >= 0 else "▼"
            col = GREEN if diff >= 0 else RED
            txt = f"{arrow} {fmt(abs(diff))} vs Budget"
            if pct is not None:
                txt += f" ({abs(pct):.1f}%)"
            lines.append([(txt, {"size": Pt(16), "bold": True, "color": col})])
        blocks.append((label, lines))

    comm_lines = [[(f"• {c}", {"size": Pt(16)})] for c in (comm or ["—"])]
    blocks.append(("Commentary", comm_lines))

    for i, (label, lines) in enumerate(blocks[:4]):
        if i < len(ordered_titles) and ordered_titles[i] is not None:
            _set_lines(ordered_titles[i].text_frame, [label])
        if i < len(ordered_bodies):
            body = ordered_bodies[i]
            # give the taller styled content room to breathe
            body.height = Inches(2.6)
            body.text_frame.word_wrap = True
            _set_lines(body.text_frame, lines)


# ---------------------------------------------------------------- chart helpers

def _add_titled_slide(prs, title, insight, insight_color=GREEN):
    slide = prs.slides.add_slide(_content_layout(prs))
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == TITLE_PLACEHOLDER_IDX:
            _set_lines(ph.text_frame, [title])
        elif idx == EYEBROW_PLACEHOLDER_IDX and insight:
            ph.width = Inches(22.0)
            _set_lines(ph.text_frame,
                       [[(insight, {"bold": True, "color": insight_color})]])
    return slide


def _style_chart(chart, base_size=Pt(16)):
    chart.has_title = False
    chart.font.size = base_size
    chart.font.bold = False
    chart.font.color.rgb = DARKNAVY


def _style_axis(axis, font_size=Pt(14), gridlines=True):
    axis.tick_labels.font.size = font_size
    axis.tick_labels.font.color.rgb = DARKNAVY
    if not gridlines:
        axis.has_major_gridlines = False
        return
    axis.has_major_gridlines = True
    mg = axis._element.find(qn("c:majorGridlines"))
    if mg is None:
        return
    # light 0.5pt grey gridlines
    for child in list(mg):
        mg.remove(child)
    sp_pr = etree.SubElement(mg, qn("c:spPr"))
    ln = etree.SubElement(sp_pr, qn("a:ln"))
    ln.set("w", "6350")
    fill = etree.SubElement(ln, qn("a:solidFill"))
    etree.SubElement(fill, qn("a:srgbClr")).set("val", "E3E3E3")


def _ensure_orientation(axis):
    """Explicitly write <c:orientation val="minMax"/> under <c:scaling>.
    python-pptx omits this element when the orientation is the default
    minMax, leaving <c:scaling/> completely empty — some viewers render a
    chart with no explicit axis orientation as blank, so make it explicit."""
    scaling = axis._element.find(qn("c:scaling"))
    if scaling is None or scaling.find(qn("c:orientation")) is not None:
        return
    orient = scaling.makeelement(qn("c:orientation"), {})
    orient.set("val", "minMax")
    scaling.insert(0, orient)


def _hide_axis_line(axis):
    ax = axis._element
    sp_pr = ax.find(qn("c:spPr"))
    if sp_pr is None:
        # insert spPr in schema position (before txPr)
        sp_pr = ax.makeelement(qn("c:spPr"), {})
        tx_pr = ax.find(qn("c:txPr"))
        if tx_pr is not None:
            tx_pr.addprevious(sp_pr)
        else:
            ax.append(sp_pr)
    ln = etree.SubElement(sp_pr, qn("a:ln"))
    etree.SubElement(ln, qn("a:noFill"))


def _set_gap_width(chart, gap_pct=120):
    for tag in ("c:barChart", "c:bar3DChart"):
        el = chart._element.find(f".//{qn(tag)}")
        if el is not None:
            gw = el.find(qn("c:gapWidth"))
            if gw is None:
                gw = etree.SubElement(el, qn("c:gapWidth"))
                # move into schema order: gapWidth comes right after the last ser
                sers = el.findall(qn("c:ser"))
                if sers:
                    sers[-1].addnext(gw)
            gw.set("val", str(gap_pct))


def _series_labels(series, font_size, color, number_format="#,##0;(#,##0)",
                   pos="outEnd"):
    """Value-only data labels above each point, built in correct schema order
    (numFmt, spPr, txPr, dLblPos, show*). `pos=None` omits the position element
    (required for chart types that don't allow dLblPos, e.g. area)."""
    ser = series._element
    d_lbls = ser.find(qn("c:dLbls"))
    if d_lbls is not None:
        ser.remove(d_lbls)
    d_lbls = ser.makeelement(qn("c:dLbls"), {})
    # dLbls must come right before c:cat (or c:val) within c:ser
    cat = ser.find(qn("c:cat"))
    if cat is not None:
        cat.addprevious(d_lbls)
    else:
        ser.append(d_lbls)

    num_fmt = etree.SubElement(d_lbls, qn("c:numFmt"))
    num_fmt.set("formatCode", number_format)
    num_fmt.set("sourceLinked", "0")

    tx_pr = etree.SubElement(d_lbls, qn("c:txPr"))
    etree.SubElement(tx_pr, qn("a:bodyPr"))
    etree.SubElement(tx_pr, qn("a:lstStyle"))
    p = etree.SubElement(tx_pr, qn("a:p"))
    p_pr = etree.SubElement(p, qn("a:pPr"))
    def_rpr = etree.SubElement(p_pr, qn("a:defRPr"))
    def_rpr.set("sz", str(int(font_size.pt * 100)))
    def_rpr.set("b", "1")
    fill = etree.SubElement(def_rpr, qn("a:solidFill"))
    etree.SubElement(fill, qn("a:srgbClr")).set("val", f"{color:06X}" if isinstance(color, int) else str(color))
    etree.SubElement(p, qn("a:endParaRPr")).set("lang", "en-US")

    if pos:
        pos_el = etree.SubElement(d_lbls, qn("c:dLblPos"))
        pos_el.set("val", pos)
    for tag, val in (("c:showLegendKey", "0"), ("c:showVal", "1"),
                     ("c:showCatName", "0"), ("c:showSerName", "0"),
                     ("c:showPercent", "0"), ("c:showBubbleSize", "0")):
        etree.SubElement(d_lbls, qn(tag)).set("val", val)


def _panel_chart(slide, x, y, w, h, label, triples, unit, max_val):
    """One small-multiple panel: PY / Budget / Actual bars for a single KPI."""
    data = CategoryChartData()
    data.categories = ["Prior Year", "Budget", "Actual"]
    data.add_series(label, triples)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, w, h, data)
    chart = gf.chart
    _style_chart(chart, base_size=Pt(14))
    chart.has_legend = False

    series = chart.plots[0].series[0]
    for i, col in enumerate((LIGHTGREEN, SKY, NAVY)):
        point = series.points[i]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = col

    _series_labels(series, Pt(16), str(DARKNAVY))

    va = chart.value_axis
    va.minimum_scale = 0
    va.maximum_scale = max_val
    va.visible = False
    va.has_major_gridlines = False
    _style_axis(chart.category_axis, font_size=Pt(14), gridlines=False)
    _hide_axis_line(chart.category_axis)
    _set_gap_width(chart, gap_pct=60)
    return chart


def _headroom(v):
    """Round a max value up to a clean axis ceiling with ~15% headroom."""
    if v <= 0:
        return 1
    target = v * 1.18
    mag = 10 ** max(0, len(str(int(target))) - 2)
    return ((int(target) // mag) + 1) * mag


# ---------------------------------------------------------------- chart slides

def _kpi_insight(kpis):
    ahead, behind = [], []
    for key, label in KPIDEFS:
        k = kpis.get(key) or {}
        a, b = _get(k, "actual"), _get(k, "budget")
        if a is None or b is None:
            continue
        (ahead if a >= b else behind).append(label)
    if ahead and behind:
        return f"{', '.join(ahead)} ahead of budget; {', '.join(behind)} behind", GREEN
    if ahead:
        return f"All KPIs at or above budget: {', '.join(ahead)}", GREEN
    if behind:
        return f"{', '.join(behind)} below budget", RED
    return None, GREEN


def _kpi_chart_slide(prs, kpis, unit, period):
    panels = []
    for key, label in KPIDEFS:
        k = kpis.get(key) or {}
        a, b, p = _get(k, "actual"), _get(k, "budget"), _get(k, "py")
        if a is None and b is None and p is None:
            continue
        panels.append((label, [p or 0, b or 0, a or 0]))
    if not panels:
        return None

    insight, col = _kpi_insight(kpis)
    title = f"KPI Overview ({unit})"
    if period:
        title = f"{period} KPI Overview ({unit})"
    slide = _add_titled_slide(prs, title, insight, col)

    # one panel per KPI so each metric gets its own scale
    n = len(panels)
    margin, gutter = Inches(1.3), Inches(0.9)
    top, height = Inches(3.4), Inches(10.2)
    total_w = prs.slide_width - 2 * margin - gutter * (n - 1)
    w = int(total_w / n)

    for i, (label, triples) in enumerate(panels):
        x = margin + i * (w + gutter)
        # panel heading above each chart
        tb = slide.shapes.add_textbox(x, Inches(2.55), w, Inches(0.7))
        tf = tb.text_frame
        tf.word_wrap = False
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        run = para.add_run()
        run.text = label
        run.font.size = Pt(22)
        run.font.bold = True
        run.font.color.rgb = DARKNAVY
        _panel_chart(slide, x, top, w, height, label, triples, unit,
                     _headroom(max(triples)))
    return slide


def _variance_chart_slide(prs, kpis, unit, period):
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

    worst = min(vals)
    best = max(vals)
    if all(v >= 0 for v in vals):
        insight, col = "All KPIs delivered at or above budget", GREEN
    elif all(v < 0 for v in vals):
        insight, col = "All KPIs came in below budget", RED
    else:
        neg = [c for c, v in zip(cats, vals) if v < 0]
        insight, col = f"{', '.join(neg)} below budget — all other KPIs ahead", RED

    title = f"Variance vs Budget ({unit})"
    if period:
        title = f"{period} Variance vs Budget ({unit})"
    slide = _add_titled_slide(prs, title, insight, col)

    data = CategoryChartData(number_format="#,##0;(#,##0)")
    data.categories = cats
    data.add_series("Variance vs Budget", vals)
    gf = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(3.2), Inches(3.2), Inches(20.2), Inches(10.2), data,
    )
    chart = gf.chart
    _style_chart(chart)
    chart.has_legend = False

    series = chart.plots[0].series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = GREEN
    for i, v in enumerate(vals):
        point = series.points[i]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = GREEN if v >= 0 else RED

    _series_labels(series, Pt(18), str(DARKNAVY))

    va = chart.value_axis
    va.tick_labels.number_format = "#,##0;(#,##0)"
    va.tick_labels.number_format_is_linked = False
    _style_axis(va, font_size=Pt(13))
    ca = chart.category_axis
    _style_axis(ca, font_size=Pt(18), gridlines=False)
    ca.tick_labels.font.bold = True
    _set_gap_width(chart, gap_pct=90)
    return slide


# ------------------------------------------------- AI-planned chart slides

YELLOW = RGBColor(0xE6, 0xAF, 0x00)
SERIES_PALETTE = (NAVY, SKY, GREEN, LIGHTGREEN, RED, YELLOW)

PLANNED_XL_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "stacked_column": XL_CHART_TYPE.COLUMN_STACKED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "area": XL_CHART_TYPE.AREA,
    "pie": XL_CHART_TYPE.PIE,
    "donut": XL_CHART_TYPE.DOUGHNUT,
}


def _spec_number_format(spec):
    vals = [v for s in spec["series"] for v in s["values"] if v is not None]
    if vals and all(abs(v) < 100 for v in vals) and any(v != int(v) for v in vals):
        return "#,##0.0;(#,##0.0)"
    return "#,##0;(#,##0)"


def _set_overlap(chart, pct):
    for tag in ("c:barChart", "c:bar3DChart"):
        el = chart._element.find(f".//{qn(tag)}")
        if el is not None:
            ov = el.find(qn("c:overlap"))
            if ov is None:
                ov = etree.SubElement(el, qn("c:overlap"))
                gw = el.find(qn("c:gapWidth"))
                if gw is not None:
                    gw.addnext(ov)
            ov.set("val", str(pct))


def _pie_labels(chart, font_size=Pt(16)):
    """Percentage labels on slices, category names in the legend."""
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.show_percentage = True
    dl.show_value = False
    dl.show_category_name = False
    dl.number_format = "0%"
    dl.number_format_is_linked = False
    dl.font.size = font_size
    dl.font.bold = True
    dl.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


def render_planned_chart_slide(prs, spec):
    """Draw one AI/user-specified chart spec on a new SALIC content slide.
    Spec: {type, title, insight, unit, categories, series:[{name, values}]}."""
    ctype = spec.get("type", "column")
    xl_type = PLANNED_XL_TYPES.get(ctype)
    if xl_type is None:
        return None
    cats = spec.get("categories") or []
    series = [s for s in (spec.get("series") or []) if s.get("values")]
    if not cats or not series:
        return None
    if ctype in ("pie", "donut"):
        series = series[:1]

    num_fmt = _spec_number_format(spec)
    title = spec.get("title") or "Chart"
    unit = (spec.get("unit") or "").strip()
    if unit and unit.lower() not in title.lower():
        title += f" ({unit})"
    slide = _add_titled_slide(prs, title, spec.get("insight") or "", DARKNAVY)

    data = CategoryChartData(number_format=num_fmt)
    data.categories = [str(c) for c in cats]
    for s in series:
        data.add_series(str(s.get("name") or "Series"),
                        [v for v in s["values"]])

    if ctype in ("pie", "donut"):
        x, y, w, h = Inches(7.3), Inches(3.0), Inches(12.0), Inches(10.6)
    else:
        x, y, w, h = Inches(2.0), Inches(3.1), Inches(22.6), Inches(10.5)
    gf = slide.shapes.add_chart(xl_type, x, y, w, h, data)
    chart = gf.chart
    _style_chart(chart, base_size=Pt(14))

    n_ser, n_cat = len(series), len(cats)
    chart.has_legend = n_ser > 1 or ctype in ("pie", "donut")
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(15)
        chart.legend.font.color.rgb = DARKNAVY

    if ctype in ("pie", "donut"):
        # one colour per slice
        ser = chart.plots[0].series[0]
        for i in range(n_cat):
            point = ser.points[i]
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = SERIES_PALETTE[i % len(SERIES_PALETTE)]
        _pie_labels(chart)
        return slide

    all_vals = [v for s in series for v in s["values"] if v is not None]
    has_negative = any(v < 0 for v in all_vals)
    # a single-series bar/column with mixed signs is a bridge/variance chart:
    # colour each point by direction instead of one flat fill colour
    per_point_sign = (n_ser == 1 and ctype in ("column", "bar")
                       and has_negative and any(v >= 0 for v in all_vals))

    # series colours (fill for bars/areas, line colour for lines)
    for i, ser in enumerate(chart.plots[0].series):
        col = SERIES_PALETTE[i % len(SERIES_PALETTE)]
        if ctype == "line":
            ser.format.line.color.rgb = col
            ser.format.line.width = Pt(3)
            try:
                ser.marker.style = 8  # circle
                ser.marker.format.fill.solid()
                ser.marker.format.fill.fore_color.rgb = col
                ser.marker.format.line.color.rgb = col
            except Exception:
                pass
            ser.smooth = False
        elif per_point_sign:
            for point, v in zip(ser.points, series[i]["values"]):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = GREEN if v >= 0 else RED
        else:
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = col

    # data labels when the chart stays readable with them
    if ctype != "stacked_column" and n_ser * n_cat <= 16:
        pos = "t" if ctype == "line" else (None if ctype == "area" else "outEnd")
        for i, ser in enumerate(chart.plots[0].series):
            col = DARKNAVY if (n_ser == 1 or per_point_sign) else SERIES_PALETTE[i % len(SERIES_PALETTE)]
            _series_labels(ser, Pt(14) if n_ser * n_cat > 8 else Pt(16),
                           str(col), number_format=num_fmt, pos=pos)

    va = chart.value_axis
    va.tick_labels.number_format = num_fmt
    va.tick_labels.number_format_is_linked = False
    if ctype in ("column", "bar", "stacked_column", "area") and not has_negative:
        va.minimum_scale = 0  # bars must start at a zero baseline; skip when
        # values go negative (bridge/variance charts) so the axis can cross zero
    _style_axis(va, font_size=Pt(13))
    _style_axis(chart.category_axis, font_size=Pt(14), gridlines=False)
    _ensure_orientation(va)
    _ensure_orientation(chart.category_axis)
    if ctype == "stacked_column":
        _set_overlap(chart, 100)
    if ctype in ("column", "bar", "stacked_column"):
        _set_gap_width(chart, gap_pct=80 if n_ser == 1 else 120)
    return slide


# ----------------------------------------------------- chat report slides

def render_narrative_slide(prs, title, text):
    """One content slide with a single block of narrative text (AI chat answer)."""
    text = (text or "").strip()
    if not text:
        return None
    slide = _add_titled_slide(prs, title or "Insight", "")
    tb = slide.shapes.add_textbox(Inches(2.0), Inches(3.2), Inches(22.6), Inches(9.5))
    tf = tb.text_frame
    tf.word_wrap = True
    paras = [p.strip() for p in text.split("\n") if p.strip()] or [text]
    _set_lines(tf, [[(p, {"size": Pt(20), "color": DARKNAVY})] for p in paras])
    for para in tf.paragraphs:
        para.space_after = Pt(14)
        para.alignment = PP_ALIGN.LEFT
    return slide


def render_table_slide(prs, spec):
    """One content slide with a native pptx table (AI chat table answer).
    Spec: {title, columns:[...], rows:[[...], ...]}."""
    cols = [str(c) for c in (spec.get("columns") or [])]
    rows = [r for r in (spec.get("rows") or []) if isinstance(r, list)]
    if not cols or not rows:
        return None
    slide = _add_titled_slide(prs, spec.get("title") or "Table", "")

    n_rows, n_cols = len(rows) + 1, len(cols)
    x, y, w, h = Inches(2.0), Inches(3.3), Inches(22.6), min(Inches(0.7) * n_rows, Inches(10.2))
    gf = slide.shapes.add_table(n_rows, n_cols, x, y, w, h)
    table = gf.table

    for j, col in enumerate(cols):
        cell = table.cell(0, j)
        cell.text = str(col)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        for para in cell.text_frame.paragraphs:
            para.font.size = Pt(16)
            para.font.bold = True
            para.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for i, row in enumerate(rows, start=1):
        for j in range(n_cols):
            val = row[j] if j < len(row) else ""
            cell = table.cell(i, j)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0xF5, 0xF8, 0xFA) if i % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)
            for para in cell.text_frame.paragraphs:
                para.font.size = Pt(15)
                para.font.color.rgb = DARKNAVY
    return slide


def build_report_pptx(blocks, company="SALIC", period=""):
    """Assemble a SALIC-template deck from AI chat report blocks.
    blocks: list of {"type": "narrative"|"chart"|"table", ...}
      narrative -> {"title", "text"}
      chart     -> {"spec"}  (validated chart spec, see planner.py)
      table     -> {"spec"}  ({"title","columns","rows"})
    Independent of build_pptx / the voice-to-slide flow."""
    company = (company or "SALIC").strip()
    period = (period or "").strip()

    prs = Presentation(TEMPLATE)
    slides = list(prs.slides)

    title = f"{company} — AI Insights Report"
    if period:
        title += f" · {period}"

    _fill_cover(slides[COVER], title)

    report_slides = []
    for block in blocks or []:
        btype = (block or {}).get("type")
        try:
            if btype == "narrative":
                s = render_narrative_slide(prs, block.get("title") or "Insight", block.get("text") or "")
            elif btype == "chart":
                s = render_planned_chart_slide(prs, block.get("spec") or {})
            elif btype == "table":
                s = render_table_slide(prs, block.get("spec") or {})
            else:
                s = None
        except Exception as e:
            print(f"[pptx] report block skipped ({e})")
            s = None
        if s is not None:
            report_slides.append(s)

    ordered = [slides[COVER].slide_id]
    ordered += [s.slide_id for s in report_slides]
    ordered.append(slides[THANKYOU].slide_id)
    _arrange(prs, ordered)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


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
    inc_agenda = include.get("agenda", True) and AGENDA is not None
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
            s = maker(prs, kpis, unit, period)
            if s is not None:
                chart_slides.append(s)

    # AI-planned charts (validated specs from planner.py / the UI cards)
    for spec in (d.get("charts") or [])[:6]:
        try:
            s = render_planned_chart_slide(prs, spec)
        except Exception as e:
            print(f"[pptx] planned chart skipped ({e})")
            s = None
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
