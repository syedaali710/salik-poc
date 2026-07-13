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
import io, os, re
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
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
GOLD = RGBColor(0xC6, 0xA1, 0x5B)          # board-deck header accent
WINS_BG = RGBColor(0xE7, 0xF1, 0xEA)       # soft green panel
WATCH_BG = RGBColor(0xFB, 0xEC, 0xEC)      # soft red panel
CARD_LINE = RGBColor(0xE3, 0xE9, 0xEF)

# Board-deck chrome text (kept generic; period/company come from the request).
BOARD_EYEBROW = "GROUP FINANCIAL PERFORMANCE"
BOARD_TAG = "A PIF COMPANY"

# Client BoD deck (standard 16:9) — matches SALIC BoD Financial Performance template.
BOARD_TEMPLATE = os.path.join(HERE, "assets", "SALIC_board_template.pptx")
BOARD_LOGO_LEFT = os.path.join(HERE, "assets", "board_logo_left.png")
BOARD_LOGO_RIGHT = os.path.join(HERE, "assets", "board_logo_right.png")
BOARD_WINS_BG = RGBColor(0xE4, 0xF2, 0xE6)
BOARD_WATCH_BG = RGBColor(0xFB, 0xED, 0xEC)
BOARD_BRIDGE_RED = RGBColor(0xC0, 0x45, 0x3B)
BOARD_PANEL_BG = RGBColor(0xF2, 0xF6, 0xFA)
BOARD_KPI_X = [Inches(0.4), Inches(3.6), Inches(6.81), Inches(10.01)]

# Vertical rhythm for legacy widescreen template slides (26.67×15 in).
TITLE_TOP = Inches(1.25)
INSIGHT_TOP = Inches(2.15)
CONTENT_TOP = Inches(3.65)
CONTENT_H = Inches(10.8)

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


# ---------------------------------------------------------------- cell helpers

_NUM_CHARS = set("0123456789")


def _num_sign(text):
    """Classify a table cell for numeric alignment/coloring.
    -> ("neg"|"pos"|"num"|None). None = non-numeric (left-aligned, default color).
    Handles 1,234 / (1,234) / -12.3 / +5% / ▲ 12 / ▼ (3) / 2.05x etc."""
    s = str(text).strip()
    if not s or s in ("—", "-", "–"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    neg = neg or s.lstrip().startswith("-") or s.lstrip().startswith("−") or "▼" in s
    pos = s.lstrip().startswith("+") or "▲" in s
    digits = [c for c in s if c in _NUM_CHARS]
    if not digits:
        return None
    # mostly-numeric check: strip allowed formatting chars, rest must be digits
    core = s
    for ch in "(),+-−▲▼%x× .\u00a0'":
        core = core.replace(ch, "")
    if not core.isdigit():
        return None
    if neg:
        return "neg"
    if pos:
        return "pos"
    return "num"


def _is_total_row(row):
    """First cell names a subtotal/total line -> emphasize the whole row."""
    head = str((row or [""])[0]).strip().lower()
    return any(k in head for k in ("total", "consolidated", "group", "net income",
                                    "profit for the period", "grand"))


def _add_footer(slide, prs, company, period):
    """Bottom-left footer line: '<company> — Financial Performance · <period> · SAR',
    like the client board deck. Page numbers are added separately once the deck
    order is known (see _add_page_numbers)."""
    label = f"{company} — Financial Performance" if company else "Financial Performance"
    parts = [p for p in (label, period, "SAR") if p]
    if not parts:
        return
    y = prs.slide_height - Inches(0.85)
    tb = slide.shapes.add_textbox(Inches(1.3), y, Inches(18.0), Inches(0.5))
    tf = tb.text_frame
    tf.word_wrap = False
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.LEFT
    run = para.add_run()
    run.text = "   ·   ".join(parts)
    run.font.size = Pt(11)
    run.font.color.rgb = GREY


def _clear_template_shapes(slide):
    """Blank inherited layout placeholders and drop the template slide-number
    token (‹#›) that renders as a large corner digit alongside our footer."""
    for ph in slide.placeholders:
        _set_lines(ph.text_frame, [""])
    drop = []
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        txt = (sh.text_frame.text or "").strip()
        if "‹#›" in txt or txt in ("#", "‹#›"):
            drop.append(sh)
    for sh in drop:
        el = sh._element
        el.getparent().remove(el)


def _insight_style(text):
    """Pick font size / box height so long bridge insights wrap instead of clipping."""
    text = (text or "").strip()
    if len(text) > 160:
        return text, Pt(13), Inches(1.55)
    if len(text) > 100:
        return text, Pt(14), Inches(1.35)
    return text, Pt(16), Inches(1.0)


def _add_page_number(slide, prs, page_no, total):
    """Bottom-right 'Page X of Y' stamp for board content slides."""
    tb = slide.shapes.add_textbox(prs.slide_width - Inches(6.3),
                                  prs.slide_height - Inches(0.85),
                                  Inches(5.0), Inches(0.5))
    tf = tb.text_frame
    tf.word_wrap = False
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.RIGHT
    run = para.add_run()
    run.text = f"Page {page_no} of {total}"
    run.font.size = Pt(11)
    run.font.color.rgb = GREY


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

def _board_header(slide, prs):
    """Navy header band across the top of a board slide: company (left),
    'GROUP FINANCIAL PERFORMANCE' (center, gold), 'A PIF COMPANY' (right).
    Mirrors the client BoD deck's masthead."""
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1.0))
    band.fill.solid()
    band.fill.fore_color.rgb = DARKNAVY
    band.line.fill.background()
    band.shadow.inherit = False

    def _cell(x, w, text, color, size, align, bold=True):
        tb = slide.shapes.add_textbox(x, Inches(0.24), w, Inches(0.55))
        tf = tb.text_frame
        tf.word_wrap = False
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text = text
        r.font.size = size
        r.font.bold = bold
        r.font.color.rgb = color

    edge = Inches(1.3)
    third = int((prs.slide_width - 2 * edge) / 3)
    _cell(edge, third, "SALIC GROUP", RGBColor(0xFF, 0xFF, 0xFF), Pt(17), PP_ALIGN.LEFT)
    _cell(edge + third, third, BOARD_EYEBROW, GOLD, Pt(15), PP_ALIGN.CENTER)
    _cell(edge + 2 * third, third, BOARD_TAG, RGBColor(0xCF, 0xDD, 0xEC), Pt(12),
          PP_ALIGN.RIGHT)


def _add_titled_slide(prs, title, insight, insight_color=GREEN):
    """A board content slide: navy masthead band, section title, and an optional
    insight line — all drawn as our own shapes (the template's placeholders are
    cleared) so the layout matches the client's Board-of-Directors deck. Content
    renderers draw below CONTENT_TOP, clear of this chrome."""
    slide = prs.slides.add_slide(_content_layout(prs))
    _clear_template_shapes(slide)

    _board_header(slide, prs)

    tb = slide.shapes.add_textbox(Inches(1.3), TITLE_TOP, Inches(24.0), Inches(0.9))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    _set_lines(tf, [[(title, {"size": Pt(30), "bold": True, "color": DARKNAVY})]])
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT

    if insight:
        insight, size, box_h = _insight_style(insight)
        ib = slide.shapes.add_textbox(Inches(1.3), INSIGHT_TOP, Inches(24.0), box_h)
        itf = ib.text_frame
        itf.word_wrap = True
        itf.vertical_anchor = MSO_ANCHOR.TOP
        _set_lines(itf, [[(insight, {"size": size, "bold": True, "color": insight_color})]])
        itf.paragraphs[0].alignment = PP_ALIGN.LEFT
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


def _find_series(series, *patterns):
    """Return the values list of the first series whose name matches any pattern."""
    for s in series:
        name = str(s.get("name") or "").lower()
        if any(p in name for p in patterns):
            return s.get("values") or []
    return None


def _looks_like_kpi_compare(cats, series):
    """True when the spec is a multi-KPI Actual vs Budget/Prior-Year comparison —
    the case that reads terribly as one shared-axis chart (Revenue dwarfs Net
    Income) and belongs as KPI cards, like the board deck's summary page."""
    if len(cats) < 2:
        return False
    actual = _find_series(series, "actual")
    ref = _find_series(series, "budget", "prior", "py", "last year", "smly")
    return actual is not None and ref is not None


def _val_at(values, i):
    if values is None or i >= len(values):
        return None
    v = values[i]
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def render_kpi_cards_slide(prs, spec, cats, series, company="", period=""):
    """Board-style KPI cards: one card per KPI with the headline Actual number
    and vs-Budget / vs-PY deltas (▲/▼, green/red). Each KPI is self-scaled, so
    small metrics stay legible next to large ones — unlike a shared-axis chart."""
    actual = _find_series(series, "actual")
    budget = _find_series(series, "budget")
    py = _find_series(series, "prior", "py", "last year", "smly")
    unit = (spec.get("unit") or "").strip()

    title = spec.get("title") or "KPI Summary"
    slide = _add_titled_slide(prs, title, spec.get("insight") or "", DARKNAVY)

    n = len(cats)
    cols = n if n <= 5 else (n + 1) // 2
    rows = 1 if n <= 5 else 2
    margin, gutter = Inches(1.3), Inches(0.5)
    top = CONTENT_TOP
    card_h = Inches(4.2) if rows == 1 else Inches(3.6)
    row_gap = Inches(0.6)
    total_w = prs.slide_width - 2 * margin - gutter * (cols - 1)
    card_w = int(total_w / cols)

    for i, label in enumerate(cats):
        r, c = divmod(i, cols)
        x = margin + c * (card_w + gutter)
        y = top + r * (card_h + row_gap)

        body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, card_h)
        body.fill.solid()
        body.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        body.line.color.rgb = RGBColor(0xE3, 0xE9, 0xEF)
        body.line.width = Pt(1)
        body.shadow.inherit = False

        # navy accent bar across the top of the card
        accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, card_w, Inches(0.22))
        accent.fill.solid()
        accent.fill.fore_color.rgb = NAVY
        accent.line.fill.background()
        accent.shadow.inherit = False

        a = _val_at(actual, i)
        b = _val_at(budget, i)
        p = _val_at(py, i)

        lines = [[(str(label).upper(), {"size": Pt(13), "bold": True, "color": GREY})]]
        big = fmt(a) if a is not None else "—"
        lines.append([(big, {"size": Pt(32), "bold": True, "color": NAVY}),
                      (f"  {unit}" if unit else "", {"size": Pt(14), "color": GREY})])
        for ref_val, ref_label in ((b, "Budget"), (p, "PY")):
            if a is not None and ref_val is not None:
                pct = _variance_pct(a, ref_val)
                up = (a - ref_val) >= 0
                arrow = "▲" if up else "▼"
                sign = "+" if up else "-"
                col = GREEN if up else RED
                txt = (f"{arrow} vs {ref_label} {sign}{abs(pct):.1f}%"
                       if pct is not None else f"{arrow} vs {ref_label}")
                lines.append([(txt, {"size": Pt(13), "bold": True, "color": col})])
        foot = []
        if b is not None:
            foot.append(f"Budget {fmt(b)}")
        if p is not None:
            foot.append(f"PY {fmt(p)}")
        if foot:
            lines.append([("   ·   ".join(foot), {"size": Pt(11), "color": GREY})])

        tf = body.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.4)
        tf.margin_right = Inches(0.4)
        tf.margin_top = Inches(0.5)
        _set_lines(tf, lines)
        for para in tf.paragraphs:
            para.alignment = PP_ALIGN.LEFT
            para.space_after = Pt(6)

    _add_footer(slide, prs, company, period)
    return slide


_BRIDGE_START = {"b", "budget", "start", "baseline", "opening", "prior year", "py"}
_BRIDGE_END = {"a", "actual", "end", "closing", "total", "reported"}


def _looks_like_bridge(cats, series):
    """A single-series chart that goes start-total → signed deltas → end-total.
    Detected so it can render as a floating waterfall instead of flat columns."""
    if len(series) != 1 or len(cats) < 3:
        return False
    first = str(cats[0]).strip().lower()
    last = str(cats[-1]).strip().lower()
    if first not in _BRIDGE_START or last not in _BRIDGE_END:
        return False
    mids = series[0]["values"][1:-1]
    return any(v is not None and v < 0 for v in mids)


def render_waterfall_slide(prs, spec, company="", period=""):
    """Render a bridge spec as a true waterfall: the first and last categories
    are full total bars (navy); the middle steps float — green for increases,
    red for decreases — each starting where the previous left off. Built with a
    stacked column and an invisible base series (python-pptx has no native
    waterfall)."""
    cats = spec.get("categories") or []
    series = spec.get("series") or []
    if not cats or not series:
        return None
    raw = series[0].get("values") or []
    if len(raw) != len(cats) or len(cats) < 3:
        return None
    vals = [float(v) if v is not None else 0.0 for v in raw]
    n = len(vals)

    unit = (spec.get("unit") or "").strip()
    title = spec.get("title") or "Bridge"
    if unit and unit.lower() not in title.lower():
        title += f" ({unit})"
    slide = _add_titled_slide(prs, title, spec.get("insight") or "", DARKNAVY)

    base, total, inc, dec = [], [], [], []
    running = 0.0
    for i, v in enumerate(vals):
        if i == 0 or i == n - 1:  # start / end totals
            base.append(0.0)
            total.append(v)
            inc.append(None)
            dec.append(None)
            running = v
        elif v >= 0:
            base.append(running)
            inc.append(v)
            total.append(None)
            dec.append(None)
            running += v
        else:
            running += v
            base.append(running)
            dec.append(-v)  # stored positive; forced negative via number format
            total.append(None)
            inc.append(None)

    data = CategoryChartData(number_format="#,##0;(#,##0)")
    data.categories = [str(c) for c in cats]
    data.add_series("base", base)
    data.add_series("total", total)
    data.add_series("increase", inc)
    data.add_series("decrease", dec)

    gf = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED,
        Inches(2.0), CONTENT_TOP, Inches(22.6), CONTENT_H, data,
    )
    chart = gf.chart
    _style_chart(chart, base_size=Pt(14))
    chart.has_legend = False

    plot_series = chart.plots[0].series
    # base is the invisible spacer that lifts each floating bar to its level
    plot_series[0].format.fill.background()
    plot_series[0].format.line.fill.background()
    for idx, col in ((1, NAVY), (2, GREEN), (3, RED)):
        plot_series[idx].format.fill.solid()
        plot_series[idx].format.fill.fore_color.rgb = col

    _series_labels(plot_series[1], Pt(14), "FFFFFF", number_format="#,##0", pos="ctr")
    _series_labels(plot_series[2], Pt(14), str(GREEN), number_format='"+"#,##0', pos="ctr")
    _series_labels(plot_series[3], Pt(14), str(RED), number_format='"-"#,##0', pos="ctr")

    va = chart.value_axis
    va.minimum_scale = 0
    va.tick_labels.number_format = "#,##0"
    va.tick_labels.number_format_is_linked = False
    _style_axis(va, font_size=Pt(13))
    _style_axis(chart.category_axis, font_size=Pt(14), gridlines=False)
    _ensure_orientation(va)
    _ensure_orientation(chart.category_axis)
    _set_overlap(chart, 100)
    _set_gap_width(chart, gap_pct=60)

    _add_footer(slide, prs, company, period)
    return slide


def render_planned_chart_slide(prs, spec, company="", period=""):
    """Draw one AI/user-specified chart spec on a new SALIC content slide.
    Spec: {type, title, insight, unit, categories, series:[{name, values}]}.
    Multi-KPI Actual/Budget/PY comparisons are routed to KPI cards instead;
    start→delta→end bridge specs render as floating waterfalls."""
    ctype = spec.get("type", "column")
    xl_type = PLANNED_XL_TYPES.get(ctype)
    if xl_type is None:
        return None
    cats = spec.get("categories") or []
    series = [s for s in (spec.get("series") or []) if s.get("values")]
    if not cats or not series:
        return None

    # Bridge / waterfall: start total → signed deltas → end total. Render as
    # floating up/down bars (tagged variant, or detected by shape).
    if ctype in ("column", "bar", "stacked_column") and (
        spec.get("variant") == "waterfall" or _looks_like_bridge(cats, series)
    ):
        wf_slide = render_waterfall_slide(prs, spec, company, period)
        if wf_slide is not None:
            return wf_slide

    # A multi-KPI Actual vs Budget/PY grid reads badly on one shared axis
    # (Revenue dwarfs Net Income); render it as self-scaled KPI cards instead.
    if ctype in ("column", "bar", "stacked_column") and _looks_like_kpi_compare(cats, series):
        card_slide = render_kpi_cards_slide(prs, spec, cats, series, company, period)
        if card_slide is not None:
            return card_slide

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
        x, y, w, h = Inches(7.3), CONTENT_TOP, Inches(12.0), CONTENT_H
    else:
        x, y, w, h = Inches(2.0), CONTENT_TOP, Inches(22.6), CONTENT_H
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
        _add_footer(slide, prs, company, period)
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
    _add_footer(slide, prs, company, period)
    return slide


# ----------------------------------------------------- chat report slides

def _split_bullets(text):
    """Turn a CFO-memo paragraph into board-style bullets. Respect explicit
    line breaks; otherwise split into sentences so a dense paragraph becomes a
    scannable 'story in one view' list."""
    text = (text or "").strip()
    lines = [l.strip(" •-\t").strip() for l in text.split("\n") if l.strip()]
    if len(lines) > 1:
        return lines
    # single block -> sentence split (keep decimals/abbreviations intact)
    parts = re.split(r"(?<=[.!?])\s+(?=[A-Z0-9])", text)
    parts = [p.strip() for p in parts if p.strip()]
    return parts or [text]


def _strip_prefix(text):
    """Drop a leading 'Wins:' / 'Watch:' / 'Watch-point:' label from a bullet."""
    return re.sub(r"^\s*(wins?|watch[- ]?points?|watch|risks?)\s*[:\-–]\s*", "",
                  text, flags=re.IGNORECASE).strip()


def _wins_watch_split(bullets):
    """Split bullets into (wins, watch) by their leading label. Returns None if
    the content isn't clearly a wins/watch list."""
    wins, watch = [], []
    for b in bullets:
        low = b.lower()
        if low.startswith("win"):
            wins.append(_strip_prefix(b))
        elif low.startswith(("watch", "risk")):
            watch.append(_strip_prefix(b))
    if wins and watch:
        return wins, watch
    return None


def _panel(slide, x, y, w, h, header, header_color, bg, bullets):
    """A soft-tinted rounded panel with a coloured header and bullet lines —
    the wins / watch-points boxes from the board deck."""
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    box.fill.solid()
    box.fill.fore_color.rgb = bg
    box.line.color.rgb = header_color
    box.line.width = Pt(1)
    box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.45)
    tf.margin_right = Inches(0.45)
    tf.margin_top = Inches(0.4)
    lines = [[(header, {"size": Pt(20), "bold": True, "color": header_color})]]
    lines += [[("•  ", {"size": Pt(15), "bold": True, "color": header_color}),
               (b, {"size": Pt(15), "color": DARKNAVY})] for b in bullets]
    _set_lines(tf, lines)
    for para in tf.paragraphs:
        para.alignment = PP_ALIGN.LEFT
        para.space_after = Pt(10)
        para.line_spacing = 1.1


def render_wins_watch_slide(prs, title, insight, wins, watch, company="", period=""):
    """Key Drivers & Risks rendered as side-by-side green wins / red watch-point
    panels, like the board deck's 'Five wins / three watch-points' page."""
    slide = _add_titled_slide(prs, title or "Key Drivers & Risks", insight, DARKNAVY)
    margin, gap = Inches(1.3), Inches(0.6)
    top, height = CONTENT_TOP, CONTENT_H
    panel_w = int((prs.slide_width - 2 * margin - gap) / 2)
    _panel(slide, margin, top, panel_w, height,
           "▲  Wins", GREEN, WINS_BG, wins)
    _panel(slide, margin + panel_w + gap, top, panel_w, height,
           "▼  Watch-points", RED, WATCH_BG, watch)
    _add_footer(slide, prs, company, period)
    return slide


def render_narrative_slide(prs, title, text, insight="", company="", period=""):
    """One content slide with the AI chat answer typeset as a board-style
    executive summary: eyebrow insight, then accented bullet points. When the
    content is a wins/watch-points list, it renders as two coloured panels."""
    text = (text or "").strip()
    if not text:
        return None
    bullets = _split_bullets(text)
    # lead sentence doubles as the eyebrow insight when none was supplied
    eyebrow = (insight or "").strip()
    if not eyebrow and len(bullets) > 1 and _wins_watch_split(bullets) is None:
        eyebrow = bullets[0]
        bullets = bullets[1:]

    ww = _wins_watch_split(bullets)
    if ww is not None:
        return render_wins_watch_slide(prs, title or "Key Drivers & Risks",
                                       eyebrow, ww[0], ww[1], company, period)

    slide = _add_titled_slide(prs, title or "Insight", eyebrow, DARKNAVY)
    tb = slide.shapes.add_textbox(Inches(2.0), CONTENT_TOP, Inches(22.6), CONTENT_H)
    tf = tb.text_frame
    tf.word_wrap = True
    if len(bullets) > 1:
        lines = [[("▪  ", {"size": Pt(18), "bold": True, "color": GREEN}),
                  (b, {"size": Pt(19), "color": DARKNAVY})] for b in bullets]
    else:
        lines = [[(bullets[0], {"size": Pt(20), "color": DARKNAVY})]]
    _set_lines(tf, lines)
    for para in tf.paragraphs:
        para.space_after = Pt(16)
        para.line_spacing = 1.15
        para.alignment = PP_ALIGN.LEFT

    _add_footer(slide, prs, company, period)
    return slide


def _col_is_numeric(rows, j):
    """A column reads as numeric if most of its populated cells parse as numbers.
    Used to right-align the whole column (headers included) like a financial table."""
    seen = filled = 0
    for r in rows:
        if j < len(r) and str(r[j]).strip():
            seen += 1
            if _num_sign(r[j]) is not None:
                filled += 1
    return seen > 0 and filled >= max(1, seen // 2)


def render_table_slide(prs, spec, company="", period=""):
    """One content slide with a native pptx table styled like the board deck:
    right-aligned numeric columns, red/green variance cells, emphasized total
    rows, alternating zebra fill. Spec: {title, columns:[...], rows:[[...], ...]}."""
    cols = [str(c) for c in (spec.get("columns") or [])]
    rows = [r for r in (spec.get("rows") or []) if isinstance(r, list)]
    if not cols or not rows:
        return None
    slide = _add_titled_slide(prs, spec.get("title") or "Table", spec.get("insight") or "")

    n_rows, n_cols = len(rows) + 1, len(cols)
    row_h = Inches(0.62)
    x, y, w = Inches(2.0), CONTENT_TOP, Inches(22.6)
    h = min(row_h * n_rows, CONTENT_H)
    gf = slide.shapes.add_table(n_rows, n_cols, x, y, w, h)
    table = gf.table

    # first column carries labels (wide); numeric columns share the rest evenly
    if n_cols > 1:
        label_w = int(w * 0.34)
        rest = int((w - label_w) / (n_cols - 1))
        table.columns[0].width = label_w
        for j in range(1, n_cols):
            table.columns[j].width = rest

    numeric_cols = {j: _col_is_numeric(rows, j) for j in range(n_cols)}

    for j, col in enumerate(cols):
        cell = table.cell(0, j)
        cell.text = str(col)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        for para in cell.text_frame.paragraphs:
            para.alignment = PP_ALIGN.RIGHT if (j > 0 and numeric_cols.get(j)) else PP_ALIGN.LEFT
            para.font.size = Pt(15)
            para.font.bold = True
            para.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for i, row in enumerate(rows, start=1):
        total = _is_total_row(row)
        for j in range(n_cols):
            val = row[j] if j < len(row) else ""
            cell = table.cell(i, j)
            cell.text = str(val)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if total:
                cell.fill.fore_color.rgb = RGBColor(0xE9, 0xF0, 0xF6)
            else:
                cell.fill.fore_color.rgb = RGBColor(0xF5, 0xF8, 0xFA) if i % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)

            sign = _num_sign(val)
            color = DARKNAVY
            # colour variance-style cells (not the first label column)
            if j > 0 and sign == "neg":
                color = RED
            elif j > 0 and sign == "pos":
                color = GREEN
            for para in cell.text_frame.paragraphs:
                para.alignment = PP_ALIGN.RIGHT if (j > 0 and numeric_cols.get(j)) else PP_ALIGN.LEFT
                para.font.size = Pt(14)
                para.font.bold = bool(total)
                para.font.color.rgb = color

    _add_footer(slide, prs, company, period)
    return slide


# ---------------------------------------------------------------- client BoD deck (13.33×7.5 in)


def _new_board_prs():
    """Standard 16:9 board deck — same dimensions as the client's BoD PPTX."""
    prs = Presentation(BOARD_TEMPLATE) if os.path.isfile(BOARD_TEMPLATE) else Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    # Drop template demo slides; we draw every page from scratch.
    while len(prs.slides):
        rId = prs.slides._sldIdLst[0].get(qn("r:id"))
        prs.part.drop_rel(rId)
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])
    return prs


def _board_text(slide, x, y, w, h, text, size, bold=False, color=DARKNAVY,
                align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    _set_lines(tf, [[(text, {"size": size, "bold": bold, "color": color})]])
    tf.paragraphs[0].alignment = align
    return tb


def _fmt_board_amount(n):
    """Client-style headline amounts: '2.40 b' for billions, '687 m' for millions."""
    if n is None:
        return "—"
    try:
        n = float(n)
    except (ValueError, TypeError):
        return "—"
    neg = n < 0
    abs_n = abs(n)
    if abs_n >= 1000:
        s = f"{abs_n / 1000:.2f}".rstrip("0").rstrip(".")
        val = f"{s} b"
    else:
        val = f"{fmt(abs_n)} m"
    return f"({val})" if neg else val


def _board_shell(prs, title, subtitle="", period="", page_no=1, total=1):
    """Navy masthead, logos, title block, and three-part footer like the client deck."""
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    sw = prs.slide_width

    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, sw, Inches(0.9))
    band.fill.solid()
    band.fill.fore_color.rgb = DARKNAVY
    band.line.fill.background()

    if os.path.isfile(BOARD_LOGO_LEFT):
        slide.shapes.add_picture(BOARD_LOGO_LEFT, Inches(0.42), Inches(0.23),
                                 Inches(1.11), Inches(0.43))
    if os.path.isfile(BOARD_LOGO_RIGHT):
        slide.shapes.add_picture(BOARD_LOGO_RIGHT, Inches(11.71), Inches(0.28),
                                 Inches(1.04), Inches(0.34))

    _board_text(slide, Inches(4.17), Inches(0.3), Inches(5.0), Inches(0.32),
                "SALIC GROUP  ·  BOARD OF DIRECTORS", Pt(10.5), bold=True,
                color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    _board_text(slide, Inches(0.4), Inches(1.0), Inches(11.0), Inches(0.5),
                title, Pt(23), bold=True)
    if subtitle:
        _board_text(slide, Inches(0.4), Inches(1.5), Inches(11.6), Inches(0.3),
                    subtitle, Pt(11.5))

    sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(7.14),
                                 Inches(12.53), Pt(0.5))
    sep.fill.background()
    sep.line.color.rgb = RGBColor(0xE3, 0xE9, 0xEF)
    sep.line.width = Pt(0.5)

    _board_text(slide, Inches(0.4), Inches(7.18), Inches(5.0), Inches(0.22),
                "SALIC Group — Board Financial Performance", Pt(8), color=GREY)
    mid = f"{period}  ·  SAR" if period else "SAR"
    _board_text(slide, Inches(4.67), Inches(7.18), Inches(4.0), Inches(0.22),
                mid, Pt(8), color=GREY, align=PP_ALIGN.CENTER)
    _board_text(slide, Inches(10.93), Inches(7.18), Inches(2.0), Inches(0.22),
                f"Page {page_no} of {total}", Pt(8), color=GREY, align=PP_ALIGN.RIGHT)
    return slide


def _board_kpi_card(slide, x, y, label, actual, budget, py):
    """One KPI tile — white box, 25pt headline, split vs-Budget / vs-PY columns."""
    w, h = Inches(2.92), Inches(1.5)
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    box.line.color.rgb = RGBColor(0xE3, 0xE9, 0xEF)
    box.line.width = Pt(0.75)

    _board_text(slide, x + Inches(0.16), y + Inches(0.11), Inches(2.62), Inches(0.24),
                str(label).upper(), Pt(10), bold=True, color=GREY)
    _board_text(slide, x + Inches(0.15), y + Inches(0.31), Inches(2.62), Inches(0.48),
                _fmt_board_amount(actual), Pt(25), bold=True)

    def _var_col(col_x, ref_label, ref_val):
        if actual is None or ref_val is None:
            return
        pct = _variance_pct(actual, ref_val)
        up = (actual - ref_val) >= 0
        arrow = "▲" if up else "▼"
        sign = "+" if up else "−"
        col = GREEN if up else BOARD_BRIDGE_RED
        var = f"{arrow}{sign}{abs(pct):.1f}%" if pct is not None else arrow
        tb = slide.shapes.add_textbox(col_x, y + Inches(0.86), Inches(1.36), Inches(0.2))
        _set_lines(tb.text_frame, [
            [(f"vs {ref_label}  ", {"size": Pt(8), "color": GREY}),
             (var, {"size": Pt(8.5), "bold": True, "color": col})],
        ])
        _board_text(slide, col_x, y + Inches(1.06), Inches(1.36), Inches(0.2),
                    f"{ref_label} {_fmt_board_amount(ref_val)}", Pt(8), color=GREY)

    _var_col(x + Inches(0.16), "Budget", budget)
    _var_col(x + Inches(1.48), "PY", py)


def _board_kpi_row(slide, spec):
    cats = spec.get("categories") or []
    series = spec.get("series") or []
    actual = _find_series(series, "actual")
    budget = _find_series(series, "budget")
    py = _find_series(series, "prior", "py", "last year", "smly")
    for i, label in enumerate(cats[:4]):
        if i >= len(BOARD_KPI_X):
            break
        _board_kpi_card(slide, BOARD_KPI_X[i], Inches(1.92), label,
                        _val_at(actual, i), _val_at(budget, i), _val_at(py, i))


def _board_story(slide, text):
    bullets = _split_bullets(text)
    if not bullets:
        return
    _board_text(slide, Inches(0.4), Inches(3.58), Inches(6.5), Inches(0.3),
                "The story in one view", Pt(14), bold=True)
    tb = slide.shapes.add_textbox(Inches(0.4), Inches(3.9), Inches(6.5), Inches(3.0))
    tf = tb.text_frame
    tf.word_wrap = True
    lines = [[(b, {"size": Pt(11), "color": DARKNAVY})] for b in bullets[:5]]
    _set_lines(tf, lines)
    for para in tf.paragraphs:
        para.space_after = Pt(6)
        para.line_spacing = 1.1


def _board_insight_panel(slide, insight):
    """Right-hand callout when we lack reported→recurring data."""
    if not (insight or "").strip():
        return
    panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.15), Inches(3.58),
                                   Inches(5.78), Inches(3.34))
    panel.fill.solid()
    panel.fill.fore_color.rgb = BOARD_PANEL_BG
    panel.line.color.rgb = RGBColor(0xE3, 0xE9, 0xEF)
    panel.line.width = Pt(0.75)
    _board_text(slide, Inches(7.39), Inches(3.74), Inches(5.3), Inches(0.3),
                "Key insight", Pt(13), bold=True)
    tb = slide.shapes.add_textbox(Inches(7.39), Inches(4.1), Inches(5.3), Inches(2.6))
    tf = tb.text_frame
    tf.word_wrap = True
    _set_lines(tf, [[(insight.strip(), {"size": Pt(10.5), "color": DARKNAVY})]])


def _board_table(slide, spec, x, y, w, max_h, header_pt=10, cell_pt=9):
    cols = [str(c) for c in (spec.get("columns") or [])]
    rows = [r for r in (spec.get("rows") or []) if isinstance(r, list)]
    if not cols or not rows:
        return
    n_rows, n_cols = len(rows) + 1, len(cols)
    row_h = Inches(0.36)
    h = min(row_h * n_rows, max_h)
    gf = slide.shapes.add_table(n_rows, n_cols, x, y, w, h)
    table = gf.table
    if n_cols > 1:
        label_w = int(w * 0.28)
        rest = int((w - label_w) / (n_cols - 1))
        table.columns[0].width = label_w
        for j in range(1, n_cols):
            table.columns[j].width = rest

    numeric_cols = {j: _col_is_numeric(rows, j) for j in range(n_cols)}
    for j, col in enumerate(cols):
        cell = table.cell(0, j)
        cell.text = str(col)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        for para in cell.text_frame.paragraphs:
            para.alignment = PP_ALIGN.RIGHT if (j > 0 and numeric_cols.get(j)) else PP_ALIGN.LEFT
            para.font.size = Pt(header_pt)
            para.font.bold = True
            para.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for i, row in enumerate(rows, start=1):
        total = _is_total_row(row)
        for j in range(n_cols):
            val = row[j] if j < len(row) else ""
            cell = table.cell(i, j)
            cell.text = str(val)
            cell.fill.solid()
            if total:
                cell.fill.fore_color.rgb = RGBColor(0xE9, 0xF0, 0xF6)
            else:
                cell.fill.fore_color.rgb = (RGBColor(0xF5, 0xF8, 0xFA)
                                            if i % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF))
            sign = _num_sign(val)
            color = DARKNAVY
            if j > 0 and sign == "neg":
                color = BOARD_BRIDGE_RED
            elif j > 0 and sign == "pos":
                color = GREEN
            for para in cell.text_frame.paragraphs:
                para.alignment = PP_ALIGN.RIGHT if (j > 0 and numeric_cols.get(j)) else PP_ALIGN.LEFT
                para.font.size = Pt(cell_pt)
                para.font.bold = bool(total)
                para.font.color.rgb = color


def _board_waterfall(slide, spec, x, y, w, h, label_size=Pt(8.5), axis_size=Pt(8)):
    cats = spec.get("categories") or []
    series = spec.get("series") or []
    if not cats or not series:
        return
    raw = series[0].get("values") or []
    if len(raw) != len(cats) or len(cats) < 3:
        return
    vals = [float(v) if v is not None else 0.0 for v in raw]
    n = len(vals)

    base, total, inc, dec = [], [], [], []
    running = 0.0
    for i, v in enumerate(vals):
        if i == 0 or i == n - 1:
            base.append(0.0)
            total.append(v)
            inc.append(None)
            dec.append(None)
            running = v
        elif v >= 0:
            base.append(running)
            inc.append(v)
            total.append(None)
            dec.append(None)
            running += v
        else:
            running += v
            base.append(running)
            dec.append(-v)
            total.append(None)
            inc.append(None)

    data = CategoryChartData(number_format="#,##0;(#,##0)")
    data.categories = [str(c) for c in cats]
    data.add_series("base", base)
    data.add_series("total", total)
    data.add_series("increase", inc)
    data.add_series("decrease", dec)

    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, x, y, w, h, data)
    chart = gf.chart
    _style_chart(chart, base_size=axis_size)
    chart.has_legend = False
    plot_series = chart.plots[0].series
    plot_series[0].format.fill.background()
    plot_series[0].format.line.fill.background()
    for idx, col in ((1, NAVY), (2, GREEN), (3, BOARD_BRIDGE_RED)):
        plot_series[idx].format.fill.solid()
        plot_series[idx].format.fill.fore_color.rgb = col

    _series_labels(plot_series[1], label_size, "FFFFFF", number_format="#,##0", pos="ctr")
    _series_labels(plot_series[2], label_size, str(GREEN), number_format='"+"#,##0', pos="ctr")
    _series_labels(plot_series[3], label_size, str(BOARD_BRIDGE_RED),
                   number_format='"-"#,##0', pos="ctr")

    va = chart.value_axis
    va.minimum_scale = 0
    va.tick_labels.number_format = "#,##0"
    va.tick_labels.number_format_is_linked = False
    _style_axis(va, font_size=axis_size)
    _style_axis(chart.category_axis, font_size=axis_size, gridlines=False)
    _ensure_orientation(va)
    _ensure_orientation(chart.category_axis)
    _set_overlap(chart, 100)
    _set_gap_width(chart, gap_pct=50)


def _board_wins_watch(slide, wins, watch):
    """Compact side-by-side panels matching client slide 4."""
    for x, header, bullets, bg, col in (
        (Inches(0.4), "▲  Wins", wins, BOARD_WINS_BG, GREEN),
        (Inches(6.87), "▼  Watch-points", watch, BOARD_WATCH_BG, BOARD_BRIDGE_RED),
    ):
        if not bullets:
            continue
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(1.95),
                                     Inches(6.07), Inches(2.62))
        box.fill.solid()
        box.fill.fore_color.rgb = bg
        box.line.color.rgb = col
        box.line.width = Pt(0.75)
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.2)
        tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.12)
        lines = [[(header, {"size": Pt(13), "bold": True, "color": col})]]
        lines += [[("•  ", {"size": Pt(10), "bold": True, "color": col}),
                   (b, {"size": Pt(10), "color": DARKNAVY})] for b in bullets[:4]]
        _set_lines(tf, lines)
        for para in tf.paragraphs:
            para.space_after = Pt(4)
            para.line_spacing = 1.05


def _classify_report_blocks(blocks):
    """Sort LLM blocks into exec / KPI / bridges / drivers / matters."""
    out = {"exec": None, "kpi_chart": None, "kpi_table": None,
           "bridges": [], "drivers": None, "matters": None, "insight": ""}
    for block in blocks or []:
        btype = (block or {}).get("type")
        if btype == "narrative":
            title = str(block.get("title") or "").lower()
            text = block.get("text") or ""
            if "driver" in title or "risk" in title:
                out["drivers"] = block
            elif "executive" in title:
                out["exec"] = block
            elif out["exec"] is None and "wins:" not in text.lower():
                out["exec"] = block
        elif btype == "chart":
            spec = block.get("spec") or {}
            cats = spec.get("categories") or []
            series = spec.get("series") or []
            if _looks_like_kpi_compare(cats, series):
                out["kpi_chart"] = block
            elif spec.get("variant") == "waterfall" or _looks_like_bridge(cats, series):
                out["bridges"].append(block)
            else:
                out["bridges"].append(block)
            if spec.get("insight") and not out["insight"]:
                out["insight"] = spec["insight"]
        elif btype == "table":
            title = str((block.get("spec") or {}).get("title") or "").lower()
            if "matter" in title or "board" in title:
                out["matters"] = block
            else:
                out["kpi_table"] = block
    return out


def _build_board_deck(blocks, company="SALIC", period=""):
    """Assemble a client-format 4-page board deck (13.33×7.5 in)."""
    parts = _classify_report_blocks(blocks)
    bridges = parts["bridges"]
    n_bridge_pages = max(1, (len(bridges) + 1) // 2) if bridges else 0
    total = 2 + n_bridge_pages + 1
    page = 1
    period_label = period or "Year-to-Date"

    # --- Page 1: Executive summary + KPI row + story ---
    sub1 = f"Consolidated results, {period_label} · Actual vs Budget vs Prior Year"
    s1 = _board_shell(prs := _new_board_prs(),
                      "Group Financial Performance — Executive Summary",
                      sub1, period_label, page, total)
    if parts["kpi_chart"]:
        _board_kpi_row(s1, parts["kpi_chart"]["spec"])
    if parts["exec"]:
        _board_story(s1, parts["exec"].get("text") or "")
    elif parts["drivers"]:
        _board_story(s1, parts["drivers"].get("text") or "")
    insight = parts["insight"]
    if parts["kpi_chart"] and parts["kpi_chart"]["spec"].get("insight"):
        insight = parts["kpi_chart"]["spec"]["insight"]
    _board_insight_panel(s1, insight)
    page += 1

    # --- Page 2: KPI table + first bridge ---
    s2 = _board_shell(prs, "Consolidated Performance & Variance Bridges",
                      "KPI summary and the main variance bridge vs budget",
                      period_label, page, total)
    if parts["kpi_table"]:
        _board_table(s2, parts["kpi_table"]["spec"],
                     Inches(0.4), Inches(1.95), Inches(6.35), Inches(3.6),
                     header_pt=9, cell_pt=8.5)
    if bridges:
        spec = bridges[0]["spec"]
        title = spec.get("title") or "Variance bridge"
        unit = (spec.get("unit") or "").strip()
        if unit and unit.lower() not in title.lower():
            title += f"  ({unit})"
        _board_text(s2, Inches(7.0), Inches(1.95), Inches(5.93), Inches(0.28),
                    title, Pt(12), bold=True)
        _board_waterfall(s2, spec, Inches(7.0), Inches(2.25), Inches(5.93), Inches(3.95))
        foot = (spec.get("insight") or "").strip()
        if foot:
            _board_text(s2, Inches(0.4), Inches(5.76), Inches(12.5), Inches(0.4),
                        foot, Pt(9), color=GREY)
    page += 1

    # --- Page 3+: remaining bridges (up to 2 per page) ---
    rest = bridges[1:]
    while rest:
        chunk = rest[:2]
        rest = rest[2:]
        s = _board_shell(prs, "Variance Bridges",
                         "Additional budget-to-actual drivers",
                         period_label, page, total)
        if len(chunk) == 1:
            spec = chunk[0]["spec"]
            title = spec.get("title") or "Bridge"
            _board_text(s, Inches(0.4), Inches(1.92), Inches(12.5), Inches(0.28),
                        title, Pt(12), bold=True)
            _board_waterfall(s, spec, Inches(0.4), Inches(2.25), Inches(12.53), Inches(4.5))
        else:
            for idx, blk in enumerate(chunk):
                spec = blk["spec"]
                top = Inches(1.95) if idx == 0 else Inches(4.35)
                title = spec.get("title") or "Bridge"
                _board_text(s, Inches(0.4), top, Inches(12.5), Inches(0.25),
                            title, Pt(11), bold=True)
                _board_waterfall(s, spec, Inches(0.4), top + Inches(0.32),
                                 Inches(12.53), Inches(2.0))
        page += 1

    # --- Last page: drivers + matters ---
    wins, watch = [], []
    if parts["drivers"]:
        ww = _wins_watch_split(_split_bullets(parts["drivers"].get("text") or ""))
        if ww:
            wins, watch = ww
    n_wins, n_watch = len(wins), len(watch)
    sub4 = f"{n_wins} wins, {n_watch} watch-points, and matters for the Board"
    s4 = _board_shell(prs, "Key Drivers, Risks & Matters for the Board",
                      sub4, period_label, page, total)
    _board_wins_watch(s4, wins, watch)
    if parts["matters"]:
        _board_text(s4, Inches(0.4), Inches(5.62), Inches(6.0), Inches(0.3),
                    "Matters for the Board", Pt(12), bold=True)
        _board_table(s4, parts["matters"]["spec"],
                     Inches(0.4), Inches(5.94), Inches(12.53), Inches(1.1),
                     header_pt=9, cell_pt=8.5)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


def build_report_pptx(blocks, company="SALIC", period=""):
    """Assemble a client-format board deck from AI report blocks.

    Renders at standard 16:9 (13.33×7.5 in) to match the client's
    SALIC BoD Financial Performance deck — compact KPI tiles, dense tables,
    side-by-side bridges, and wins/watch panels. Falls back to the legacy
    widescreen template only if board assembly yields no slides."""
    company = (company or "SALIC").strip()
    period = (period or "").strip()
    if blocks:
        try:
            return _build_board_deck(blocks, company=company, period=period)
        except Exception as e:
            print(f"[pptx] board deck failed ({e}); falling back to legacy template")
    return _build_legacy_report_pptx(blocks, company=company, period=period)


def _build_legacy_report_pptx(blocks, company="SALIC", period=""):
    """Original 26.67×15 in template path — kept as fallback."""

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
                s = render_narrative_slide(prs, block.get("title") or "Insight",
                                           block.get("text") or "",
                                           company=company, period=period)
            elif btype == "chart":
                s = render_planned_chart_slide(prs, block.get("spec") or {},
                                               company=company, period=period)
            elif btype == "table":
                s = render_table_slide(prs, block.get("spec") or {},
                                       company=company, period=period)
            else:
                s = None
        except Exception as e:
            print(f"[pptx] report block skipped ({e})")
            s = None
        if s is not None:
            report_slides.append(s)

    total = len(report_slides)
    for i, s in enumerate(report_slides, start=1):
        _add_page_number(s, prs, i, total)

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
