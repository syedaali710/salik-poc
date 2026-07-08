"""
Builds a SALIC-branded 'Performance Summary' slide (widescreen 16:9) from the
structured fields, using the official SALIC colours and logo. Returns an in-memory
.pptx ready to stream back to the browser.
"""
import io, os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---- SALIC brand palette (from the template) ----
NAVY = RGBColor(0x0B, 0x4D, 0x89)
SKY = RGBColor(0x33, 0x9B, 0xD6)
GREEN = RGBColor(0x2A, 0x7C, 0x62)
DARKNAVY = RGBColor(0x08, 0x30, 0x5C)
RED = RGBColor(0xD9, 0x3F, 0x40)
LIGHTBLUE = RGBColor(0xEB, 0xF5, 0xFB)
CARDLINE = RGBColor(0xD6, 0xE8, 0xF5)
INK = RGBColor(0x1C, 0x2B, 0x36)
MUTED = RGBColor(0x6B, 0x7C, 0x88)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = os.path.join(HERE, "static", "salic_logo.png")

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


def _txt(slide, x, y, w, h, text, size, color, bold=False, italic=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Pt(2)
    tf.margin_top = tf.margin_bottom = Pt(1)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    f = r.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    f.name = "Calibri"
    return tb


def _rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(0.75)
    sp.shadow.inherit = False
    return sp


def _get(k, key):
    v = (k or {}).get(key)
    return v


def build_pptx(d):
    d = d or {}
    company = (d.get("company") or "Company").strip()
    period = (d.get("period") or "").strip()
    unit = (d.get("unit") or "SAR m").strip()
    kpis = d.get("kpis") or {}
    comm = [c for c in (d.get("comm") or []) if str(c).strip()][:5]

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # left accent bar
    _rect(slide, 0, 0, 0.09, 7.5, NAVY)
    # logo
    if os.path.exists(LOGO):
        slide.shapes.add_picture(LOGO, Inches(0.5), Inches(0.33), height=Inches(0.33))
    # period top-right
    _txt(slide, 9.4, 0.33, 3.4, 0.35, period, 12, MUTED, align=PP_ALIGN.RIGHT)
    # title + underline
    _txt(slide, 0.5, 0.82, 12.3, 0.55, f"{company} — Performance Summary", 26, NAVY, bold=True)
    _rect(slide, 0.52, 1.44, 0.9, 0.045, GREEN)
    _txt(slide, 0.52, 1.52, 8, 0.3, f"All figures in Million {unit.replace('m','').strip() or 'SAR'} unless stated",
         11, MUTED, italic=True)

    # KPI cards
    cx, cw, gap = 0.5, 4.0, 0.28
    for key, label in KPIDEFS:
        k = kpis.get(key) or {}
        card = _rect(slide, cx, 2.0, cw, 1.18, LIGHTBLUE, line=CARDLINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        _set_radius(card, 0.06)
        _txt(slide, cx + 0.15, 2.08, cw - 0.3, 0.3, label.upper(), 11, NAVY, bold=True)
        _txt(slide, cx + 0.13, 2.32, cw - 0.3, 0.6, fmt(_get(k, "actual")), 30, DARKNAVY, bold=True)
        sub = []
        if _get(k, "budget") is not None:
            sub.append("Budget " + fmt(_get(k, "budget")))
        if _get(k, "py") is not None:
            sub.append("PY " + fmt(_get(k, "py")))
        _txt(slide, cx + 0.15, 2.86, cw - 0.3, 0.25, "   ·   ".join(sub) or unit, 10, MUTED)
        cx += cw + gap

    # financial table
    _build_table(slide, kpis, unit)

    # commentary
    _txt(slide, 8.0, 3.5, 4.8, 0.35, (period + " Performance") if period else "Performance",
         13, GREEN, bold=True)
    if comm:
        tb = slide.shapes.add_textbox(Inches(8.0), Inches(3.92), Inches(4.85), Inches(2.9))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, c in enumerate(comm):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(7)
            r = p.add_run()
            r.text = "•  " + str(c)
            r.font.size = Pt(12)
            r.font.color.rgb = INK
            r.font.name = "Calibri"

    # footer bar
    _rect(slide, 0, 7.16, 13.333, 0.34, DARKNAVY)
    _txt(slide, 0.5, 7.17, 8, 0.32, "Saudi Agricultural and Livestock Investment Company",
         9, WHITE, anchor=MSO_ANCHOR.MIDDLE)
    _txt(slide, 9.0, 7.17, 3.8, 0.32, (period + " · Finance Update").strip(" ·"),
         9, WHITE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


def _set_radius(shape, frac):
    try:
        shape.adjustments[0] = frac
    except Exception:
        pass


def _build_table(slide, kpis, unit):
    cols = [unit, "Prior Yr", "Budget", "Actual", "Var vs B"]
    rows = 1 + len(KPIDEFS)
    gtbl = slide.shapes.add_table(rows, len(cols), Inches(0.5), Inches(3.5),
                                  Inches(7.1), Inches(0.42 * rows)).table
    gtbl.first_row = False
    gtbl.horz_banding = False
    widths = [1.7, 1.35, 1.35, 1.35, 1.35]
    for i, w in enumerate(widths):
        gtbl.columns[i].width = Inches(w)

    # header
    for c, name in enumerate(cols):
        cell = gtbl.cell(0, c)
        _cell(cell, name, WHITE, bold=True, fill=NAVY,
              align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT)

    for ri, (key, label) in enumerate(KPIDEFS, start=1):
        k = kpis.get(key) or {}
        a, b, py = _get(k, "actual"), _get(k, "budget"), _get(k, "py")
        band = RGBColor(0xF7, 0xFA, 0xFC) if ri % 2 == 0 else WHITE
        _cell(gtbl.cell(ri, 0), label, NAVY, bold=True, fill=band, align=PP_ALIGN.LEFT)
        _cell(gtbl.cell(ri, 1), fmt(py), INK, fill=band, align=PP_ALIGN.RIGHT)
        _cell(gtbl.cell(ri, 2), fmt(b), INK, fill=band, align=PP_ALIGN.RIGHT)
        _cell(gtbl.cell(ri, 3), fmt(a), INK, bold=True, fill=band, align=PP_ALIGN.RIGHT)
        if a is not None and b is not None:
            diff = a - b
            vtxt = ("▲ " if diff >= 0 else "▼ ") + fmt(abs(diff))
            vcol = GREEN if diff >= 0 else RED
        else:
            vtxt, vcol = "—", MUTED
        _cell(gtbl.cell(ri, 4), vtxt, vcol, bold=True, fill=band, align=PP_ALIGN.RIGHT)


def _cell(cell, text, color, bold=False, fill=None, align=PP_ALIGN.RIGHT):
    cell.margin_left = Inches(0.06)
    cell.margin_right = Inches(0.06)
    cell.margin_top = Inches(0.02)
    cell.margin_bottom = Inches(0.02)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    else:
        cell.fill.background()
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(11.5)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = "Calibri"
