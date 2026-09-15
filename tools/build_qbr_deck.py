#!/usr/bin/env python3
"""Builds the QBR deck template (PowerPoint, 16:9) that sits on top of the QBR cockpit workbook.

Every quantitative slide names the cockpit tab / block / table it is filled from, so the deck is a narrative layer
over the sheet rather than a second set of hand-typed numbers. Square-bracket tokens ([TEAM], [Qx-FYyy], $X.XM ...)
are the placeholders to replace; grey italic text is guidance to delete. Speaker notes carry the fill instructions.

    pip install -r tools/requirements.txt
    python3 tools/build_qbr_deck.py out/QBR_Deck_Template.pptx
"""
import re
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------- palette / type
INK = RGBColor(0x0F, 0x17, 0x2A)
TEXT = RGBColor(0x1E, 0x29, 0x3B)
MUTED = RGBColor(0x64, 0x74, 0x8B)
FAINT = RGBColor(0x94, 0xA3, 0xB8)
LINE = RGBColor(0xE2, 0xE8, 0xF0)
PANEL = RGBColor(0xF8, 0xFA, 0xFC)
PANEL2 = RGBColor(0xF1, 0xF5, 0xF9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x10, 0xB9, 0x81)
GREEN_DARK = RGBColor(0x04, 0x78, 0x57)
GREEN_BG = RGBColor(0xD1, 0xFA, 0xE5)
AMBER = RGBColor(0xD9, 0x77, 0x06)
AMBER_BG = RGBColor(0xFE, 0xF3, 0xC7)
RED = RGBColor(0xDC, 0x26, 0x26)
RED_BG = RGBColor(0xFE, 0xE2, 0xE2)
BLUE = RGBColor(0x25, 0x63, 0xEB)
BLUE_BG = RGBColor(0xDB, 0xEA, 0xFE)

FONT = "Inter"
MONO = "Roboto Mono"

W, H = Inches(13.333), Inches(7.5)
MX = Inches(0.55)  # page margin
CONTENT_TOP = Inches(1.35)
CONTENT_W = W - 2 * MX

# Placeholder tokens used everywhere so a find & replace fills the boilerplate.
# Machine-readable placeholders: src/Deck.gs (QBR > Build deck) replaces every {{token}}, fills every table whose
# first data cell is {{rows:name}} and swaps every shape holding {{chart:title}} for the cockpit chart of that title.
TEAM = "{{team}}"
Q = "{{quarter}}"
Q1 = "{{q1}}"
Q2 = "{{q2}}"
STATUS = "{{status}}"
REFRESH = "{{refreshed}}"


def tok(name):
    return "{{" + name + "}}"


def upper(label):
    """Upper-case a label but leave {{tokens}} as they are (Deck.gs matches them case-sensitively)."""
    return "".join(p if p.startswith("{{") else p.upper() for p in re.split(r"(\{\{[^}]*\}\})", label))


def rows_for(marker, template, n=3):
    """Table body for a {{rows:marker}} table: the marker in the first cell, `template` (manual-column guidance or '')
    in the others, repeated so the template looks like a table; Deck.gs resizes it to the data."""
    return [[tok("rows:" + marker) if r == 0 else ""] + list(template) for r in range(n)]

prs = Presentation()
prs.slide_width, prs.slide_height = W, H
BLANK = prs.slide_layouts[6]
slide_no = 0


# ---------------------------------------------------------------- low-level helpers
def rgb(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color


def no_line(shape):
    shape.line.fill.background()


def text(slide, x, y, w, h, runs, size=11, color=TEXT, bold=False, font=FONT, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, italic=False, spacing=None, fit=False):
    """runs: str | list[str] (one paragraph each) | list[list[(text, {overrides})]].
    fit=True (or a line count): render a long {{token}} at a reduced size so it fits the box; Deck.gs restores `size`
    (recorded in the shape description) once the token is replaced by a value."""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    if fit and isinstance(runs, str):
        size = fit_font(box, runs, size, w, h, lines=fit if isinstance(fit, int) and not isinstance(fit, bool) else 1)
    tf.vertical_anchor = anchor
    paragraphs = [runs] if isinstance(runs, str) else runs
    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if spacing:
            p.space_after = Pt(spacing)
        pieces = [(para, {})] if isinstance(para, str) else para
        for piece, over in pieces:
            r = p.add_run()
            r.text = piece
            f = r.font
            f.name = over.get("font", font)
            f.size = Pt(over.get("size", size))
            f.bold = over.get("bold", bold)
            f.italic = over.get("italic", italic)
            f.color.rgb = over.get("color", color)
    return box


FONT_SIZE_TAG = "qbr:fontSize="  # shape description, read by Deck.gs restoreFontSize()


def fit_font(shape, s, size, w, h, lines=1):
    """Font size (pt) at which `s` fits the box on `lines` lines; tags the shape with the intended size if reduced.
    No autofit XML: Google Slides import is picky about it, so the template carries a plain smaller size."""
    tf = shape.text_frame
    usable_w = (w - tf.margin_left - tf.margin_right) / 12700
    usable_h = (h - tf.margin_top - tf.margin_bottom) / 12700
    per_line = max(1, len(s) / lines)
    scale = min(1.0, usable_w / (0.7 * per_line * size), usable_h / (1.25 * lines * size))
    if scale >= 1:
        return size
    shape._element.nvSpPr.cNvPr.set("descr", f"{FONT_SIZE_TAG}{size:g}")
    return max(6, round(size * scale * 2) / 2)


def rect(slide, x, y, w, h, fill=PANEL, line=None, shape=MSO_SHAPE.RECTANGLE, dash=False):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    if fill is None:
        s.fill.background()
    else:
        rgb(s, fill)
    if line is None:
        no_line(s)
    else:
        s.line.color.rgb = line
        s.line.width = Pt(0.75)
        if dash:
            s.line.dash_style = 4  # MSO_LINE_DASH_STYLE.DASH
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = 0.08
    s.shadow.inherit = False
    return s


def pill(slide, x, y, label, fill, color, w=None, size=8.5):
    w = w or Inches(0.12 + 0.078 * len(label))
    s = rect(slide, x, y, w, Inches(0.26), fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    s.adjustments[0] = 0.5
    tf = s.text_frame
    tf.margin_left = tf.margin_right = Inches(0.06)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    if "{{" in label:
        tf.word_wrap = True
        size = fit_font(s, label, size, w, Inches(0.26))
    r.font.name, r.font.size, r.font.bold, r.font.color.rgb = MONO, Pt(size), True, color
    return s


def set_cell_border(cell, color="E2E8F0", width=Pt(0.5), sides=("a:lnB",)):
    tcPr = cell._tc.get_or_add_tcPr()
    for side in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for old in tcPr.findall(qn(side)):
            tcPr.remove(old)
    # Schema order: lnL, lnR, lnT, lnB come before the cell fill.
    for i, side in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
        ln = tcPr.makeelement(qn(side), {})
        if side in sides:
            ln.set("w", str(int(width)))
            fill = ln.makeelement(qn("a:solidFill"), {})
            clr = fill.makeelement(qn("a:srgbClr"), {"val": color})
            fill.append(clr)
            ln.append(fill)
        else:
            ln.set("w", "0")
            ln.append(ln.makeelement(qn("a:noFill"), {}))
        tcPr.insert(i, ln)


def table(slide, x, y, w, headers, rows, col_w=None, size=9, row_h=Inches(0.3), header_fill=INK,
          zebra=True, align=None, bold_first_col=False, guidance_cols=()):
    """Flat table: dark header, hairline rows. `guidance_cols` are rendered grey italic (manual commentary columns)."""
    n_rows, n_cols = len(rows) + 1, len(headers)
    shape = slide.shapes.add_table(n_rows, n_cols, x, y, w, row_h * n_rows)
    tbl = shape.table
    tblPr = tbl._tbl.tblPr
    tblPr.set("bandRow", "0")
    tblPr.set("firstRow", "0")
    for sid in tblPr.findall(qn("a:tableStyleId")):
        tblPr.remove(sid)
    weights = col_w or [1] * n_cols
    total = sum(weights)
    for i, wgt in enumerate(weights):
        tbl.columns[i].width = Emu(int(w * wgt / total))
    for r in range(n_rows):
        tbl.rows[r].height = row_h
        for c in range(n_cols):
            cell = tbl.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            val = headers[c] if r == 0 else rows[r - 1][c]
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = (align[c] if align else PP_ALIGN.LEFT) if r else (align[c] if align else PP_ALIGN.LEFT)
            run = p.add_run()
            run.text = str(val)
            f = run.font
            f.name = FONT
            f.size = Pt(size)
            if r == 0:
                f.bold, f.color.rgb = True, WHITE
                rgb(cell, header_fill)
                set_cell_border(cell, sides=())
            else:
                guidance = c in guidance_cols
                f.color.rgb = FAINT if guidance else TEXT
                f.italic = guidance
                f.bold = bold_first_col and c == 0
                rgb(cell, PANEL if (zebra and r % 2 == 0) else WHITE)
                set_cell_border(cell)
    return shape


# ---------------------------------------------------------------- page furniture
def status_pill(slide, x, y):
    return pill(slide, x, y, STATUS, BLUE_BG, BLUE, w=Inches(4.1))


def new_slide(kicker, title, source=None, subtitle=None, status=True):
    """Standard content slide: mono kicker, title, status pill, footer with cockpit source + confidentiality."""
    global slide_no
    slide_no += 1
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, W, Inches(0.08), fill=INK)
    text(s, MX, Inches(0.28), Inches(6), Inches(0.25), upper(kicker), size=8.5, color=MUTED, font=MONO, bold=True)
    title_w = CONTENT_W - Inches(4.4) if status else Inches(9.6)
    text(s, MX, Inches(0.5), title_w, Inches(0.6), title, size=20 if len(title) > 70 else 22, color=INK, bold=True,
         anchor=MSO_ANCHOR.TOP)
    if subtitle:
        text(s, MX, Inches(1.0), Inches(9.6), Inches(0.3), subtitle, size=10.5, color=MUTED)
    if status:
        status_pill(s, W - MX - Inches(4.1), Inches(0.3))
        text(s, W - MX - Inches(4.1), Inches(0.6), Inches(4.1), Inches(0.25), f"{TEAM}  ·  {Q}  ·  {REFRESH}",
             size=8.5, color=FAINT, font=MONO, align=PP_ALIGN.RIGHT)
    # footer
    rect(s, MX, H - Inches(0.5), CONTENT_W, Emu(9525), fill=LINE)
    text(s, MX, H - Inches(0.45), Inches(9.5), Inches(0.3),
         f"Source: {source}" if source else "Commentary slide - no cockpit source, filled by the presenter",
         size=8, color=MUTED, font=MONO)
    text(s, W - MX - Inches(3), H - Inches(0.45), Inches(3), Inches(0.3),
         f"Cognition  ·  Proprietary & Confidential  ·  {slide_no}", size=8, color=FAINT, align=PP_ALIGN.RIGHT,
         font=MONO)
    return s


def notes(slide, body):
    slide.notes_slide.notes_text_frame.text = body


def callout(slide, x, y, w, h, heading, body, accent=GREEN, fill=PANEL):
    rect(slide, x, y, w, h, fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(slide, x, y + Inches(0.1), Inches(0.06), h - Inches(0.2), fill=accent)
    text(slide, x + Inches(0.18), y + Inches(0.08), w - Inches(0.3), Inches(0.28), upper(heading), size=8.5,
         color=accent if accent != GREEN else GREEN_DARK, font=MONO, bold=True)
    text(slide, x + Inches(0.18), y + Inches(0.36), w - Inches(0.3), h - Inches(0.45), body, size=9.5, color=TEXT,
         spacing=3)


def guidance(slide, x, y, w, h, body, size=9):
    return text(slide, x, y, w, h, body, size=size, color=FAINT, italic=True, spacing=2)


def chart_placeholder(slide, x, y, w, h, chart_title, instruction, chart=None):
    """One dashed frame. With `chart` (cockpit chart title prefix) it carries {{chart:...}} and Build deck replaces the
    whole frame with the linked Sheets chart at this position; without it stays as guidance for a manual chart."""
    frame = rect(slide, x, y, w, h, fill=WHITE, line=FAINT, dash=True, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    tf = frame.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.2)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    paras = [[(chart_title, {"bold": True, "color": INK, "size": 10.5})],
             [(tok("chart:" + chart) if chart else "Linked chart placeholder", {"bold": True, "color": MUTED, "size": 10})],
             [(instruction, {"color": FAINT, "size": 9, "italic": True})],
             [("QBR > Build deck links this chart from the cockpit; or Insert > Chart > From Sheets and keep 'Link to spreadsheet'.",
               {"color": FAINT, "size": 8.5, "italic": True})]]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        for piece, over in para:
            r = p.add_run()
            r.text = piece
            r.font.name, r.font.size = FONT, Pt(over["size"])
            r.font.bold, r.font.italic, r.font.color.rgb = over.get("bold", False), over.get("italic", False), over["color"]
    return frame


def kpi_card(slide, x, y, w, h, label, value, sub, tone=None):
    rect(slide, x, y, w, h, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(slide, x + Inches(0.15), y + Inches(0.1), w - Inches(0.3) - (Inches(1.4) if tone else 0), Inches(0.3),
         upper(label), size=8, color=MUTED, font=MONO, bold=True)
    text(slide, x + Inches(0.15), y + Inches(0.38), w - Inches(0.3), Inches(0.42), value, size=24, color=INK,
         bold=True, fit=True)
    text(slide, x + Inches(0.15), y + h - Inches(0.42), w - Inches(0.3), Inches(0.35), sub, size=8.5, color=MUTED,
         fit=2)
    if tone:
        # tone is a {{rag.*}} token: Build deck writes ON / ABOVE PLAN, WATCH or BEHIND and colours the pill.
        pill(slide, x + w - Inches(1.45), y + Inches(0.1), tone, PANEL, MUTED, w=Inches(1.3), size=7)


def section_label(slide, x, y, w, label):
    text(slide, x, y, w, Inches(0.25), upper(label), size=8.5, color=MUTED, font=MONO, bold=True)
    rect(slide, x, y + Inches(0.27), w, Emu(9525), fill=LINE)


def bullets(slide, x, y, w, h, items, size=10, color=TEXT):
    paras = []
    for it in items:
        if isinstance(it, tuple):
            head, body = it
            paras.append([("\u2022  ", {"color": GREEN, "bold": True}), (head + " ", {"bold": True}),
                          (body, {"color": MUTED})])
        else:
            paras.append([("\u2022  ", {"color": GREEN, "bold": True}), (it, {})])
    return text(slide, x, y, w, h, paras, size=size, color=color, spacing=5)


def asks_table(slide, y, rows=4, h_col="Ask / decision needed"):
    table(slide, MX, y, CONTENT_W, [h_col, "Why it matters (deal / ARR at stake)", "Owner", "Decision by", "Status"],
          [["[ask]", "[$X.XM / account]", "[name]", "[date]", "Open"] for _ in range(rows)],
          col_w=[3.2, 3.4, 1.2, 1.1, 1.0], guidance_cols=(0, 1, 2, 3, 4))


# ================================================================ SLIDES
# 1 - Cover ------------------------------------------------------------------------------------------
slide_no += 1
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, fill=INK)
rect(s, MX, Inches(2.9), Inches(0.9), Inches(0.06), fill=GREEN)
text(s, MX, Inches(1.4), Inches(8), Inches(0.35), "QUARTERLY BUSINESS REVIEW", size=11, color=GREEN, font=MONO,
     bold=True)
text(s, MX, Inches(1.8), Inches(11), Inches(1.0), f"{TEAM}  ·  {Q}", size=40, color=WHITE, bold=True)
text(s, MX, Inches(3.15), Inches(11), Inches(0.5), "[One-line headline: the quarter in a sentence, e.g. "
     "\"Added $X.XM Net ARR at NN% of goal, N logos, zero churn; Q+1 coverage is the risk\"]", size=14,
     color=FAINT, italic=True)
pill(s, MX, Inches(4.1), STATUS, BLUE, WHITE, w=Inches(4.1))
text(s, MX, Inches(4.5), Inches(8), Inches(0.3), f"Numbers from the QBR cockpit  ·  {REFRESH}  ·  Salesforce as of that date",
     size=9.5, color=FAINT, font=MONO)
text(s, MX, Inches(6.3), Inches(8), Inches(0.6),
     [[("Presented by ", {"color": FAINT}), ("[Sales leader]  ·  [Engineering leader]", {"color": WHITE})],
      [("Date ", {"color": FAINT}), ("[QBR date]", {"color": WHITE})]], size=11)
text(s, W - MX - Inches(4), H - Inches(0.6), Inches(4), Inches(0.3), "Cognition  ·  Proprietary & Confidential",
     size=8.5, color=FAINT, font=MONO, align=PP_ALIGN.RIGHT)
notes(s, "QBR > Build deck (this tab) fills {{team}}, {{quarter}}, the status pill and the refresh date from the team "
         "tab (B1 / B2 and the row-1 banner). The headline is the one sentence a reader should remember; write it last.")

# 2 - How to use --------------------------------------------------------------------------------------
s = new_slide("How to use this template", "One cockpit, one story per slide, explicit asks",
              source="Template guidance - delete this slide before presenting", status=False)
col_w = (CONTENT_W - Inches(0.3)) / 2
bullets(s, MX, CONTENT_TOP, col_w, Inches(4.6), [
    ("Numbers come from the QBR cockpit, not from memory.", "Set B1 = team, B2 = quarter, QBR > Build deck (this tab): "
     "it refreshes the tab and fills every number, table and chart. Each slide's footer names the source block; the appendix maps each metric."),
    ("Actuals vs forecast is always visible.", "The blue pill on every slide mirrors the tab banner: while the quarter is "
     "open all quarter numbers are FORECAST (quarter in progress); after quarter end they are ACTUALS."),
    ("QoQ is like for like.", "While the quarter is open the cockpit compares to the previous quarter at the same elapsed "
     "day; open pipeline / coverage / conversion have no same-point value and show '-'."),
    ("Charts are linked, not pasted.", "Build deck inserts the cockpit charts linked to the sheet; Tools > Linked objects > "
     "Update all refreshes them after a cockpit refresh."),
], size=10)
bullets(s, MX + col_w + Inches(0.3), CONTENT_TOP, col_w, Inches(4.6), [
    ("Commentary is structured.", "Sales and Engineering leadership each have a look-back and a look-ahead slide with "
     "fixed prompts, plus one 'Asks & decisions' table (owner, date) - no free-floating prose."),
    ("Every insight needs a 'so what'.", "Each slide has one headline (the takeaway) and, where relevant, a 'So what / "
     "action' callout. If a slide has no takeaway, drop it."),
    ("Keep it to ~25 slides.", "Details (full tables, rep list, definitions, data-quality issues) live in the appendix "
     "and in the cockpit itself; the cockpit PDF (cell M2) is the pre-read attachment."),
    ("Placeholders.", "Double-brace tokens and tables are filled by QBR > Build deck (this tab) from the selected team tab; "
     "[square brackets] and grey italic text are commentary to write or delete."),
], size=10)
notes(s, "Delete before presenting.")

# 3 - Agenda -----------------------------------------------------------------------------------------
s = new_slide("Agenda", f"{Q} review and {Q1} / {Q2} outlook", source="-", status=False)
agenda = [
    ("01", "Executive summary", "Headline results, key risks, decisions needed", "4"),
    ("02", f"{Q} look back", "Scorecard, ARR bridge, deals won, logos, retention, pilots", "5-10"),
    ("03", "Leadership look back", "Sales and Engineering commentary on what happened and why", "11-12"),
    ("04", f"{Q1} / {Q2} look ahead", "Forecast, pipeline quality, deals to win, renewals at risk, path to goal", "13-17"),
    ("05", "Leadership look ahead", "Sales and Engineering plan, capacity, asks", "18-19"),
    ("06", "Team", "Rep activity and performance, hiring and capacity", "20-21"),
    ("07", "Asks & decisions", "Single table of decisions needed from leadership", "22"),
    ("A", "Appendix", "Cockpit mapping, definitions, data quality", "23-25"),
]
y = CONTENT_TOP
for num, head, desc, pages in agenda:
    rect(s, MX, y, CONTENT_W, Inches(0.55), fill=PANEL if agenda.index((num, head, desc, pages)) % 2 == 0 else WHITE)
    text(s, MX + Inches(0.15), y + Inches(0.12), Inches(0.6), Inches(0.3), num, size=12, color=GREEN, font=MONO, bold=True)
    text(s, MX + Inches(0.8), y + Inches(0.08), Inches(4.5), Inches(0.4), head, size=13, color=INK, bold=True)
    text(s, MX + Inches(5.2), y + Inches(0.12), Inches(6), Inches(0.35), desc, size=10, color=MUTED)
    text(s, W - MX - Inches(1.2), y + Inches(0.12), Inches(1.05), Inches(0.35), pages, size=9, color=FAINT, font=MONO,
         align=PP_ALIGN.RIGHT)
    y += Inches(0.6)
notes(s, "Fixed structure - keep the section order so decks are comparable quarter to quarter. Update the slide ranges "
         "only if you add or remove slides.")

# 4 - Executive summary ------------------------------------------------------------------------------
s = new_slide("01 · Executive summary", "[Headline: the quarter in one sentence - result, cause, implication]",
              source=f"{TEAM} tab > KPI cards (row 5) and {Q} QUARTER block; Definitions tab for every metric")
cw = (CONTENT_W - 3 * Inches(0.2)) / 4
cards = [("Net Added ARR", tok("netAddedArr"),
          f"Goal {tok('revenueGoal')}  ·  {tok('attainment')} attainment  ·  pace {tok('pace')}", tok("rag.attainment")),
         ("Logos won (Majors)", tok("logosWonMajors"),
          f"Goal {tok('logoGoal')}  ·  Land only {tok('logosWonLand')}  ·  Land + MSP {tok('logosWonMajors')}  ·  "
          f"expected incl. open {tok('logoAttainment')}", tok("rag.logoAttainmentPct")),
         ("Churn / renewals", tok("churnArr"),
          f"{tok('churnCustomers')} customers  ·  {tok('wonRenewals')} of {tok('renewals')} renewals won  ·  "
          f"GRR {tok('grr')}  ·  NRR {tok('nrr')}", tok("rag.renewalRate")),
         ("Ending ARR", tok("endingArr"), f"Starting {tok('startingArr')}  ·  {tok('endingArrQoq')} vs Starting ARR", None)]
for i, (lab, val, sub, tone) in enumerate(cards):
    kpi_card(s, MX + i * (cw + Inches(0.2)), CONTENT_TOP, cw, Inches(1.45), lab, val, sub, tone)
y = CONTENT_TOP + Inches(1.7)
half = (CONTENT_W - Inches(0.3)) / 2
callout(s, MX, y, half, Inches(3.55), "What happened - 3 to 5 takeaways",
        ["1. [Result vs goal and what drove it - name the deals/accounts]",
         "2. [Logo performance - landed vs expected, Land vs MSP]",
         "3. [Retention - churn/downgrades and why; renewals won/lost]",
         "4. [Pipeline / pilots - what changed in coverage or stage mix]",
         "5. [Team - hiring, coverage, notable rep performance]"])
callout(s, MX + half + Inches(0.3), y, half, Inches(1.65), "Top risks to " + Q1,
        ["1. [Risk - deal / renewal / capacity, $ at stake, owner]", "2. [Risk]", "3. [Risk]"], accent=AMBER)
callout(s, MX + half + Inches(0.3), y + Inches(1.9), half, Inches(1.65), "Decisions needed today",
        ["1. [Decision - what, from whom, by when]  (detail on slide 22)", "2. [Decision]"], accent=RED)
notes(s, "KPI cards are filled by Build deck from the tab's KPI cards / metric block; the pills carry the cockpit's RAG "
         "(green >= 100% / on plan, amber 70-99% / watch, red < 70% / behind). Write the takeaways as conclusions "
         "('Barclays land carried the quarter; without it attainment is 30%'), not as data.")

# 5 - Scorecard --------------------------------------------------------------------------------------
s = new_slide(f"02 · {Q} look back", "Scorecard: goal, result, attainment and like-for-like QoQ",
              source=f"{TEAM} tab > 'ACTUALS | FORECAST QUARTER' block (Metric / {Q} / QoQ / Trend columns) "
                     f"and Accounts block; Goals tab for goals")
rows = [
    ["Net Added ARR", tok("revenueGoal"), tok("netAddedArr"), tok("attainment"), tok("qoq.netAddedArr"),
     "Closed Won Delta ARR + full-churn renewals (MSP ARR when NACV = 0)"],
    ["Quarter elapsed / pace", "-", tok("quarterElapsedPct"), tok("pace"), tok("qoq.pace"), "Attainment / elapsed: 100% = straight-line to goal"],
    ["Logo goal / Logos Won (Land only, Majors)", tok("logoGoal"), tok("logosWonLand"), "-", tok("qoq.logosWonLand"), "Closed Won Land on Major accounts"],
    ["Logos Won (Land + MSP, Majors)", tok("logoGoal"), tok("logosWonMajors"), "-", tok("qoq.logosWonMajors"), "MSP on a new account counts as a landed logo"],
    ["Logo attainment (expected, incl. open opps)", tok("logoGoal"), tok("logoAttainment"), tok("logoAttainmentPct"), tok("qoq.logoAttainment"),
     "Sum of Expected Logo Impact - forecast until closed"],
    ["# Renewals / # won / renewal rate", "-", f"{tok('renewals')} / {tok('wonRenewals')} / {tok('renewalRate')}", "-", tok("qoq.renewalRate"),
     "Closed renewals in the quarter, by count"],
    ["Churn ARR / # customers", "-", f"{tok('churnArr')} / {tok('churnCustomers')}", "-", tok("qoq.churnArr"), "Full churn + downgrades"],
    ["GRR / NRR", "-", f"{tok('grr')} / {tok('nrr')}", "-", tok("qoq.grr"), "Of Starting ARR"],
    ["Starting -> Ending ARR", "-", f"{tok('startingArr')} -> {tok('endingArr')}", "-", tok("qoq.endingArr"),
     "ARR Ledger balances (full quarter, not same-point)"],
    ["Open pipeline (this quarter) / coverage", "-", f"{tok('openPipelineArr')} / {tok('pipelineCoverage')}", "-", tok("qoq.openPipelineArr"),
     "No same-point value while the quarter is open"],
]
table(s, MX, CONTENT_TOP, CONTENT_W, ["Metric", "Goal", Q, "Attainment", tok("qoqHeader"), "Definition (cockpit)"],
      rows, col_w=[3.0, 0.9, 1.6, 1.0, 1.7, 4.0], size=8.5, row_h=Inches(0.36), bold_first_col=True, guidance_cols=(5,))
pill(s, MX, CONTENT_TOP + Inches(4.1), tok("rag.attainment"), PANEL, MUTED, w=Inches(1.5), size=7)
guidance(s, MX + Inches(1.7), CONTENT_TOP + Inches(4.08), CONTENT_W - Inches(1.7), Inches(0.5),
         "Revenue attainment RAG (green >= 100%, amber 70-99%, red < 70%). The QoQ header states the cockpit's basis: "
         "same elapsed day of the previous quarter while open, full quarter once closed; '-' = no like-for-like value.")
notes(s, "Filled from the metric block by Build deck. Definitions are the Definitions tab wording - do not retype them. "
         "A '-' means the cockpit has no same-point baseline; do not fill it from another report.")

# 6 - ARR bridge -------------------------------------------------------------------------------------
s = new_slide(f"02 · {Q} look back", "[Headline: e.g. 'Ending ARR $X.XM, +NN% QoQ, driven by N expansions; churn limited to N account']",
              source=f"{TEAM} tab > ARR bridge block (Starting, Added, Downgrade, Full Churn, Ending, GRR, NRR) "
                     f"and CHARTS > 'ARR bridge {Q}'; ARR Ledger tab for Starting / Ending")
chart_placeholder(s, MX, CONTENT_TOP, Inches(7.4), Inches(4.3), f"ARR bridge {Q}",
                  "Waterfall from the team tab CHARTS section (Starting, Added, Downgrade, Full churn, Ending)", chart="ARR bridge")
x2 = MX + Inches(7.7)
table(s, x2, CONTENT_TOP, CONTENT_W - Inches(7.7), ["Bridge", "$", "% of Starting"],
      [["Starting ARR (ledger)", tok("startingArr"), "100%"], ["+ Added ARR (won)", tok("addedArr"), tok("bridge.addedPct")],
       [f"- Downgrades ({tok('downgradeCount')})", tok("downgradeArr"), tok("bridge.downgradePct")],
       [f"- Full churn ({tok('fullChurnCount')})", tok("fullChurnArr"), tok("bridge.fullChurnPct")],
       ["= Ending ARR", tok("endingArr"), tok("bridge.endingPct")], ["GRR", tok("grr"), ""], ["NRR", tok("nrr"), ""]],
      col_w=[2.2, 1.1, 1.1], size=9, row_h=Inches(0.32), bold_first_col=True)
callout(s, x2, CONTENT_TOP + Inches(2.75), CONTENT_W - Inches(7.7), Inches(1.55), "So what",
        "[What the bridge says about the book: concentration of Added ARR in 1-2 deals, whether churn was expected, "
        "and what Ending ARR implies for next quarter's starting point.]")
notes(s, "Starting / Ending ARR are ARR Ledger balances (seeded from the cockpit for Q1/Q2-2026, rolled forward from "
         "Salesforce Net Added ARR afterwards). Added ARR includes the MSP fallback (ARR when NACV = 0).")

# 7 - Deals won --------------------------------------------------------------------------------------
s = new_slide(f"02 · {Q} look back", "[Headline: e.g. 'N deals won for $X.XM Delta ARR - Net Added ARR ties out to these rows']",
              source=f"{TEAM} tab > 'Top 10 Deals Won {Q}' table (Account, Opportunity link, Type, Deal value (TCV), Delta ARR, Owner)")
table(s, MX, CONTENT_TOP, Inches(8.6), ["Account", "Type", "Deal value (TCV)", "Delta ARR", "Owner", "Why we won / what it unlocks"],
      rows_for("dealsWon", ["", "", "", "", "[1 line - why we won, what it unlocks]"], n=10),
      col_w=[2.0, 1.6, 1.2, 1.1, 1.3, 3.0], size=8.5, row_h=Inches(0.34), bold_first_col=True, guidance_cols=(5,))
x2 = MX + Inches(8.9)
callout(s, x2, CONTENT_TOP, CONTENT_W - Inches(8.9), Inches(1.5), "Tie-out",
        f"{tok('wonCount')} wins, {tok('wonArr')} Delta ARR ({tok('dealsWonShown')} shown; other wins {tok('dealsWonOther')} / "
        f"{tok('dealsWonOtherArr')}) + full-churn renewals {tok('fullChurnArr')} = Net Added ARR {tok('netAddedArr')}. "
        "Attribution is by account team: a deal owned by another team's rep on this team's account counts here.")
callout(s, x2, CONTENT_TOP + Inches(1.7), CONTENT_W - Inches(8.9), Inches(1.3), "TCV vs Delta ARR",
        "MSP deals: TCV is the contract value; Delta ARR is the ARR booked (NACV, or the opp's ARR when NACV is 0). "
        "Explain any large gap in one line.", accent=BLUE)
callout(s, x2, CONTENT_TOP + Inches(3.2), CONTENT_W - Inches(8.9), Inches(1.1), "Concentration",
        f"Top deal = {tok('topDealShare')} of won Delta ARR. [Say whether that is a risk.]", accent=AMBER)
notes(s, "Filled from the Top 10 Deals Won table (account names link to Salesforce). With more than ten wins Build deck "
         "adds an 'Other wins not shown' row so the tie-out holds. 'Why we won' is the only manual column.")

# 8 - Logos ------------------------------------------------------------------------------------------
s = new_slide(f"02 · {Q} look back", "[Headline: e.g. 'N of N logos landed; N more expected from open Major opps']",
              source=f"{TEAM} tab > metric block rows 'Logo Goal', 'Logos Won (Land only, Majors)', 'Logos Won (Land + MSP, Majors)', "
                     f"'Logo Attainment (expected, incl. open opps)' and 'Logos Won (Land + MSP)' table")
cw = (Inches(6.2) - 2 * Inches(0.2)) / 3
for i, (lab, val, sub) in enumerate([("Logo goal", tok("logoGoal"), "Goals tab"), ("Won - Land only", tok("logosWonLand"), "Closed Won Land, Majors"),
                                     ("Won - Land + MSP", tok("logosWonMajors"), "MSP on new account = logo")]):
    kpi_card(s, MX + i * (cw + Inches(0.2)), CONTENT_TOP, cw, Inches(1.3), lab, val, sub)
pill(s, MX, CONTENT_TOP + Inches(1.45), f"Expected logo attainment {tok('logoAttainment')} incl. open opps ({tok('logoAttainmentPct')} of goal) - "
     "forecast until closed", BLUE_BG, BLUE, w=Inches(6.2), size=8)
table(s, MX, CONTENT_TOP + Inches(1.9), Inches(6.2), ["Account (logo won)", "Type", "TCV", "Delta ARR", "Owner"],
      rows_for("logosWon", ["", "", "", ""], n=5), col_w=[2.4, 1.0, 1.0, 1.0, 1.4],
      size=8.5, row_h=Inches(0.32), bold_first_col=True)
x2 = MX + Inches(6.5)
table(s, x2, CONTENT_TOP, CONTENT_W - Inches(6.5), ["Open Major logo opp (this quarter)", "Stage", "Expected logo impact", "Close date"],
      rows_for("openLogoOpps", ["", "", ""], n=5), col_w=[2.6, 1.6, 1.4, 1.1], size=8.5,
      row_h=Inches(0.32), bold_first_col=True)
callout(s, x2, CONTENT_TOP + Inches(2.2), CONTENT_W - Inches(6.5), Inches(1.6), "So what",
        "[Landed vs expected gap: which open opps must close to hit the logo goal, and whether Expected Logo Impact in "
        "Salesforce is set correctly (an MSP logo with impact 0 under-states the forecast).]")
notes(s, "Won logos: 'Logos Won (Land + MSP)' table. Open Major logo opps: open Land / MSP opps on Major accounts with a "
         "close date in the quarter, sorted by Expected Logo Impact (same rows as Raw Data).")

# 9 - Retention --------------------------------------------------------------------------------------
s = new_slide(f"02 · {Q} look back", "[Headline: e.g. 'N of N renewals won; $X.XM churned across N accounts, all [reason]']",
              source=f"{TEAM} tab > 'Churned Customers', 'Downgrade Customers', 'Renewals Won', 'Renewals Lost' tables; "
                     f"metric rows '# Renewals', '# Won Renewals', 'Renewal Rate', 'Churn - ARR $', 'Reasons for Churn'; CHARTS > 'Renewals {Q}'")
cw = (CONTENT_W - 3 * Inches(0.2)) / 4
for i, (lab, val, sub) in enumerate([("Renewals closed", tok("renewals"), f"Won {tok('wonRenewals')} · Lost {tok('lostRenewals')}"),
                                     ("Renewal rate", tok("renewalRate"), "By count; '-' if none closed"),
                                     ("Churn ARR", tok("churnArr"), f"Full churn {tok('fullChurnCount')} · Downgrade {tok('downgradeCount')}"),
                                     ("GRR / NRR", f"{tok('grr')} / {tok('nrr')}", "Of Starting ARR")]):
    kpi_card(s, MX + i * (cw + Inches(0.2)), CONTENT_TOP, cw, Inches(1.2), lab, val, sub)
table(s, MX, CONTENT_TOP + Inches(1.45), Inches(8.2), ["Account", "Outcome", "Record type", "Close date", "Delta ARR", "Reason / lesson"],
      rows_for("retention", ["", "", "", "", "[Salesforce lost reason if any; add what we learned]"], n=6),
      col_w=[2.0, 1.9, 1.2, 1.0, 1.0, 3.0], size=8.5, row_h=Inches(0.34), bold_first_col=True, guidance_cols=(5,))
callout(s, MX + Inches(8.5), CONTENT_TOP + Inches(1.45), CONTENT_W - Inches(8.5), Inches(1.6), "Was it predicted?",
        "[Compare to last QBR's 'Predicted churn' slide: which churns were on it, which were not, and why.]", accent=AMBER)
callout(s, MX + Inches(8.5), CONTENT_TOP + Inches(3.25), CONTENT_W - Inches(8.5), Inches(1.15), "Engineering read-across",
        "[Any churn / downgrade with a product or deployment root cause - hand to slide 12.]", accent=BLUE)
notes(s, "A 0% renewal rate with N closed means all N renewals closed in the quarter were lost; '-' means none closed "
         "yet. Say which in the headline so it is not misread.")

# 10 - Pilots ----------------------------------------------------------------------------------------
s = new_slide(f"02 · {Q} look back", "[Headline: e.g. 'N pilots completed (N converted), N active; DE capacity is the constraint for Q+1']",
              source=f"{TEAM} tab > '# Active Pilots (open, 3- Tech Validation)', '# Pilots Completed (this quarter)', "
                     f"'Active Pilots' and 'Pilots Completed {Q}' tables; Accounts block '# Activated Prospects', 'Conversion Rate'")
cw = (Inches(5.0) - Inches(0.2)) / 2
kpi_card(s, MX, CONTENT_TOP, cw, Inches(1.2), "Pilots completed", tok("pilotsCompleted"), "Pilot Status = Complete, ended this quarter")
kpi_card(s, MX + cw + Inches(0.2), CONTENT_TOP, cw, Inches(1.2), "Active pilots", tok("activePilots"), "Open opps in 3- Tech Validation")
kpi_card(s, MX, CONTENT_TOP + Inches(1.4), cw, Inches(1.2), "Activated prospects", tok("activatedProspects"), "No ARR, open opp")
kpi_card(s, MX + cw + Inches(0.2), CONTENT_TOP + Inches(1.4), cw, Inches(1.2), "Conversion rate", tok("conversionRate"),
         "Won lands / (won lands + activated) - full quarter only")
table(s, MX + Inches(5.3), CONTENT_TOP, CONTENT_W - Inches(5.3),
      ["Pilot (account)", "Status / stage now", "End / expected end", "Delta ARR", "Outcome + engineering notes"],
      rows_for("pilots", ["", "", "", "[Converted / lost / extended; DE effort, blockers]"], n=7),
      col_w=[2.0, 1.7, 1.2, 0.9, 2.9], size=8.5, row_h=Inches(0.34), bold_first_col=True, guidance_cols=(4,))
callout(s, MX, CONTENT_TOP + Inches(2.8), Inches(5.0), Inches(1.5), "So what",
        "[Pilot -> win conversion this quarter; pilots that ran long and why; DE hours per pilot if known. Feeds the "
        "Engineering look-ahead (slide 19).]", accent=BLUE)
notes(s, "Active = open opps in stage '3- Tech Validation' as of the refresh (whatever the close date). Completed = "
         "Pilot Status 'Complete' with Pilot Actual End Date in the quarter. Conversion rate is blank in the QoQ "
         "column by design (no historical account state).")

# 11 - Sales leadership look back --------------------------------------------------------------------
s = new_slide("03 · Leadership look back", "Sales leadership commentary - what happened and why", status=False)
text(s, W - MX - Inches(4.1), Inches(0.3), Inches(4.1), Inches(0.3), "[Sales leader name]  ·  " + Q, size=9, color=MUTED,
     font=MONO, align=PP_ALIGN.RIGHT)
half = (CONTENT_W - Inches(0.3)) / 2
qh = Inches(2.5)
callout(s, MX, CONTENT_TOP, half, qh, "1 · What worked - and is repeatable",
        "[Deals / motions that drove the result. Name accounts, owners, partner involvement. Which of it can the team "
        "repeat next quarter?]")
callout(s, MX + half + Inches(0.3), CONTENT_TOP, half, qh, "2 · What did not work - root cause",
        "[Slipped or lost deals (see 'Lost Pipeline' table), stalled pilots, churn. Root cause, not symptom: "
        "qualification, pricing, competition, execution, product.]", accent=AMBER)
callout(s, MX, CONTENT_TOP + qh + Inches(0.2), half, qh, "3 · Forecast accuracy vs last QBR",
        "[Last QBR's commit for this quarter: $X.XM / N logos. Landed: $X.XM / N. What moved in or out and what that "
        "says about how we forecast.]", accent=BLUE)
callout(s, MX + half + Inches(0.3), CONTENT_TOP + qh + Inches(0.2), half, qh, "4 · Account movements that matter",
        "[Top customers up / down (Top 10 Major / Enterprise Customers tables, '% of active ARR'), concentration risk, "
        "new executive relationships, competitive displacements.]", accent=GREEN)
notes(s, "Fill all four boxes; 3-5 lines each. Reference cockpit tables rather than re-typing numbers. Prompt 3 needs "
         "the previous deck's slide 16 (Q+1 commit) - keep decks so this comparison is always possible.")

# 12 - Engineering leadership look back ---------------------------------------------------------------
s = new_slide("03 · Leadership look back", "Engineering leadership commentary - delivery and product", status=False)
text(s, W - MX - Inches(4.1), Inches(0.3), Inches(4.1), Inches(0.3), "[Engineering leader name]  ·  " + Q, size=9,
     color=MUTED, font=MONO, align=PP_ALIGN.RIGHT)
callout(s, MX, CONTENT_TOP, half, qh, "1 · Pilots and deployments delivered",
        "[Pilots completed (slide 10) - outcome, DE effort, time to value. Production deployments live / expanded. "
        "What made the successful ones succeed.]", accent=BLUE)
callout(s, MX + half + Inches(0.3), CONTENT_TOP, half, qh, "2 · Technical blockers hit this quarter",
        "[Security / compliance reviews, integrations, environment access, model or product gaps that slowed or cost "
        "deals. Which were resolved, which remain.]", accent=AMBER)
callout(s, MX, CONTENT_TOP + qh + Inches(0.2), half, qh, "3 · Customer health from the engineering side",
        "[Adoption / consumption signals on the top accounts, support escalations, accounts where usage does not match "
        "the ARR (expansion or churn risk). Link to churn on slide 9.]", accent=GREEN)
callout(s, MX + half + Inches(0.3), CONTENT_TOP + qh + Inches(0.2), half, qh, "4 · Product feedback to route",
        "[Top 3 feature / roadmap requests from this team's customers, each with the deal or ARR attached and the "
        "product owner it was routed to.]", accent=RED)
notes(s, "Written by the Engineering / Deployed Engineering leader for the region. Keep each box to what changed this "
         "quarter; capacity and asks go on slide 19.")

# 13 - Forecast summary -------------------------------------------------------------------------------
s = new_slide(f"04 · {Q1} / {Q2} look ahead", f"[Headline: e.g. '{Q1} forecast $X.XM vs $X.XM goal (NN%); coverage N.Nx is below the 3x bar']",
              source=f"{TEAM} tab > FUTURE QUARTER(S) block ({Q1} and {Q2} columns: Revenue Goal, Net Forecast, Logo Goal, Pipeline, "
                     f"# Renewals due, ARR up for renewal, Starting -> Forecast Ending ARR, Forecast Churn); CHARTS > 'Q+1 / Q+2 forecast vs goal'")
pill(s, MX, CONTENT_TOP - Inches(0.05), "FORECAST (quarter not started) - Net forecast = Expected Delta ARR; Pipeline = open opps' Delta ARR",
     BLUE_BG, BLUE, w=Inches(6.6), size=8)


def f12(names, sep=" / "):
    """The Q+1 and Q+2 cells of a forecast row: the given metrics joined by `sep`, as f1./f2. tokens."""
    return [sep.join(tok(f"f{i}.{n}") for n in names) for i in (1, 2)]


rows = [["Revenue goal", tok("f1.revenueGoal"), tok("f2.revenueGoal")],
        ["Net forecast ($) / (%)"] + f12(["netForecastArr", "netForecastPct"]),
        ["Logo goal / Net forecast (#)"] + f12(["logoGoal", "logoForecast"]),
        ["Open pipeline / coverage (x goal)"] + f12(["pipelineArr", "pipelineCoverage"]),
        ["# Renewals due / ARR up for renewal"] + f12(["renewalsDueCount", "renewalArrDue"]),
        ["Forecast churn ARR / #"] + f12(["forecastChurnArr", "forecastChurnCount"]),
        ["Starting ARR -> Forecast Ending ARR"] + f12(["startingArr", "forecastEndingArr"], sep=" -> ")]
table(s, MX, CONTENT_TOP + Inches(0.35), Inches(6.6), ["Metric", Q1, Q2], rows, col_w=[3.0, 1.8, 1.8], size=9,
      row_h=Inches(0.36), bold_first_col=True)
chart_placeholder(s, MX + Inches(6.9), CONTENT_TOP, CONTENT_W - Inches(6.9), Inches(2.6), "Q+1 / Q+2 forecast vs goal",
                  "Goal, Net Forecast and Pipeline per future quarter (team tab CHARTS section)", chart="Q+1 / Q+2 forecast vs goal")
callout(s, MX + Inches(6.9), CONTENT_TOP + Inches(2.8), CONTENT_W - Inches(6.9), Inches(1.5), "So what",
        f"{Q1}: {tok('f1.gapToForecast')} between goal and net forecast, coverage {tok('f1.pipelineCoverage')} (rule of thumb 3x), "
        f"logo gap {tok('f1.logoGap')}. [What has to be true for the forecast to land.]", accent=AMBER)
notes(s, "Net Forecast = Expected Delta ARR of every Land / MSP / Expand opp and growing renewal with a close date in the "
         "quarter (closed ones included, at their expected value) plus forecast churn (renewals with Expected Delta ARR < 0, "
         "the 'Predicted Churn' tables). Open pipeline / coverage = Delta ARR of open opps only. Both are labelled FORECAST "
         "until the quarter starts and then follow the tab banner.")

# 14 - Pipeline quality -------------------------------------------------------------------------------
s = new_slide(f"04 · {Q1} / {Q2} look ahead", "[Headline: e.g. 'NN% of Q+1 pipeline is still in Discovery / Scope - it is early, not real']",
              source=f"{TEAM} tab > 'Pipeline by stage {Q}', 'Pipeline by stage Q+1 {Q1}', 'Pipeline by stage Q+2 {Q2}' tables "
                     f"(Stage, # Open opps, Delta ARR, % of pipeline)")
tw = (CONTENT_W - 2 * Inches(0.25)) / 3
for i, (q, marker) in enumerate([(Q, "stages"), (Q1, "stagesQ1"), (Q2, "stagesQ2")]):
    x = MX + i * (tw + Inches(0.25))
    section_label(s, x, CONTENT_TOP, tw, f"Pipeline by stage {q}")
    table(s, x, CONTENT_TOP + Inches(0.35), tw, ["Stage", "# Opps", "Delta ARR", "% of pipe"], rows_for(marker, ["", "", ""], n=7),
          col_w=[1.8, 0.8, 1.1, 0.9], size=8.5, row_h=Inches(0.3), bold_first_col=True)
callout(s, MX, CONTENT_TOP + Inches(3.05), Inches(6.2), Inches(1.3), "Early-stage share",
        f"{Q1}: {tok('f1.earlySharePct')} of pipeline in stages 0-2  ·  {Q2}: {tok('f2.earlySharePct')}.  [Rule: pipeline in stages 0-2 "
        "inside the quarter it is meant to close in is at risk; say how much of the forecast depends on it.]", accent=AMBER)
callout(s, MX + Inches(6.5), CONTENT_TOP + Inches(3.05), CONTENT_W - Inches(6.5), Inches(1.3), "Pipeline hygiene",
        "[Stalled >60d ARR from the Reps table; data-quality issues (slide 25) that distort this view: missing close "
        "dates, opps without Expected Delta ARR.]", accent=RED)
notes(s, "Stage names are the Salesforce picklist, sorted by their numeric prefix; '% of pipeline' is blank when total "
         "open Delta ARR is 0 or negative. Build deck lists only the stages that have opps, plus a Total open row.")

# 15 - Deals to win -----------------------------------------------------------------------------------
s = new_slide(f"04 · {Q1} / {Q2} look ahead", "[Headline: e.g. 'Top N deals = NN% of the Q+1 forecast; two are pilot-gated']",
              source=f"{TEAM} tab > 'Top 10 Deals Q+1 {Q1}' and 'Top 10 Deals Q+2 {Q2}' tables (Account, Opportunity link, Stage, Close date, Delta ARR, Owner)")
half = (CONTENT_W - Inches(0.3)) / 2
for i, (q, marker) in enumerate([(Q1, "topDealsQ1"), (Q2, "topDealsQ2")]):
    x = MX + i * (half + Inches(0.3))
    section_label(s, x, CONTENT_TOP, half, f"Top deals {q}  ·  {tok(f'f{i + 1}.topDealsArr')} = {tok(f'f{i + 1}.topDealsSharePct')} of open pipeline")
    table(s, x, CONTENT_TOP + Inches(0.35), half, ["Account", "Stage", "Close", "Delta ARR", "Next step / blocker / help needed"],
          rows_for(marker, ["", "", "", "[1 line - owner's commitment]"], n=8),
          col_w=[1.7, 1.2, 0.8, 0.9, 2.3], size=8.5, row_h=Inches(0.33), bold_first_col=True, guidance_cols=(4,))
notes(s, "Keep the Salesforce links. The last column is the manual part and the basis for the deal review in the "
         "meeting: what happens next, by when, and who outside the team needs to help (exec sponsor, DE, legal).")

# 16 - Renewals at risk -------------------------------------------------------------------------------
s = new_slide(f"04 · {Q1} / {Q2} look ahead", "[Headline: e.g. '$X.XM of Q+1 renewals carry -$X.XM expected churn; mitigation owners assigned']",
              source=f"{TEAM} tab > 'Renewals due Q+1 {Q1}' / 'Renewals due Q+2 {Q2}' and 'Predicted Churn Q+1 {Q1}' / 'Predicted Churn Q+2 {Q2}' tables "
                     f"(Account, Opportunity link, Close date, Current ARR, Expected Delta ARR, Owner)")
cw = (Inches(4.0) - Inches(0.2)) / 2
kpi_card(s, MX, CONTENT_TOP, cw, Inches(1.2), f"Renewals due {Q1}", tok("f1.renewalsDueCount"), f"{tok('f1.renewalArrDue')} up for renewal")
kpi_card(s, MX + cw + Inches(0.2), CONTENT_TOP, cw, Inches(1.2), f"Predicted churn {Q1}", tok("f1.forecastChurnArr"),
         f"{tok('f1.forecastChurnCount')} renewals with Expected Delta ARR < 0")
kpi_card(s, MX, CONTENT_TOP + Inches(1.4), cw, Inches(1.2), f"Renewals due {Q2}", tok("f2.renewalsDueCount"), f"{tok('f2.renewalArrDue')} up for renewal")
kpi_card(s, MX + cw + Inches(0.2), CONTENT_TOP + Inches(1.4), cw, Inches(1.2), f"Predicted churn {Q2}", tok("f2.forecastChurnArr"),
         f"{tok('f2.forecastChurnCount')} renewals with Expected Delta ARR < 0")
table(s, MX + Inches(4.3), CONTENT_TOP, CONTENT_W - Inches(4.3),
      ["Account (renewal)", "Close", "Current ARR", "Expected Delta ARR", "Risk driver", "Mitigation + owner"],
      rows_for("renewalsAtRisk", ["", "", "", "[usage / budget / champion / product]", "[action, owner, date]"], n=7),
      col_w=[1.9, 0.8, 1.0, 1.2, 1.7, 2.2], size=8.5, row_h=Inches(0.34), bold_first_col=True, guidance_cols=(4, 5))
callout(s, MX, CONTENT_TOP + Inches(2.8), Inches(4.0), Inches(1.5), "Engineering read-across",
        "[Renewals where adoption / consumption is below contract: what engineering can do before the renewal date.]",
        accent=BLUE)
notes(s, "Predicted churn = renewal opps (renewal record types) with Expected Delta ARR < 0, Q+1 rows then Q+2 rows. "
         "Current ARR is the account's Current ARR. 'Risk driver' and 'Mitigation' are manual.")

# 17 - Path to goal -----------------------------------------------------------------------------------
s = new_slide(f"04 · {Q1} / {Q2} look ahead", f"[Headline: e.g. 'Path to the {Q1} goal: $X.XM won + $X.XM commit leaves a $X.XM gap to close from best case']",
              source=f"{TEAM} tab > FUTURE QUARTER(S) block (Revenue Goal, Net Forecast, Pipeline) + Raw Data (Stage, Expected Delta ARR) for the commit / best-case split; Goals tab")
chart_placeholder(s, MX, CONTENT_TOP, Inches(7.4), Inches(3.0), f"Path to goal {Q1}",
                  "Optional: build a stacked bar (Won | Commit | Best case | Gap) vs Goal in the cockpit and link it here")
table(s, MX + Inches(7.7), CONTENT_TOP, CONTENT_W - Inches(7.7), ["Component", Q1, "Basis"],
      [["Goal", tok("f1.revenueGoal"), "Goals tab"], [f"Won to date ({tok('f1.wonCount')})", tok("f1.wonArr"), "Closed Won in quarter"],
       ["Commit (stages 4-5)", tok("f1.commitArr"), "Expected Delta ARR"], ["Best case (stage 3)", tok("f1.bestCaseArr"), "Expected Delta ARR"],
       ["Early (stages 0-2)", tok("f1.earlyArr"), "not counted"], ["Gap to goal", tok("f1.gapArr"), "Goal - Won - Commit"]],
      col_w=[1.9, 1.1, 1.6], size=9, row_h=Inches(0.33), bold_first_col=True)
callout(s, MX, CONTENT_TOP + Inches(3.2), Inches(7.4), Inches(1.1), "How the gap closes",
        "[Named deals that must pull in, expansion to create, or an explicit 'we will miss by $X.XM' - no unnamed upside.]",
        accent=AMBER)
callout(s, MX + Inches(7.7), CONTENT_TOP + Inches(2.3), CONTENT_W - Inches(7.7), Inches(2.0), "Logo path",
        f"Goal {tok('f1.logoGoal')}  ·  net forecast {tok('f1.logoForecast')} from {tok('f1.openLogoMajors')} open Major logo opps  ·  "
        f"gap {tok('f1.logoGap')}.  [Which accounts.]", accent=GREEN)
notes(s, "The stage split is a convention for the deck (commit = stages 4-5, best case = stage 3 Tech Validation); "
         "state it on the slide if the team uses different stage cut-offs.")

# 18 - Sales leadership look ahead --------------------------------------------------------------------
s = new_slide("05 · Leadership look ahead", f"Sales leadership commentary - plan for {Q1}", status=False)
text(s, W - MX - Inches(4.1), Inches(0.3), Inches(4.1), Inches(0.3), "[Sales leader name]  ·  " + Q1, size=9, color=MUTED,
     font=MONO, align=PP_ALIGN.RIGHT)
callout(s, MX, CONTENT_TOP, half, Inches(1.9), "1 · Commit and upside",
        f"Cockpit: won {tok('f1.wonArr')}  ·  commit {tok('f1.commitArr')}  ·  best case {tok('f1.bestCaseArr')}  ·  goal "
        f"{tok('f1.revenueGoal')} / {tok('f1.logoGoal')} logos.  [Your commit and upside in $ and logos, and the named deals "
        "behind each number. This is the line item compared on next QBR's slide 11.]", accent=GREEN)
callout(s, MX + half + Inches(0.3), CONTENT_TOP, half, Inches(1.9), "2 · How we close the gap",
        "[Pipeline creation plan (meetings / pipeline created targets per rep), partner-sourced pipeline, expansion "
        "plays on the top 10 customers, pricing / packaging moves.]", accent=AMBER)
callout(s, MX, CONTENT_TOP + Inches(2.1), half, Inches(1.9), "3 · Team and coverage",
        "[Reps in seat vs plan, ramp status, accounts uncovered ('Accts owned' / 'Acct coverage' in the Reps table), "
        "hiring needed and by when (detail slide 21).]", accent=BLUE)
callout(s, MX + half + Inches(0.3), CONTENT_TOP + Inches(2.1), half, Inches(1.9), "4 · Risks and what would change the forecast",
        "[Top 3 risks with $ and probability; leading indicators to watch monthly; the one thing that would make you "
        "raise or cut the commit.]", accent=RED)
notes(s, "Asks from Sales go on slide 22 (single table) - do not put them here.")

# 19 - Engineering leadership look ahead --------------------------------------------------------------
s = new_slide("05 · Leadership look ahead", f"Engineering leadership commentary - capacity and dependencies for {Q1}", status=False)
text(s, W - MX - Inches(4.1), Inches(0.3), Inches(4.1), Inches(0.3), "[Engineering leader name]  ·  " + Q1, size=9,
     color=MUTED, font=MONO, align=PP_ALIGN.RIGHT)
table(s, MX, CONTENT_TOP, Inches(6.6), ["Pilot demand vs DE capacity", Q1, Q2],
      [["Pilots active + planned (from pipeline slide 15)", "N", "N"], ["DE hours / pilot (avg, last quarter)", "N", "N"],
       ["DE capacity (FTE in seat x hours)", "N", "N"], ["Utilisation", "NN%", "NN%"], ["Pilots we cannot staff", "N", "N"]],
      col_w=[3.6, 1.5, 1.5], size=9, row_h=Inches(0.36), bold_first_col=True)
callout(s, MX, CONTENT_TOP + Inches(2.4), Inches(6.6), Inches(1.9), "1 · Roadmap / product dependencies on deals",
        "[Feature or integration each top deal / renewal depends on, owner, expected date, and what happens to the "
        "deal if it slips.]", accent=BLUE)
callout(s, MX + Inches(6.9), CONTENT_TOP, CONTENT_W - Inches(6.9), Inches(2.05), "2 · Deployment and technical risk on the book",
        "[Accounts with security reviews, migrations, or environment work due next quarter; renewals with technical "
        "risk (slide 16).]", accent=AMBER)
callout(s, MX + Inches(6.9), CONTENT_TOP + Inches(2.25), CONTENT_W - Inches(6.9), Inches(2.05), "3 · Hiring, enablement, local support",
        "[DE / support hiring plan vs demand, partner enablement, language / time-zone coverage. Asks go on slide 22.]",
        accent=GREEN)
notes(s, "Pilot demand comes from the Active Pilots table plus opps expected to enter Tech Validation (stage 2 with "
         "a pilot planned). Capacity numbers are the engineering leader's.")

# 20 - Rep performance --------------------------------------------------------------------------------
s = new_slide("06 · Team", f"[Headline: e.g. '{tok('repsOnPace')} of {tok('repCount')} reps at or above pace; activity coverage lowest on ramping reps']",
              source=f"{TEAM} tab > REP ACTIVITY & PERFORMANCE > 'Reps' table (Months in seat, Accts owned, Rep goal (FY), Won ARR (QTD / FY), "
                     f"Attainment (FY), Meetings, Activities, Acct coverage, Pipeline created (QTD), Stalled >60d, Renewal risk (FY))")
table(s, MX, CONTENT_TOP, CONTENT_W,
      ["Rep", "Months in seat", "Accts", "Goal (FY)", "Won ARR (QTD)", "Won ARR (FY)", "Attain. (FY)", "Meetings (QTD)",
       "Activities (QTD)", "Acct cov. (QTD)", "Pipe created (QTD)", "Stalled >60d", "Manager read"],
      rows_for("reps", [""] * 11 + ["[on track / coach / ramping]"], n=8),
      col_w=[1.6, 0.8, 0.6, 0.9, 1.0, 0.9, 0.8, 0.8, 0.9, 0.8, 1.0, 0.9, 1.8], size=8, row_h=Inches(0.33),
      bold_first_col=True, guidance_cols=(12,))
guidance(s, MX, CONTENT_TOP + Inches(3.15), CONTENT_W, Inches(0.7),
         "Rep goal and attainment are FY-level (Salesforce rep goals have no quarter); activity, Won ARR and pipeline "
         "created are for the selected quarter to date (header reads '(Q)' once closed). A rep shown as 'Name (team)' "
         "owns opps on this team's accounts but sits in another team - attribution stays at account level. Meetings are "
         "credited to the event owner (no participant credit); slippage rate is not reproduced.")
callout(s, MX, CONTENT_TOP + Inches(3.85), CONTENT_W, Inches(0.55), "So what",
        "[Who needs what: coaching, territory change, pipeline help, recognition.]", accent=GREEN)
notes(s, "Filled from the Reps table (sorted by Won ARR QTD; names link to Salesforce). The 'Manager read' column is the "
         "only manual column and replaces the old free-text stack rank.")

# 21 - Hiring / capacity ------------------------------------------------------------------------------
s = new_slide("06 · Team", "[Headline: e.g. 'N of N planned heads in seat; DE hiring is the gating item for Q+1 pilots']",
              source="Manual (recruiting system) - not in the cockpit", status=False)
table(s, MX, CONTENT_TOP, CONTENT_W,
      ["Role", "Team", "# HC", "In seat", "Offer / onsite", "Target in-seat", "Impact if late (deals / pilots / accounts)", "Status"],
      [["[Account Director]", TEAM, "N", "N", "N / N", "[Qx]", "[e.g. N Major accounts uncovered, $X.XM pipeline]", "On track / At risk"],
       ["[Deployed Engineer]", TEAM, "N", "N", "N / N", "[Qx]", "[e.g. N pilots unstaffed in Q+1]", "At risk"],
       ["[Solutions / Support]", TEAM, "N", "N", "N / N", "[Qx]", "[...]", "Not started"],
       ["[...]", "", "", "", "", "", "", ""]],
      col_w=[1.9, 1.3, 0.6, 0.7, 1.1, 1.0, 3.6, 1.2], size=8.5, row_h=Inches(0.38), bold_first_col=True, guidance_cols=(6,))
callout(s, MX, CONTENT_TOP + Inches(2.3), half, Inches(1.6), "Capacity math",
        "[Accounts per rep today vs target; pilots per DE; what the plan assumes about ramp time.]", accent=BLUE)
callout(s, MX + half + Inches(0.3), CONTENT_TOP + Inches(2.3), half, Inches(1.6), "Blockers",
        "[Recruiting pipeline volume, comp, location, process - and the ask, if any, on slide 22.]", accent=AMBER)
notes(s, "Manual slide: headcount comes from the recruiting system / hiring plan, not the cockpit. Tie 'Impact if late' "
         "to cockpit facts (uncovered Major accounts, unstaffed Active Pilots, Q+1 pipeline).")

# 22 - Asks & decisions -------------------------------------------------------------------------------
s = new_slide("07 · Asks & decisions", "Decisions needed from leadership - owners and dates", status=False)
section_label(s, MX, CONTENT_TOP, CONTENT_W, "Sales asks")
asks_table(s, CONTENT_TOP + Inches(0.35), rows=3)
section_label(s, MX, CONTENT_TOP + Inches(1.75), CONTENT_W, "Engineering asks")
asks_table(s, CONTENT_TOP + Inches(2.1), rows=3)
section_label(s, MX, CONTENT_TOP + Inches(3.5), CONTENT_W, "Carried over from last QBR - status update")
table(s, MX, CONTENT_TOP + Inches(3.85), CONTENT_W, ["Ask / decision (last QBR)", "Outcome", "Owner", "Closed on", "Status"],
      [["[ask]", "[what was decided / delivered]", "[name]", "[date]", "Done / In progress / Dropped"] for _ in range(2)],
      col_w=[3.2, 3.4, 1.2, 1.1, 1.0], guidance_cols=(0, 1, 2, 3, 4))
guidance(s, MX, CONTENT_TOP + Inches(4.8), CONTENT_W, Inches(0.5),
         "Every ask names the deal / ARR it protects or unlocks, a decision owner and a date. Carry the table to the next "
         "QBR with the Status column updated (Done / In progress / Dropped) - this is the accountability loop the old deck lacked.")
notes(s, "This is the meeting's output. Review it first if time is short.")

# 23 - Appendix: cockpit mapping ---------------------------------------------------------------------
s = new_slide("Appendix", "Cockpit to deck mapping - where every number comes from",
              source="QBR cockpit workbook: team tabs (B1 team, B2 quarter), Raw Data, Goals, ARR Ledger, Definitions", status=False)
mapping = [
    ["4 · Executive summary", "KPI cards (row 5) + metric block", "Formulas over Raw Data / Goals / ARR Ledger keyed to B1, B2"],
    ["5 · Scorecard", "'ACTUALS | FORECAST QUARTER' block: Metric / Quarter / QoQ / Trend", "QoQ basis in the section header"],
    ["6 · ARR bridge", "ARR bridge block + 'ARR bridge' chart", "Starting / Ending = ARR Ledger"],
    ["7 · Deals won", "'Top 10 Deals Won' table", "Links to Salesforce opps; TCV = Amount"],
    ["8 · Logos", "'Logo Goal', 'Logos Won (Land only / Land + MSP)', 'Logo Attainment (expected)', 'Logos Won' table", "Open Major logo opps from Raw Data"],
    ["9 · Retention", "'Churned', 'Downgrade', 'Renewals Won / Lost' tables; renewal & churn rows; 'Renewals' chart", ""],
    ["10 · Pilots", "'# Active Pilots', '# Pilots Completed', 'Active Pilots', 'Pilots Completed' tables", "Accounts block for conversion"],
    ["13 · Forecast", "FUTURE QUARTER(S) block, 'Q+1 / Q+2 forecast vs goal' chart", "Net forecast = Expected Delta ARR; pipeline = open opps"],
    ["14 · Pipeline quality", "'Pipeline by stage' tables (Q, Q+1, Q+2)", "Stalled >60d from Reps table"],
    ["15 · Deals to win", "'Top 10 Deals Q+1 / Q+2' tables", ""],
    ["16 · Renewals at risk", "'Renewals due Q+1 / Q+2', 'Predicted Churn Q+1 / Q+2' tables", ""],
    ["17 · Path to goal", "FUTURE QUARTER(S) block + Raw Data (Stage, Expected Delta ARR)", "Commit / best-case split is a deck convention"],
    ["20 · Rep performance", "REP ACTIVITY & PERFORMANCE > 'Reps' table", "Goal / attainment FY, activity QTD"],
    ["25 · Data quality", "OWNERS & DATA QUALITY > 'Data quality' table", ""],
]
table(s, MX, CONTENT_TOP, CONTENT_W, ["Slide", "Cockpit block / table (team tab unless stated)", "Notes"], mapping,
      col_w=[2.0, 6.4, 3.6], size=8.5, row_h=Inches(0.3), bold_first_col=True)
notes(s, "Reference slide - nothing to fill. Keep it in the pre-read so readers can audit any number against the cockpit.")

# 24 - Appendix: definitions --------------------------------------------------------------------------
s = new_slide("Appendix", "Definitions and reading rules", source="Definitions tab of the cockpit (authoritative wording)", status=False)
defs = [
    ("Net Added ARR", "Delta ARR of Closed Won opps + Delta ARR of Closed Lost renewals (full churn), close date in the quarter. Closed Won MSP with NACV = 0 uses the opp's ARR."),
    ("Attainment / pace", "Net Added ARR / Revenue Goal; pace = attainment / quarter elapsed (100% = straight line)."),
    ("ACTUALS vs FORECAST", "All quarter numbers are FORECAST while the quarter is in progress (label shows % elapsed) and ACTUALS from the day after quarter end. Nothing is frozen: refresh re-reads Salesforce."),
    ("QoQ (same point)", "While the quarter is open, previous quarter counted only to the same elapsed day; full quarter vs full quarter once closed. Ledger balances, open pipeline, coverage and conversion have no same-point value ('-')."),
    ("Logos Won", "Land only = Closed Won Land on Major accounts; Land + MSP also counts MSP deals on new accounts. Logo attainment (expected) sums Expected Logo Impact incl. open opps."),
    ("Renewal rate / churn", "Won renewals / closed renewals in the quarter (count). Churn ARR = full churn + downgrades; GRR / NRR of Starting ARR."),
    ("Pilots", "Active = open opps in '3- Tech Validation'; Completed = Pilot Status 'Complete' with actual end date in the quarter."),
    ("Attribution", "Opportunities belong to the team of their account (account-level Team / Sub-team), whatever the opp owner's team."),
    ("Rep metrics", "Won ARR, meetings, activities, coverage and pipeline created are quarter to date; rep goal and attainment are FY (no quarterly rep goal in Salesforce)."),
]
half = (CONTENT_W - Inches(0.3)) / 2
for i, (k, v) in enumerate(defs):
    col = i % 2
    row = i // 2
    x = MX + col * (half + Inches(0.3))
    y = CONTENT_TOP + row * Inches(0.98)
    text(s, x, y, half, Inches(0.25), k, size=9.5, color=INK, bold=True)
    text(s, x, y + Inches(0.24), half, Inches(0.7), v, size=8.5, color=MUTED)
notes(s, "Nothing to fill. If a definition changes in the cockpit's Definitions tab, change it here in the same PR.")

# 25 - Appendix: data quality -------------------------------------------------------------------------
s = new_slide("Appendix", "Data quality - what would change these numbers",
              source=f"{TEAM} tab > OWNERS & DATA QUALITY > 'Data quality' table (Issue, Account, Opportunity, Detail, Close date, Owner)", status=False)
table(s, MX, CONTENT_TOP, CONTENT_W, ["Issue", "Account", "Detail", "Close date", "Owner", "Fix by"],
      rows_for("dataQuality", ["", "", "", "", "[date]"], n=8),
      col_w=[2.8, 2.2, 3.6, 1.0, 1.4, 1.0], size=8.5, row_h=Inches(0.32), bold_first_col=True, guidance_cols=(5,))
guidance(s, MX, CONTENT_TOP + Inches(3.1), CONTENT_W, Inches(0.5),
         "Own the hygiene: each open issue has an owner and a fix-by date. Known model limits (not issues): no per-quarter "
         "account history, rep goals FY-only, MSP Delta ARR = opp ARR when NACV is 0.")
notes(s, "Filled from the 'Data quality' table (OWNERS & DATA QUALITY section); 'Fix by' is the only manual column. Rows "
         "here explain why a number may move on the next refresh (e.g. a Closed Won opp with no Delta ARR).")

out = sys.argv[1] if len(sys.argv) > 1 else "QBR_Deck_Template.pptx"
prs.save(out)
print(f"{out}: {len(prs.slides)} slides")
