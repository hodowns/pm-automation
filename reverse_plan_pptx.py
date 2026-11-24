"""
Generates Reverse-Plan-Skill.pptx on the Desktop.
Run: py -3 reverse_plan_pptx.py
"""

import os
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── palette ───────────────────────────────────────────────────────────────────
C_BG      = RGBColor(0x07, 0x18, 0x36)
C_SURFACE = RGBColor(0x0E, 0x28, 0x4A)
C_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
C_LIGHT   = RGBColor(0xD8, 0xE8, 0xFF)
C_SUB     = RGBColor(0x68, 0x98, 0xCC)
C_AUTO    = RGBColor(0x20, 0x68, 0xC8)
C_AUTO_BG = RGBColor(0x0C, 0x22, 0x48)
C_DEC     = RGBColor(0xF0, 0xA0, 0x30)
C_DEC_BG  = RGBColor(0x2C, 0x1E, 0x06)
C_NEW     = RGBColor(0x20, 0xB8, 0x98)
C_NEW_BG  = RGBColor(0x04, 0x18, 0x14)
C_EX      = RGBColor(0x90, 0x60, 0xF0)
C_EX_BG   = RGBColor(0x0A, 0x06, 0x1E)
C_GATE    = RGBColor(0xF0, 0x68, 0x58)
C_GATE_BG = RGBColor(0x20, 0x08, 0x06)
C_PROC    = RGBColor(0x50, 0x90, 0xE8)
C_PROC_BG = RGBColor(0x06, 0x10, 0x28)
C_CAL     = RGBColor(0x90, 0x80, 0xE0)
C_OUT     = RGBColor(0x30, 0xC8, 0x70)
C_OUT_BG  = RGBColor(0x03, 0x0C, 0x08)
C_ACCENT  = RGBColor(0x38, 0xB8, 0xF0)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


# ── slide factory ─────────────────────────────────────────────────────────────
def new_slide(prs, layout_idx=6):
    layout = prs.slide_layouts[layout_idx]
    slide = prs.slides.add_slide(layout)
    for ph in slide.placeholders:
        ph._element.getparent().remove(ph._element)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = C_BG
    return slide


# ── drawing primitives ────────────────────────────────────────────────────────
def box(slide, x, y, w, h, bg=None, border=None):
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid() if bg else shape.fill.background()
    if bg:
        shape.fill.fore_color.rgb = bg
    if border:
        shape.line.color.rgb = border
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape


def label(slide, text, x, y, w, h, size=12, bold=False, color=None,
          align=PP_ALIGN.LEFT, italic=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.word_wrap = True
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color or C_WHITE
    return tb


def _add_arrowhead(cxn):
    """Append a tail arrowhead to a connector via lxml."""
    try:
        ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        ln = cxn.line._ln          # creates <a:ln> if absent
        for te in ln.findall(f'{{{ns}}}tailEnd'):
            ln.remove(te)
        te = etree.SubElement(ln, f'{{{ns}}}tailEnd')
        te.set('type', 'arrow')
        te.set('w', 'med')
        te.set('len', 'med')
    except Exception:
        pass


def arrow_down(slide, cx, y1, y2, color=None, thick=1.8):
    c = color or C_SUB
    cxn = slide.shapes.add_connector(1, Inches(cx), Inches(y1), Inches(cx), Inches(y2))
    cxn.line.color.rgb = c
    cxn.line.width = Pt(thick)
    _add_arrowhead(cxn)


def arrow_right(slide, x1, y, x2, color=None, thick=1.8):
    c = color or C_SUB
    cxn = slide.shapes.add_connector(1, Inches(x1), Inches(y), Inches(x2), Inches(y))
    cxn.line.color.rgb = c
    cxn.line.width = Pt(thick)
    _add_arrowhead(cxn)


def line_seg(slide, x1, y1, x2, y2, color=None, thick=1.8):
    """Connector with no arrowhead."""
    c = color or RGBColor(0x30, 0x58, 0x88)
    cxn = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cxn.line.color.rgb = c
    cxn.line.width = Pt(thick)


def card(slide, x, y, w, h, title, title_color, body_lines,
         bg=None, border=None, eyebrow=None, eyebrow_color=None):
    box(slide, x, y, w, h, bg=bg or C_SURFACE, border=border or C_PROC)
    cy = y + 0.12
    if eyebrow:
        label(slide, eyebrow, x+0.15, cy, w-0.3, 0.2,
              size=7, bold=True, color=eyebrow_color or C_SUB)
        cy += 0.2
    label(slide, title, x+0.15, cy, w-0.3, 0.28, size=11, bold=True, color=title_color)
    cy += 0.3
    for ln in body_lines:
        label(slide, ln, x+0.15, cy, w-0.3, 0.22, size=9, color=C_LIGHT)
        cy += 0.22


# ── SLIDE 1: Title ────────────────────────────────────────────────────────────
def slide_title(prs):
    s = new_slide(prs)
    box(s, 0, 0, 0.06, 7.5, bg=C_ACCENT)
    label(s, "Reverse Plan Skill", 0.5, 1.6, 8, 1.2, size=44, bold=True, color=C_WHITE)
    label(s, "Cisco CX PM Automation", 0.5, 2.8, 8, 0.4, size=14, color=C_SUB)
    label(s, "End-to-end visual reference: from /reverse-plan to a live card on the dashboard.",
          0.5, 3.4, 9, 0.5, size=13, color=C_LIGHT)
    box(s, 0.5, 4.1, 2.2, 0.45, bg=C_PROC_BG, border=C_PROC)
    label(s, "TRIGGER", 0.65, 4.12, 0.8, 0.2, size=7, bold=True, color=C_SUB)
    label(s, "/reverse-plan", 0.65, 4.3, 2.0, 0.22, size=12, bold=True, color=C_ACCENT)
    box(s, 0.5, 4.75, 1.8, 0.35, bg=C_NEW_BG, border=C_NEW)
    label(s, "● New Plan path", 0.65, 4.8, 1.6, 0.25, size=9, bold=True, color=C_NEW)
    box(s, 2.5, 4.75, 2.2, 0.35, bg=C_EX_BG, border=C_EX)
    label(s, "● Edit Existing path", 2.65, 4.8, 2.0, 0.25, size=9, bold=True, color=C_EX)
    label(s, "Both paths share the same validation, build, and publish steps.",
          0.5, 6.6, 10, 0.4, size=10, italic=True, color=C_SUB)


# ── SLIDE 2: Full Workflow — linear flowchart ─────────────────────────────────
def slide_overview(prs):
    s = new_slide(prs)

    # ── layout constants ──
    CX   = 6.67    # horizontal center of slide
    BW   = 5.6     # width of center-column boxes
    BX   = CX - BW / 2   # left edge of center boxes
    # column geometry
    COL_W = 5.65
    LX    = 0.35               # left col left edge
    RX    = 13.33 - 0.35 - COL_W  # right col left edge  = 7.33
    LCX   = LX + COL_W / 2    # left col center x  ≈ 3.18
    RCX   = RX + COL_W / 2    # right col center x ≈ 10.155

    # ── title ──
    label(s, "Full Workflow at a Glance", 0.4, 0.1, 12.5, 0.38,
          size=18, bold=True, color=C_WHITE)
    label(s, "Read top to bottom. Both paths merge before the shared gate steps.",
          0.4, 0.48, 12.5, 0.26, size=10, color=C_SUB)

    # ─────────────────────────────────────────────────────────────
    # 1  START BOX
    # ─────────────────────────────────────────────────────────────
    Y_START = 0.82
    box(s, BX, Y_START, BW, 0.52, bg=C_AUTO_BG, border=C_AUTO)
    label(s, "START", BX+0.18, Y_START+0.04, BW-0.36, 0.17,
          size=7, bold=True, color=C_AUTO)
    label(s, "Need to create or update a Cisco CX reverse plan",
          BX+0.18, Y_START+0.22, BW-0.36, 0.26,
          size=11, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    # ─────────────────────────────────────────────────────────────
    # arrow 1
    # ─────────────────────────────────────────────────────────────
    Y_A1 = Y_START + 0.52
    arrow_down(s, CX, Y_A1, Y_A1 + 0.18)

    # ─────────────────────────────────────────────────────────────
    # 2  DECISION  (rectangle styled as decision)
    # ─────────────────────────────────────────────────────────────
    Y_DEC = Y_A1 + 0.18
    H_DEC = 0.66
    box(s, BX, Y_DEC, BW, H_DEC, bg=C_DEC_BG, border=C_DEC)
    label(s, "⬦  DECISION", BX+0.18, Y_DEC+0.05, BW-0.36, 0.18,
          size=7, bold=True, color=C_DEC)
    label(s, "New Plan  or  Edit Existing?",
          BX+0.18, Y_DEC+0.24, BW-0.36, 0.32,
          size=13, bold=True, color=C_DEC, align=PP_ALIGN.CENTER)

    # ─────────────────────────────────────────────────────────────
    # FORK  (vertical drop → horizontal spread → two arrows down)
    # ─────────────────────────────────────────────────────────────
    Y_DEC_BOT  = Y_DEC + H_DEC
    Y_FORK_H   = Y_DEC_BOT + 0.18   # horizontal crossbar y
    Y_COL_TOP  = Y_FORK_H + 0.18    # top of two columns

    line_seg(s, CX, Y_DEC_BOT, CX, Y_FORK_H)           # short drop
    line_seg(s, LCX, Y_FORK_H, RCX, Y_FORK_H)          # crossbar
    arrow_down(s, LCX, Y_FORK_H, Y_COL_TOP, color=C_NEW)  # left arrow ↓
    arrow_down(s, RCX, Y_FORK_H, Y_COL_TOP, color=C_EX)   # right arrow ↓

    # branch labels
    label(s, "New Plan", LX, Y_FORK_H - 0.03, 2.0, 0.22,
          size=9, bold=True, color=C_NEW)
    label(s, "Edit Existing", RX + COL_W - 2.2, Y_FORK_H - 0.03, 2.2, 0.22,
          size=9, bold=True, color=C_EX, align=PP_ALIGN.RIGHT)

    # ─────────────────────────────────────────────────────────────
    # 3  TWO PATH COLUMNS
    # ─────────────────────────────────────────────────────────────
    H_COL = 1.72
    Y_COL = Y_COL_TOP

    # ── NEW PLAN column ──
    box(s, LX, Y_COL, COL_W, H_COL, bg=C_NEW_BG, border=C_NEW)
    label(s, "NEW PLAN PATH", LX+0.18, Y_COL+0.07, COL_W-0.36, 0.2,
          size=8, bold=True, color=C_NEW)

    new_steps = [
        ("①  Choose working mode",
         "Step-by-step  (Claude guides field-by-field)\nor  Describe-it  (brain-dump, Claude fills gaps)"),
        ("②  Share project collateral  (optional)",
         "SOW · project plan URL · Webex space name\nClaude pre-fills as many fields as possible"),
        ("③  Gather 13 project fields",
         "Market · PID · Milestone · Revenue · Revenue Type\nCompletion Month · Partner · Customer Contact\nEscalation Path · Review Windows · Internal Review"),
    ]
    ny = Y_COL + 0.32
    for i, (t, d) in enumerate(new_steps):
        label(s, t, LX+0.2, ny, COL_W-0.4, 0.22, size=8.5, bold=True, color=C_WHITE)
        ny += 0.22
        for ln in d.split('\n'):
            label(s, ln, LX+0.2, ny, COL_W-0.4, 0.19, size=8, color=C_LIGHT)
            ny += 0.19
        if i < 2:
            label(s, "↓", LX + COL_W/2 - 0.1, ny, 0.2, 0.16, size=9,
                  color=C_NEW, align=PP_ALIGN.CENTER)
            ny += 0.17

    # ── EXISTING PLAN column ──
    box(s, RX, Y_COL, COL_W, H_COL, bg=C_EX_BG, border=C_EX)
    label(s, "EDIT EXISTING PATH", RX+0.18, Y_COL+0.07, COL_W-0.36, 0.2,
          size=8, bold=True, color=C_EX)

    ex_steps = [
        ("①  List plans in PLAN_DIR",
         "Claude lists every .json file found locally\nand numbers them for easy selection"),
        ("②  Pick plan  —  or fetch from GitHub",
         "If not found locally, provide PID or SharePoint link\nand Claude fetches it via gh CLI"),
        ("③  Choose working mode",
         "Step-by-step  or  Describe-it\nDescribe-it is ideal: say what changed,\nClaude applies it without re-asking set fields"),
    ]
    ey = Y_COL + 0.32
    for i, (t, d) in enumerate(ex_steps):
        label(s, t, RX+0.2, ey, COL_W-0.4, 0.22, size=8.5, bold=True, color=C_WHITE)
        ey += 0.22
        for ln in d.split('\n'):
            label(s, ln, RX+0.2, ey, COL_W-0.4, 0.19, size=8, color=C_LIGHT)
            ey += 0.19
        if i < 2:
            label(s, "↓", RX + COL_W/2 - 0.1, ey, 0.2, 0.16, size=9,
                  color=C_EX, align=PP_ALIGN.CENTER)
            ey += 0.17

    # ─────────────────────────────────────────────────────────────
    # MERGE  (two lines up → horizontal → center arrow down)
    # ─────────────────────────────────────────────────────────────
    Y_COL_BOT  = Y_COL + H_COL
    Y_MERGE_H  = Y_COL_BOT + 0.18
    Y_GATE_TOP = Y_MERGE_H + 0.2

    line_seg(s, LCX, Y_COL_BOT, LCX, Y_MERGE_H, color=C_NEW)
    line_seg(s, RCX, Y_COL_BOT, RCX, Y_MERGE_H, color=C_EX)
    line_seg(s, LCX, Y_MERGE_H, RCX, Y_MERGE_H)
    arrow_down(s, CX, Y_MERGE_H, Y_GATE_TOP)

    label(s, "← both paths continue →",
          CX - 1.5, Y_MERGE_H - 0.04, 3.0, 0.18,
          size=8, italic=True, color=C_SUB, align=PP_ALIGN.CENTER)

    # ─────────────────────────────────────────────────────────────
    # 4  SHARED GATE STEPS  (horizontal row of 5 boxes)
    # ─────────────────────────────────────────────────────────────
    Y_GATE = Y_GATE_TOP
    H_GATE = 0.72
    GBW = 2.06   # gate box width
    GAP = 0.22   # gap between gate boxes
    # 5 boxes: total = 5*GBW + 4*GAP = 10.3 + 0.88 = 11.18
    G_START_X = (13.33 - (5 * GBW + 4 * GAP)) / 2  # ≈ 1.075

    gate_boxes = [
        ("⛔", "STEP 2", "Validate\nAll Inputs", C_GATE,   C_GATE_BG),
        ("📄", "STEP 3", "Generate\nFilename",   C_PROC,   C_PROC_BG),
        ("📅", "STEP 4", "FY27\nCalendar",       C_CAL,    RGBColor(0x0E, 0x0C, 0x2A)),
        ("✅", "STEP 5", "Pre-build\nReview",    C_PROC,   C_PROC_BG),
        ("🔨", "STEP 6–7", "Build Activity\nTable", C_ACCENT, RGBColor(0x04, 0x18, 0x2C)),
    ]

    for i, (icon, step, title, tc, bg) in enumerate(gate_boxes):
        gx = G_START_X + i * (GBW + GAP)
        box(s, gx, Y_GATE, GBW, H_GATE, bg=bg, border=tc)
        label(s, step, gx+0.1, Y_GATE+0.05, GBW-0.2, 0.17, size=7, bold=True, color=tc)
        label(s, f"{icon}  {title}",
              gx+0.1, Y_GATE+0.22, GBW-0.2, 0.42, size=9, bold=True, color=C_WHITE)
        # arrow between boxes
        if i < 4:
            ax = gx + GBW
            arrow_right(s, ax, Y_GATE + H_GATE/2, ax + GAP, color=C_SUB)

    # ─────────────────────────────────────────────────────────────
    # arrow from BUILD box center-bottom → output row
    # ─────────────────────────────────────────────────────────────
    BUILD_CX = G_START_X + 4 * (GBW + GAP) + GBW / 2
    Y_GATE_BOT = Y_GATE + H_GATE
    Y_OUT_TOP  = Y_GATE_BOT + 0.38

    # draw: build-bottom → horizontal to center → center → output
    line_seg(s, BUILD_CX, Y_GATE_BOT, BUILD_CX, Y_GATE_BOT + 0.2)
    line_seg(s, BUILD_CX, Y_GATE_BOT + 0.2, CX, Y_GATE_BOT + 0.2)
    arrow_down(s, CX, Y_GATE_BOT + 0.2, Y_OUT_TOP)

    # ─────────────────────────────────────────────────────────────
    # 5  OUTPUT ROW  (3 destination cards)
    # ─────────────────────────────────────────────────────────────
    OBW = 3.55
    OBH = 0.72
    OBG = 0.3
    OBX = (13.33 - 3 * OBW - 2 * OBG) / 2  # ≈ 1.035

    out_items = [
        ("💾", "STEP 8a–b", "Local JSON copy",   C_OUT,  C_OUT_BG,
         "PLAN_DIR/*.json · saved immediately"),
        ("📤", "STEP 8c",   "GitHub Pull Request", C_PROC, C_PROC_BG,
         "cx-usps-auto/Reverse-plans · awaits merge"),
        ("🟢", "LIVE",      "Dashboard",           C_OUT,  C_OUT_BG,
         "~30 sec after merge · wwwin-github.cisco.com"),
    ]
    for i, (icon, step, title, tc, bg, desc) in enumerate(out_items):
        ox = OBX + i * (OBW + OBG)
        box(s, ox, Y_OUT_TOP, OBW, OBH, bg=bg, border=tc)
        label(s, step, ox+0.15, Y_OUT_TOP+0.05, OBW-0.3, 0.17, size=7, bold=True, color=tc)
        label(s, f"{icon}  {title}",
              ox+0.15, Y_OUT_TOP+0.22, OBW-0.3, 0.26, size=10, bold=True, color=tc)
        label(s, desc, ox+0.15, Y_OUT_TOP+0.48, OBW-0.3, 0.2, size=8, color=C_LIGHT)
        if i == 1:  # GitHub PR → Dashboard arrow
            arrow_right(s, ox + OBW, Y_OUT_TOP + OBH/2,
                        ox + OBW + OBG, color=C_PROC)

    # update cadence note
    label(s, "Update EOD every Friday — mark steps Complete with real Actual dates",
          OBX, Y_OUT_TOP + OBH + 0.1, 3*OBW + 2*OBG, 0.22,
          size=8.5, italic=True, color=C_SUB, align=PP_ALIGN.CENTER)


# ── SLIDE 3: Startup ──────────────────────────────────────────────────────────
def slide_startup(prs):
    s = new_slide(prs)
    label(s, "Step 0 — Startup (silent & automatic)", 0.4, 0.15, 12, 0.4,
          size=18, bold=True, color=C_WHITE)
    label(s, "These two steps run every time before anything else.", 0.4, 0.55, 10, 0.3,
          size=11, color=C_SUB)

    box(s, 0.4, 1.1, 5.8, 4.5, bg=C_AUTO_BG, border=C_AUTO)
    label(s, "AUTO-UPDATE CHECK", 0.6, 1.15, 5.5, 0.25, size=8, bold=True, color=C_AUTO)
    label(s, "Silently pulls latest skill from git", 0.6, 1.4, 5.4, 0.3,
          size=13, bold=True, color=C_WHITE)
    body1 = [
        "Runs a git pull --ff-only on the Reverse-plans repo.",
        "Only updates if the local clone is on the main branch.",
        "If the clone is on a feature branch: warns you to switch",
        "  (git switch main && git pull) but continues normally.",
        "",
        "If the skill was updated → tells you, reloads, restarts.",
        "If no change → says nothing, continues.",
        "",
        "Times out after ~20 seconds so being off VPN never blocks.",
    ]
    cy = 1.78
    for line in body1:
        label(s, line, 0.6, cy, 5.4, 0.26, size=9, color=C_LIGHT)
        cy += 0.26
    label(s, "WHY: Ensures every PM always runs the current version without manual updates.",
          0.6, 4.8, 5.4, 0.5, size=8.5, italic=True, color=C_SUB)

    box(s, 6.8, 1.1, 6.1, 4.5, bg=C_PROC_BG, border=C_PROC)
    label(s, "PLAN_DIR SETUP", 7.0, 1.15, 5.8, 0.25, size=8, bold=True, color=C_PROC)
    label(s, "Resolve your local plans folder", 7.0, 1.4, 5.6, 0.3,
          size=13, bold=True, color=C_WHITE)
    body2 = [
        "Checks for these folders in order:",
        "",
        "  1.  ~/Documents/Automation/PM-Reverse-Plans/",
        "  2.  ~/reverse-plans/",
        "",
        "Uses the first one that exists.",
        "If neither exists → creates ~/reverse-plans/ and tells you.",
        "",
        "All JSON files are read from and written to this folder.",
        "Claude loads your local plan automatically when editing.",
    ]
    cy = 1.78
    for line in body2:
        label(s, line, 7.0, cy, 5.7, 0.26, size=9, color=C_LIGHT)
        cy += 0.26
    label(s, "WHY: One consistent local copy of every plan — no searching for the latest version.",
          7.0, 4.8, 5.7, 0.5, size=8.5, italic=True, color=C_SUB)


# ── SLIDE 4: Both paths ───────────────────────────────────────────────────────
def slide_paths(prs):
    s = new_slide(prs)
    label(s, "Step 0 — Choose Your Path", 0.4, 0.15, 12, 0.4, size=18, bold=True, color=C_WHITE)
    label(s, "First question Claude asks: \"Are you building a new reverse plan, or editing an existing one?\"",
          0.4, 0.55, 12.5, 0.3, size=11, color=C_SUB)

    box(s, 5.1, 1.0, 3.1, 0.75, bg=C_DEC_BG, border=C_DEC)
    label(s, "New Plan or Edit Existing?", 5.25, 1.1, 2.8, 0.5, size=12, bold=True,
          color=C_DEC, align=PP_ALIGN.CENTER)

    box(s, 0.3, 1.9, 5.8, 4.7, bg=C_NEW_BG, border=C_NEW)
    label(s, "NEW PLAN PATH", 0.5, 1.93, 5.5, 0.25, size=8, bold=True, color=C_NEW)
    label(s, "Building from scratch", 0.5, 2.18, 5.5, 0.28, size=12, bold=True, color=C_NEW)
    new_steps = [
        ("1", "Choose Mode", "Step-by-step: Claude walks you through 13 fields one by one.\nDescribe-it: Brain-dump everything and Claude fills the gaps."),
        ("2", "Share Project Collateral (optional)", "Project plan URL, meeting minutes, Webex space name.\nClaude reads them and pre-fills as many fields as possible."),
        ("3", "Answer any remaining fields", "Market · PID · Milestone # · Revenue · Revenue Type\nRequired Completion Month · Partner / Integrator\nCustomer Approval Contact · Escalation Path\nReview windows (default 5 biz days / 2 biz days)\nCisco internal review? (default yes, 2 biz days)"),
        ("4", "List known risks & external deps", "Claude populates Risk/Mitigation and Dependency columns from your answers + collateral."),
    ]
    cy = 2.5
    for num, title, desc in new_steps:
        box(s, 0.45, cy, 0.3, 0.28, bg=C_NEW, border=C_NEW)
        label(s, num, 0.45, cy+0.02, 0.3, 0.24, size=9, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        label(s, title, 0.82, cy, 5.1, 0.26, size=9, bold=True, color=C_WHITE)
        cy += 0.28
        for line in desc.split('\n'):
            label(s, line, 0.82, cy, 5.1, 0.22, size=8, color=C_LIGHT)
            cy += 0.22
        cy += 0.1

    box(s, 7.1, 1.9, 5.8, 4.7, bg=C_EX_BG, border=C_EX)
    label(s, "EDIT EXISTING PATH", 7.3, 1.93, 5.5, 0.25, size=8, bold=True, color=C_EX)
    label(s, "Updating an existing plan", 7.3, 2.18, 5.5, 0.28, size=12, bold=True, color=C_EX)
    ex_steps = [
        ("1", "List plans in PLAN_DIR", "Claude lists all .json files found locally:\n  1. SLED West — CoH LCS (F51119) — Milestone 9\n  2. NSO — GDNN NASSCO (F20569) — MS9007\n  ..."),
        ("2", "Pick a plan", "Choose by number. If the plan isn't local,\nprovide PID / SharePoint link and Claude fetches it from GitHub."),
        ("3", "Choose Mode", "Same two options: Step-by-step or Describe-it.\nDescribe-it is ideal for edits: say what changed and\nClaude applies it without re-asking what's already set."),
        ("4", "Apply changes", "Claude updates only what you specify.\nRequired dates (baseline) are never touched.\nDependent dates cascade only when explicitly changed."),
    ]
    cy = 2.5
    for num, title, desc in ex_steps:
        box(s, 7.25, cy, 0.3, 0.28, bg=C_EX, border=C_EX)
        label(s, num, 7.25, cy+0.02, 0.3, 0.24, size=9, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        label(s, title, 7.62, cy, 5.1, 0.26, size=9, bold=True, color=C_WHITE)
        cy += 0.28
        for line in desc.split('\n'):
            label(s, line, 7.62, cy, 5.1, 0.22, size=8, color=C_LIGHT)
            cy += 0.22
        cy += 0.1

    label(s, "↓ Both paths continue to the same validation, build, and publish steps",
          0.3, 6.85, 13.0, 0.35, size=9, italic=True, color=C_SUB, align=PP_ALIGN.CENTER)


# ── SLIDE 5: Shared gates ─────────────────────────────────────────────────────
def slide_shared(prs):
    s = new_slide(prs)
    label(s, "Steps 2–5 — Shared Gates (both paths)", 0.4, 0.15, 12, 0.4,
          size=18, bold=True, color=C_WHITE)
    label(s, "These run in order after New or Existing path is resolved.", 0.4, 0.55, 10, 0.3,
          size=11, color=C_SUB)
    gates = [
        {"eyebrow": "STEP 2 — GATE", "title": "⛔  Validate All Inputs",
         "color": C_GATE, "bg": C_GATE_BG,
         "lines": [
             "Runs before any file is built. ALL checks must pass.",
             "", "Field rules:  Market in {CIV, NSO, DOW, SLE, SLW, DIB}",
             "  PID: 1 letter + 5 digits  ·  Revenue > 0  ·  Escalation path ≥ 2 names",
             "  Required completion month maps to a real FY27 MCC month",
             "  Revenue type: committed | upside | at_risk",
             "", "Activity table rules (every row, no exceptions):",
             "  Status: Not Started / In Progress / Complete / No Longer Required",
             "  Dependency: must be reasoned and non-blank — 'None' fails",
             "  Risk/Mitigation: must state risk, recovery action, and exit criteria",
             "  Actual dates: never a future date",
         ]},
        {"eyebrow": "STEP 3", "title": "Generate Filename",
         "color": C_PROC, "bg": C_PROC_BG,
         "lines": [
             "Format:  MARKET-PID-project-slug.json",
             "Example: SLED_West-f51119-coh-lcs-milestone-9.json",
             "", "⚠  PID must appear in the filename.",
             "The TRF calendar matches files to TRF line items by PID.",
             "A missing PID = plan does not appear on the leadership calendar.",
             "", "Claude shows the proposed name and waits for your confirmation.",
         ]},
        {"eyebrow": "STEP 4 — CALENDAR CHECK", "title": "📅  FY27 Calendar Compliance",
         "color": C_CAL, "bg": RGBColor(0x0C, 0x08, 0x28),
         "lines": [
             "Fetches the live MCC deadline calendar from GPMO SharePoint PDF.",
             "(Falls back to the dashboard FAQ table if the PDF is unavailable.)",
             "", "Checks: Required Completion Date ≤ MCC Submission Deadline.",
             "", "If date > deadline → STOP. Options: adjust date OR target next month.",
             "If date ≤ deadline → proceed, note the Send-By date in the output.",
             "If Dec–Jan shutdown window → flags Finance verification needed.",
         ]},
        {"eyebrow": "STEP 5 — CONFIRM", "title": "Pre-build Review",
         "color": C_PROC, "bg": C_PROC_BG,
         "lines": [
             "Claude shows a full summary of everything about to be built:",
             "  File name · Market · PID · Milestone · Revenue · Revenue type",
             "  Required completion date (MCC deadline) · Required/Expected quarter",
             "  Partner · Customer contact · Cisco team · Escalation path",
             "  Review windows · Cisco internal review · Risk status · Block note",
             "", "SLIPPING flag: shown if Expected Quarter ≠ Required Quarter.",
             "PM types yes to confirm, or specifies what to change.",
             "Nothing is built until explicit confirmation.",
         ]},
    ]
    gx = 0.3
    for g in gates:
        box(s, gx, 0.95, 3.1, 5.6, bg=g["bg"], border=g["color"])
        label(s, g["eyebrow"], gx+0.12, 0.98, 2.9, 0.2, size=7, bold=True, color=g["color"])
        label(s, g["title"], gx+0.12, 1.18, 2.9, 0.3, size=11, bold=True, color=g["color"])
        cy = 1.5
        for line in g["lines"]:
            label(s, line, gx+0.12, cy, 2.85, 0.24, size=8, color=C_LIGHT)
            cy += 0.24
        gx += 3.3
        if gx < 12:
            label(s, "→", gx - 0.22, 3.7, 0.22, 0.3, size=14, color=C_SUB, align=PP_ALIGN.CENTER)


# ── SLIDE 6: Build ────────────────────────────────────────────────────────────
def slide_build(prs):
    s = new_slide(prs)
    label(s, "Step 6–7 — Build the Activity Table", 0.4, 0.15, 12, 0.4,
          size=18, bold=True, color=C_WHITE)
    label(s, "N rows, calculated backwards from the milestone date using business days only.",
          0.4, 0.55, 12, 0.3, size=11, color=C_SUB)

    box(s, 0.3, 0.95, 4.6, 5.8, bg=C_PROC_BG, border=C_PROC)
    label(s, "DATE CALCULATION RULES", 0.5, 0.98, 4.3, 0.22, size=8, bold=True, color=C_PROC)
    label(s, "How dates are calculated", 0.5, 1.2, 4.3, 0.28, size=12, bold=True, color=C_WHITE)
    rules = [
        ("Business days only", "Mon–Fri. Never Saturday, Sunday,\nor US federal holidays."),
        ("Work backwards", "Start from the milestone date.\nCount back in business days for each row."),
        ("5-day rule", "MCC must be sent to partner/customer\nexactly 5 business days before milestone."),
        ("Bump rule", "End/completion dates → bump to previous\nbusiness day. Start dates → bump forward."),
        ("Cisco Drafts start", "Row 1 always starts today (or next biz day)."),
        ("Baseline is permanent", "Required Start and Required End are set\nonce and never changed — even if missed."),
        ("Actual dates", "Never set to a future date. Leave null\nuntil the step actually starts or ends."),
    ]
    cy = 1.55
    for title, desc in rules:
        label(s, f"▸  {title}", 0.5, cy, 4.3, 0.22, size=9, bold=True, color=C_ACCENT)
        cy += 0.22
        for line in desc.split('\n'):
            label(s, f"    {line}", 0.5, cy, 4.3, 0.22, size=8.5, color=C_LIGHT)
            cy += 0.22
        cy += 0.06

    box(s, 5.2, 0.95, 7.8, 5.8, bg=C_SURFACE, border=C_PROC)
    label(s, "ACTIVITY CHAIN — always displayed N → 1 (top to bottom)", 5.4, 0.98, 7.5, 0.22,
          size=8, bold=True, color=C_PROC)
    label(s, "Standard row sequence for every plan:", 5.4, 1.2, 7.5, 0.25,
          size=11, bold=True, color=C_WHITE)
    rows = [
        ("N",   True,  "MCC Submission in SDC",                      "On the milestone date"),
        ("N-1", True,  "MCC Approval by [Partner / Customer]",       "Within the 5-day window"),
        ("N-2", True,  "MCC Submission with 5 Day Rule",             "≥ 5 biz days before milestone"),
        ("N-3", False, "Customer Written [Deliverable] Approval",    "By 5-day rule date"),
        ("N-4", False, "Customer Reviews [Deliverable] Revision",    "Customer revision window (default 2 days)"),
        ("N-5", False, "Deliver [Deliverable] Revision to Customer", "1 business day"),
        ("N-6", False, "Cisco [Deliverable] Revision",               "2 business days"),
        ("N-7", False, "Customer Reviews [Deliverable] Iteration 1", "Customer initial window (default 5 days)"),
        ("N-8", False, "Deliver [Deliverable] Iteration 1 to Customer", "1 business day"),
        ("N-9", False, "Cisco Internal [Deliverable] Review",        "2 biz days (if PM said yes)"),
        ("1",   False, "Cisco Drafts Iteration 1 of [Deliverable]",  "Starts today → until internal review"),
    ]
    cy = 1.52
    for row_num, highlight, activity, timing in rows:
        if highlight:
            box(s, 5.25, cy - 0.02, 7.65, 0.28, bg=C_PROC_BG)
        tc = C_ACCENT if highlight else C_LIGHT
        label(s, row_num, 5.35, cy, 0.55, 0.24, size=9, bold=True, color=tc, align=PP_ALIGN.CENTER)
        label(s, activity, 5.95, cy, 4.3, 0.24, size=8.5, bold=highlight, color=tc)
        label(s, timing, 10.3, cy, 2.5, 0.24, size=8, color=C_SUB)
        cy += 0.29
    label(s, "[Deliverable] = document type per milestone: SDD, SRD, Implementation Plan, Test Plan, etc.",
          5.35, 6.05, 7.5, 0.35, size=8, italic=True, color=C_SUB)


# ── SLIDE 7: Output ───────────────────────────────────────────────────────────
def slide_output(prs):
    s = new_slide(prs)
    label(s, "Step 8 — Publish & Where Data Goes", 0.4, 0.15, 12, 0.4,
          size=18, bold=True, color=C_WHITE)
    label(s, "Claude does all three automatically after you confirm the plan.", 0.4, 0.55, 12, 0.3,
          size=11, color=C_SUB)
    cards_out = [
        {"n": "8a–8b", "title": "💾  Write JSON Locally", "color": C_OUT, "bg": C_OUT_BG,
         "lines": [
             "Saved to PLAN_DIR immediately — before the PR is opened.",
             "", "Path: PLAN_DIR/MARKET-pid-project.json",
             "  e.g. ~/reverse-plans/SLED_West-f51119-coh-lcs.json",
             "", "Schema includes:",
             "  market · pid · project · milestone · revenue",
             "  revenue_type · req_date · req_qtr · exp_qtr",
             "  risk · pm · block_note · last_updated",
             "  change_log [ { updated_at, pm } ]",
             "  tasks [ { num, status, activity, req_start, req_end,",
             "    actual_start, actual_end, dependency, risk,",
             "    owner, escalation } ]",
             "", "Tasks stored 1 → N. Dashboard reverses for display.",
             "", "To edit later: run /reverse-plan → 'editing an existing plan'.",
         ]},
        {"n": "8c", "title": "📤  Open GitHub Pull Request", "color": C_PROC, "bg": C_PROC_BG,
         "lines": [
             "Plans are never pushed directly to main.",
             "A branch + PR is created for maintainer review.",
             "", "Repo:    cx-usps-auto/Reverse-plans",
             "Branch:  reverse-plan/{pid}-{project-slug}",
             "Files:   data/{filename}.json",
             "         data/all.json  (rebuilt, N plans)",
             "", "PR body includes:",
             "  Market · PID · Project · Milestone · Revenue · PM",
             "", "No local clone required — uses GitHub API via gh CLI.",
             "", "After PR is open:",
             "  Maintainer reviews data/{filename}.json",
             "  Merges to main when ready",
         ]},
        {"n": "Dashboard", "title": "🟢  Live on the Dashboard", "color": C_OUT, "bg": C_OUT_BG,
         "lines": [
             "Plan card appears ~30 seconds after merge to main.",
             "", "Dashboard URL:",
             "  wwwin-github.cisco.com/pages/",
             "  cx-usps-auto/Reverse-plans/",
             "", "What the card shows:",
             "  PID · Project name · Milestone",
             "  Revenue · Revenue type · Risk status",
             "  Required quarter · PM · Last updated",
             "  View Plan → full table in new tab",
             "", "Update cadence:",
             "  EOD every Friday — mark completed steps",
             "  Complete with real Actual Start/End dates.",
             "", "SLIPPING: shown if Exp Qtr ≠ Req Qtr.",
         ]},
    ]
    cx = 0.3
    for c in cards_out:
        box(s, cx, 0.95, 4.1, 6.15, bg=c["bg"], border=c["color"])
        label(s, f"STEP {c['n']}", cx+0.15, 0.98, 3.8, 0.2, size=7, bold=True, color=c["color"])
        label(s, c["title"], cx+0.15, 1.18, 3.8, 0.28, size=12, bold=True, color=c["color"])
        cy = 1.5
        for line in c["lines"]:
            label(s, line, cx+0.15, cy, 3.8, 0.23, size=8.5, color=C_LIGHT)
            cy += 0.23
        cx += 4.4
        if cx < 9:
            label(s, "→", cx - 0.32, 3.95, 0.3, 0.35, size=14, color=C_SUB, align=PP_ALIGN.CENTER)


# ── SLIDE 8: Key rules ────────────────────────────────────────────────────────
def slide_rules(prs):
    s = new_slide(prs)
    label(s, "Key Rules — Quick Reference", 0.4, 0.15, 12, 0.4, size=18, bold=True, color=C_WHITE)
    label(s, "The most important rules to know when building or reviewing a reverse plan.",
          0.4, 0.55, 12, 0.3, size=11, color=C_SUB)
    rules_data = [
        ("Business Days Only",
         "All dates are Mon–Fri. No weekends, no US federal holidays.\nEnd/completion dates bump to previous business day.\nStart dates bump to next business day."),
        ("Required Dates Are Permanent",
         "Required Start and Required Completion are set once and never changed,\neven if the PM misses them. They are the fixed baseline.\nActual Start/End are only filled when work actually happens."),
        ("Table Order vs JSON Order",
         "The activity table always displays N → 1 (MCC Submission on top, Cisco Drafts on bottom).\nThe JSON file stores tasks 1 → N (Row 1 first).\nThe dashboard handles the reversal automatically."),
        ("5-Day Rule",
         "MCC must be sent to partner/customer exactly 5 business days before the milestone date.\nThis must fall on or before the month's MCC to Customer Send-By date from the GPMO calendar.\nRow N-2 in the table is always 'MCC Submission with 5 Day Rule'."),
        ("MCC Deadline Gate",
         "The milestone date must be on or before the MCC Submission Deadline for its month.\nIf it falls after the deadline, Claude stops and asks you to adjust the date\nor target the next month's window. Do not proceed until resolved."),
        ("SLIPPING Flag",
         "If Expected Completion Quarter ≠ Required Completion Quarter,\nleadership sees the milestone flagged as SLIPPING on the dashboard.\nThis is set in the pre-build review and stored in exp_qtr in the JSON."),
        ("Dependency Field — Never Blank",
         "Every activity table row must have a reasoned, non-blank Dependency.\n'None' fails validation. Use: 'Completion of [prior task]', 'Customer provides written approval',\n'Partner signs MCC', etc. Leadership reads this when a date is at risk."),
        ("Risk/Mitigation — Exit Criteria Required",
         "Every row must state the risk, the recovery action, and a clear exit criteria.\nA plan cannot be published without this.\nFor thin context rows: [PM to complete — state risk, recovery, and exit criteria]"),
        ("Cisco Fiscal Year (FY27)",
         "Q1: Aug · Sep · Oct     Q2: Nov · Dec · Jan\nQ3: Feb · Mar · Apr     Q4: May · Jun · Jul\nUse this when populating Required/Expected Completion Quarter fields."),
    ]
    cols = 3
    cw, ch = 4.1, 1.55
    mx, my = 0.35, 0.95
    pad = 0.18
    for i, (title, body) in enumerate(rules_data):
        rx = mx + (i % cols) * (cw + 0.2)
        ry = my + (i // cols) * (ch + 0.12)
        box(s, rx, ry, cw, ch, bg=C_SURFACE, border=C_PROC)
        label(s, title, rx+pad, ry+0.1, cw-pad*2, 0.26, size=9, bold=True, color=C_ACCENT)
        cy = ry + 0.36
        for line in body.split('\n'):
            label(s, line, rx+pad, cy, cw-pad*2, 0.24, size=8, color=C_LIGHT)
            cy += 0.24


# ── main ──────────────────────────────────────────────────────────────────────
def main():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    slide_title(prs)
    slide_overview(prs)
    slide_startup(prs)
    slide_paths(prs)
    slide_shared(prs)
    slide_build(prs)
    slide_output(prs)
    slide_rules(prs)

    out = os.path.join(os.path.expanduser("~"), "Desktop", "Reverse-Plan-Skill-v2.pptx")
    prs.save(out)
    print(f"Saved: {out}")

if __name__ == "__main__":
    main()
