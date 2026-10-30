"""
VVS Executive Update PowerPoint Generator
Based on: Oct 29, 2026 "Make a VVS Update" call + meeting notes from C16507
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.enum.dml import MSO_THEME_COLOR
import copy

# ── Cisco color palette ──────────────────────────────────────────────────────
CISCO_BLUE     = RGBColor(0x00, 0x5E, 0xBD)   # Primary Cisco Blue
CISCO_DARK     = RGBColor(0x1B, 0x2A, 0x4A)   # Deep navy
CISCO_TEAL     = RGBColor(0x00, 0xBC, 0xEB)   # Cisco accent teal
WHITE          = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY     = RGBColor(0xF2, 0xF4, 0xF6)
MED_GRAY       = RGBColor(0xD0, 0xD5, 0xDB)
TEXT_DARK      = RGBColor(0x1A, 0x1A, 0x2E)
GREEN_OK       = RGBColor(0x00, 0x87, 0x4A)   # Resolved
AMBER          = RGBColor(0xF0, 0x8C, 0x00)   # In progress
RED            = RGBColor(0xC0, 0x39, 0x2B)   # Critical
LIGHT_BLUE     = RGBColor(0xE8, 0xF4, 0xFF)   # row fill
HEADER_BLUE    = RGBColor(0x1B, 0x47, 0x8A)   # table header

# ── Slide dimensions (Widescreen 13.33 x 7.5) ───────────────────────────────
W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

BLANK = prs.slide_layouts[6]  # totally blank


# ── helpers ──────────────────────────────────────────────────────────────────
def add_rect(slide, x, y, w, h, fill_rgb, alpha=None):
    shape = slide.shapes.add_shape(1, x, y, w, h)  # MSO_SHAPE_TYPE.RECTANGLE
    shape.line.fill.background()
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    return shape


def add_text_box(slide, text, x, y, w, h, font_size=14, bold=False,
                 color=TEXT_DARK, align=PP_ALIGN.LEFT, wrap=True, italic=False):
    txb = slide.shapes.add_textbox(x, y, w, h)
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txb


def add_header_bar(slide, title, subtitle=None):
    """Full-width top bar with title."""
    add_rect(slide, 0, 0, W, Inches(1.1), CISCO_DARK)
    add_text_box(slide, title, Inches(0.35), Inches(0.1), Inches(12.5), Inches(0.65),
                 font_size=24, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    if subtitle:
        add_text_box(slide, subtitle, Inches(0.35), Inches(0.72), Inches(12.5), Inches(0.32),
                     font_size=12, bold=False, color=CISCO_TEAL, align=PP_ALIGN.LEFT)
    # bottom accent line
    add_rect(slide, 0, Inches(1.1), W, Inches(0.04), CISCO_TEAL)


def add_footer(slide, page_num):
    add_rect(slide, 0, Inches(7.2), W, Inches(0.3), CISCO_DARK)
    add_text_box(slide, "CISCO CONFIDENTIAL  |  VVS Executive Update  |  October 29, 2026",
                 Inches(0.3), Inches(7.22), Inches(10), Inches(0.25),
                 font_size=8, color=MED_GRAY)
    add_text_box(slide, str(page_num), Inches(12.8), Inches(7.22), Inches(0.3), Inches(0.25),
                 font_size=8, color=MED_GRAY, align=PP_ALIGN.RIGHT)


def bullet_box(slide, items, x, y, w, h, font_size=13, title=None, title_color=CISCO_DARK,
               bullet_color=TEXT_DARK, bg_color=None, indent=0.18):
    if bg_color:
        add_rect(slide, x, y, w, h, bg_color)
    if title:
        add_text_box(slide, title, x + Inches(0.12), y + Inches(0.05), w - Inches(0.2),
                     Inches(0.35), font_size=font_size + 1, bold=True, color=title_color)
        y_off = y + Inches(0.42)
    else:
        y_off = y + Inches(0.1)
    item_h = (h - (Inches(0.42) if title else Inches(0.1)) - Inches(0.05)) / max(len(items), 1)
    for item in items:
        add_text_box(slide, f"▸  {item}", x + Inches(indent), y_off, w - Inches(0.3),
                     item_h + Inches(0.1), font_size=font_size, color=bullet_color)
        y_off += item_h


def status_badge(slide, label, x, y, w=Inches(1.5), h=Inches(0.35), color=GREEN_OK):
    add_rect(slide, x, y, w, h, color)
    add_text_box(slide, label, x, y, w, h, font_size=9, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 1 — TITLE
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)

# Full navy background
add_rect(slide, 0, 0, W, H, CISCO_DARK)

# Teal accent stripe
add_rect(slide, 0, Inches(4.4), W, Inches(0.08), CISCO_TEAL)

# Large title
add_text_box(slide, "Virtual Visitation System (VVS)",
             Inches(0.7), Inches(1.4), Inches(11.9), Inches(1.1),
             font_size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

add_text_box(slide, "Executive Status Update",
             Inches(0.7), Inches(2.55), Inches(11.9), Inches(0.9),
             font_size=28, bold=False, color=CISCO_TEAL, align=PP_ALIGN.CENTER)

add_text_box(slide, "October 29, 2026",
             Inches(0.7), Inches(3.55), Inches(11.9), Inches(0.55),
             font_size=18, bold=False, color=MED_GRAY, align=PP_ALIGN.CENTER)

# Account list
add_text_box(slide, "Key Accounts: MNIT / MN DOC  ·  CDCR  ·  Ohio",
             Inches(0.7), Inches(4.7), Inches(11.9), Inches(0.45),
             font_size=14, bold=False, color=CISCO_TEAL, align=PP_ALIGN.CENTER)

add_text_box(slide, "Prepared by: Houston Downs, Customer Program Manager",
             Inches(0.7), Inches(5.5), Inches(11.9), Inches(0.38),
             font_size=11, color=MED_GRAY, align=PP_ALIGN.CENTER)

add_text_box(slide, "CISCO CONFIDENTIAL",
             Inches(0.7), Inches(6.9), Inches(11.9), Inches(0.3),
             font_size=9, color=MED_GRAY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 2 — EXECUTIVE SUMMARY (1-page brief)
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, LIGHT_GRAY)
add_header_bar(slide, "Executive Summary", "VVS Platform — Past Issues, Current Standing & Path Forward")
add_footer(slide, 2)

# Three-column boxes
col_w = Inches(3.9)
col_h = Inches(5.4)
col_y = Inches(1.3)
gap   = Inches(0.26)

# Col 1 — Past Issues
add_rect(slide, Inches(0.25), col_y, col_w, col_h, WHITE)
add_rect(slide, Inches(0.25), col_y, col_w, Inches(0.45), RED)
add_text_box(slide, "PAST ISSUES", Inches(0.3), col_y + Inches(0.05), col_w, Inches(0.38),
             font_size=12, bold=True, color=WHITE)
issues = [
    "Sept 25: Mass device cancellation — all MN DOC virtual court hearings wiped out (TAC 701470746)",
    "Cisco lost core dev & test team — no engineering capacity for ~5 months",
    "~15 open bugs across accounts; customers resorting to manual workarounds",
    "CI/CD pipeline broken — no automated test validation before releases",
    "Test environment out of sync with production — unreliable hotfix verification",
    "Customer sentiment critically low — Sara Larson (MNIT): 'I'm not seeing any progress'",
    "Education dashboard blank-row issue (MN DOC) — no RCA provided",
    "Webex Control Hub device naming conflict causing periodic device drops",
]
bullet_box(slide, issues, Inches(0.25), col_y + Inches(0.48),
           col_w, col_h - Inches(0.5), font_size=10.5, bullet_color=TEXT_DARK)

# Col 2 — Current Standing
x2 = Inches(0.25) + col_w + gap
add_rect(slide, x2, col_y, col_w, col_h, WHITE)
add_rect(slide, x2, col_y, col_w, Inches(0.45), AMBER)
add_text_box(slide, "CURRENT STANDING", x2 + Inches(0.05), col_y + Inches(0.05), col_w, Inches(0.38),
             font_size=12, bold=True, color=WHITE)
standing = [
    "New dev resources onboarded ~Oct 21 — Ricky (engineer) dedicated 100% to VVS bug fixes",
    "Backend DevOps infrastructure rebuilt & stabilized — pipelines now validate all pushes",
    "Test environment mirrors production across multiple versions — reliable QA cycle restored",
    "Platform stable — no new incidents since Sept 25 mass cancellation",
    "MFA successfully deployed for 53 Control Hub accounts (Oct 26) — no disruptions",
    "CDCR: Quick wins completed; coordinated release ready to deploy",
    "MNIT / MN DOC: Dashboard issue actively being worked today",
    "Ohio: Queued — targeted for early next week",
    "Customer sentiment: low but stable; all three accounts re-engaged",
]
bullet_box(slide, standing, x2, col_y + Inches(0.48),
           col_w, col_h - Inches(0.5), font_size=10.5, bullet_color=TEXT_DARK)

# Col 3 — Next Steps
x3 = x2 + col_w + gap
add_rect(slide, x3, col_y, col_w, col_h, WHITE)
add_rect(slide, x3, col_y, col_w, Inches(0.45), CISCO_BLUE)
add_text_box(slide, "NEXT STEPS", x3 + Inches(0.05), col_y + Inches(0.05), col_w, Inches(0.38),
             font_size=12, bold=True, color=WHITE)
nextsteps = [
    "Ricky continues MN DOC dashboard fix this week (Oct 29–31)",
    "CDCR: Deploy coordinated bug fix release — this week",
    "Ohio: Begin targeted bug fixes early next week",
    "Muba returns Monday — takes over deployment approvals; frees Ricky for fixes",
    "Phased batch releases with user-guide-style release notes per account",
    "Ryan: 'All three big customers should see improvement by next week'",
    "Establish every-other-day status update cadence with customers",
    "Formalize version naming convention & bug tracking spreadsheet for transparency",
    "Monday/Tuesday maintenance windows (preferred — avoids weekend visitation)",
]
bullet_box(slide, nextsteps, x3, col_y + Inches(0.48),
           col_w, col_h - Inches(0.5), font_size=10.5, bullet_color=TEXT_DARK)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 3 — CRITICAL INCIDENT & HISTORICAL TIMELINE
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, LIGHT_GRAY)
add_header_bar(slide, "Historical Context & Critical Incident", "How We Got Here — Root Cause & Impact")
add_footer(slide, 3)

# Timeline spine
spine_y = Inches(4.0)
add_rect(slide, Inches(0.5), spine_y, Inches(12.33), Inches(0.06), MED_GRAY)

# Timeline events
events = [
    ("May–Sep\n2026",      "~5 months",     "Dev team lost",          "Engineering capacity gap;\ncustomer issues unaddressed",     RED),
    ("Sep 25\n2026",       "Critical",      "Mass Cancellation",      "All MN DOC virtual court\nhearings canceled (TAC 701470746)", RED),
    ("Sep 28\n2026",       "Response",      "Executive Write-Up",     "Cisco PM issued formal exec\nincident summary to MNIT",        AMBER),
    ("Oct 19–21\n2026",    "Recovery",      "Dev Team Rebuilt",       "3 new engineers onboarded;\nbackend work begins",             AMBER),
    ("Oct 21–26\n2026",    "Progress",      "Infrastructure Fixed",   "CI/CD pipelines stabilized;\ntest environment restored",      GREEN_OK),
    ("Oct 26\n2026",       "Win",           "MFA Deployed Clean",     "53 Control Hub accounts;\nno incidents",                     GREEN_OK),
    ("Oct 29\n2026",       "Active",        "Bug Fix Sprint",         "CDCR done; MN active now;\nOhio next week",                   CISCO_BLUE),
]

x_positions = [Inches(0.55), Inches(2.35), Inches(4.15), Inches(5.95),
               Inches(7.75), Inches(9.55), Inches(11.35)]

for i, (date, badge, title, desc, color) in enumerate(events):
    xp = x_positions[i]
    # dot on spine
    add_rect(slide, xp + Inches(0.5), spine_y - Inches(0.09), Inches(0.18), Inches(0.24), color)
    # card above the spine
    card_h = Inches(1.75)
    card_y = spine_y - card_h - Inches(0.12)
    add_rect(slide, xp, card_y, Inches(1.58), card_h, WHITE)
    add_rect(slide, xp, card_y, Inches(1.58), Inches(0.28), color)
    add_text_box(slide, badge, xp + Inches(0.04), card_y + Inches(0.03), Inches(1.5), Inches(0.23),
                 font_size=8, bold=True, color=WHITE)
    add_text_box(slide, title, xp + Inches(0.06), card_y + Inches(0.32), Inches(1.46), Inches(0.38),
                 font_size=9.5, bold=True, color=TEXT_DARK)
    add_text_box(slide, desc, xp + Inches(0.06), card_y + Inches(0.72), Inches(1.46), Inches(0.95),
                 font_size=8.5, color=TEXT_DARK)
    # date below spine
    add_text_box(slide, date, xp + Inches(0.05), spine_y + Inches(0.1), Inches(1.58), Inches(0.55),
                 font_size=8.5, bold=True, color=CISCO_DARK, align=PP_ALIGN.CENTER)

# Callout box for the Sept 25 incident details
add_rect(slide, Inches(0.3), Inches(4.8), Inches(12.73), Inches(1.95), WHITE)
add_rect(slide, Inches(0.3), Inches(4.8), Inches(0.18), Inches(1.95), RED)
add_text_box(slide, "Sept 25 Incident — Scope of Impact",
             Inches(0.6), Inches(4.85), Inches(12.2), Inches(0.3),
             font_size=11, bold=True, color=RED)
impact_items = [
    "Virtual Court: All active & upcoming hearings auto-canceled when devices dropped from location mappings.",
    "Education & Non-Contact Visitation: Collateral impact — MNIT staff forced to manually intervene and reschedule all sessions.",
    "Device Behavior: Endpoints vanished from VVS location maps; re-add attempts blocked until backend sync was manually restored by Cisco engineering.",
    "Response: Cisco TAC and Engineering engaged same day despite internal holiday. Root cause linked to Webex Control Hub sync/naming conflict (TAC 701334792).",
]
for i, item in enumerate(impact_items):
    add_text_box(slide, f"▸  {item}", Inches(0.62), Inches(5.2) + i * Inches(0.36),
                 Inches(12.2), Inches(0.34), font_size=9.5, color=TEXT_DARK)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 4 — INTERNAL ISSUES RESOLUTION ROADMAP (the diagram)
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, LIGHT_GRAY)
add_header_bar(slide, "Internal Recovery Roadmap",
               "How Each Root-Cause Issue Is Being Addressed — Problem → Action → Resolution Status")
add_footer(slide, 4)

# Columns headers
col_labels = ["ROOT CAUSE ISSUE", "ACTION TAKEN", "OWNER(S)", "STATUS"]
col_xs     = [Inches(0.2), Inches(3.95), Inches(8.55), Inches(10.6)]
col_ws     = [Inches(3.68), Inches(4.52), Inches(1.98), Inches(2.5)]
header_y   = Inches(1.22)
header_h   = Inches(0.36)

for i, label in enumerate(col_labels):
    add_rect(slide, col_xs[i], header_y, col_ws[i] - Inches(0.04), header_h, HEADER_BLUE)
    add_text_box(slide, label, col_xs[i] + Inches(0.06), header_y + Inches(0.04),
                 col_ws[i] - Inches(0.1), header_h - Inches(0.06),
                 font_size=9.5, bold=True, color=WHITE)

rows = [
    (
        "1. Development Team Staffing Gap\n(No engineering capacity, ~5 months)",
        "Onboarded Ricky (dedicated engineer) + additional resources ~Oct 21; removed from all other assignments to focus solely on VVS",
        "Ryan Schaefer\nRicky",
        ("RESOLVED", GREEN_OK),
    ),
    (
        "2. Broken CI/CD Pipeline\n(No automated validation of code before pushing to production)",
        "Rebuilt DevOps infrastructure with end-to-end pipeline validation; all pushes now verified against automated tests before release",
        "Ryan Schaefer\nDev Team",
        ("RESOLVED", GREEN_OK),
    ),
    (
        "3. Test Environment Out of Sync\n(Could not safely test fixes; environment didn't mirror production)",
        "Test environment rebuilt to mirror production across multiple version configurations; QA cycles restored",
        "Ryan Schaefer\nRicky",
        ("RESOLVED", GREEN_OK),
    ),
    (
        "4. Bug Backlog (~15 open issues)\n(Customers running workarounds; no fix timeline)",
        "Triage by account priority: CDCR quick wins done; MN DOC dashboard fix active today; Ohio targeted early next week",
        "Ricky\nMuba (deploys)",
        ("IN PROGRESS", AMBER),
    ),
    (
        "5. Deployment Bottleneck\n(Single engineer doing both fixes & deployments)",
        "Muba returns Monday — takes over all deployment approvals. Ricky stays focused on code fixes. Roles split.",
        "Muba Khan\nRicky",
        ("IN PROGRESS", AMBER),
    ),
    (
        "6. Customer Transparency Gap\n(Customers had no visibility into fix status or timelines)",
        "Every-other-day status updates to customers; user-guide-style release notes per fix; version naming convention; bug-tracking spreadsheet",
        "Houston Downs\nPraveen Matanam",
        ("IN PROGRESS", AMBER),
    ),
    (
        "7. Webex Control Hub MFA Risk\n(MFA enablement could disrupt VBS service accounts)",
        "Validated MFA implementation approach; deployed for 53 Control Hub accounts Oct 26 — clean, zero incidents. Advisory documented.",
        "Mike Reinhart\nPraveen Matanam",
        ("RESOLVED", GREEN_OK),
    ),
]

row_y = header_y + header_h + Inches(0.05)
row_h = Inches(0.66)

for idx, (issue, action, owners, (status_lbl, status_color)) in enumerate(rows):
    bg = WHITE if idx % 2 == 0 else LIGHT_BLUE
    for i, (cx, cw) in enumerate(zip(col_xs, col_ws)):
        add_rect(slide, cx, row_y, cw - Inches(0.04), row_h, bg)

    # Issue
    add_text_box(slide, issue, col_xs[0] + Inches(0.07), row_y + Inches(0.04),
                 col_ws[0] - Inches(0.12), row_h - Inches(0.06), font_size=9, color=TEXT_DARK)
    # Action
    add_text_box(slide, action, col_xs[1] + Inches(0.07), row_y + Inches(0.04),
                 col_ws[1] - Inches(0.12), row_h - Inches(0.06), font_size=9, color=TEXT_DARK)
    # Owners
    add_text_box(slide, owners, col_xs[2] + Inches(0.07), row_y + Inches(0.04),
                 col_ws[2] - Inches(0.12), row_h - Inches(0.06), font_size=9, bold=True, color=CISCO_DARK)
    # Status badge
    bx = col_xs[3] + Inches(0.2)
    by = row_y + (row_h - Inches(0.28)) / 2
    add_rect(slide, bx, by, Inches(2.0), Inches(0.28), status_color)
    add_text_box(slide, status_lbl, bx, by, Inches(2.0), Inches(0.28),
                 font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    row_y += row_h + Inches(0.03)

# Legend
add_rect(slide, Inches(0.2), Inches(7.0), Inches(2.5), Inches(0.22), GREEN_OK)
add_text_box(slide, "RESOLVED", Inches(0.2), Inches(7.0), Inches(2.5), Inches(0.22),
             font_size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_rect(slide, Inches(2.82), Inches(7.0), Inches(2.5), Inches(0.22), AMBER)
add_text_box(slide, "IN PROGRESS", Inches(2.82), Inches(7.0), Inches(2.5), Inches(0.22),
             font_size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 5 — ACCOUNT STATUS: MNIT / MN DOC
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, LIGHT_GRAY)
add_header_bar(slide, "Account Status: MNIT / MN DOC  (C16507)",
               "Minnesota Department of Corrections — Virtual Visitation System")
add_footer(slide, 5)

# Sentiment meter (left panel)
add_rect(slide, Inches(0.2), Inches(1.3), Inches(3.5), Inches(5.5), WHITE)
add_rect(slide, Inches(0.2), Inches(1.3), Inches(3.5), Inches(0.38), HEADER_BLUE)
add_text_box(slide, "CUSTOMER SENTIMENT", Inches(0.28), Inches(1.33), Inches(3.4), Inches(0.3),
             font_size=10, bold=True, color=WHITE)

# Sentiment thermometer visual
levels = [
    ("Critical", RED,      False),
    ("Very Low",  RED,      False),
    ("Low",       AMBER,    True),   # current
    ("Stable",    AMBER,    False),
    ("Improving", GREEN_OK, False),
    ("Strong",    GREEN_OK, False),
]
bar_y = Inches(1.8)
bar_h = Inches(0.6)
for lbl, color, current in reversed(levels):
    fill = color if current else MED_GRAY
    add_rect(slide, Inches(0.6), bar_y, Inches(1.8), bar_h - Inches(0.04), fill)
    marker = "◀  CURRENT" if current else ""
    add_text_box(slide, lbl, Inches(0.64), bar_y + Inches(0.16), Inches(1.72), Inches(0.28),
                 font_size=9, bold=current, color=WHITE if current else TEXT_DARK)
    if current:
        add_text_box(slide, marker, Inches(2.5), bar_y + Inches(0.16), Inches(1.0), Inches(0.28),
                     font_size=9, bold=True, color=AMBER)
    bar_y += bar_h

add_text_box(slide, "Trend: ↑ Stabilizing", Inches(0.3), Inches(5.55), Inches(3.3), Inches(0.32),
             font_size=10, bold=True, color=GREEN_OK)
add_text_box(slide, '"Low but stable — we may have just hit rock bottom."\n— Houston Downs, PM (Oct 29, 2026)',
             Inches(0.3), Inches(5.9), Inches(3.3), Inches(0.65),
             font_size=8.5, italic=True, color=TEXT_DARK)

# Right panel — status details
right_x = Inches(3.9)
right_w = Inches(9.2)

# Key Issues / Active Cases
add_rect(slide, right_x, Inches(1.3), right_w, Inches(2.6), WHITE)
add_rect(slide, right_x, Inches(1.3), right_w, Inches(0.38), HEADER_BLUE)
add_text_box(slide, "Active Issues & Open Cases", right_x + Inches(0.1), Inches(1.33),
             right_w - Inches(0.15), Inches(0.3), font_size=10, bold=True, color=WHITE)
issues_mn = [
    "Dashboard blank-row issue (VVS-4499) — Ricky assigned TODAY; multiple related sub-cases being parsed",
    "Jenkins CI/CD intermittent failures — under investigation by Ryan's team; not customer-impacting",
    "Education dashboard records inconsistency — pending RCA; user guide being updated as indirect visibility tool",
    "TAC 701470746 (Mass Cancellation) — Sev 3, open; root cause traced, monitoring for recurrence",
    "Kernel update submitted (Praveen) — assessed as likely sufficient; pending team review",
]
for i, item in enumerate(issues_mn):
    add_text_box(slide, f"▸  {item}", right_x + Inches(0.15), Inches(1.75) + i * Inches(0.37),
                 right_w - Inches(0.25), Inches(0.35), font_size=9.5, color=TEXT_DARK)

# Wins / Positive Signals
add_rect(slide, right_x, Inches(4.0), right_w, Inches(2.0), WHITE)
add_rect(slide, right_x, Inches(4.0), right_w, Inches(0.38), GREEN_OK)
add_text_box(slide, "Wins & Positive Signals", right_x + Inches(0.1), Inches(4.03),
             right_w - Inches(0.15), Inches(0.3), font_size=10, bold=True, color=WHITE)
wins_mn = [
    "Platform STABLE — no new production incidents since Sept 25 mass cancellation (34 days)",
    "MFA for Control Hub deployed Oct 26 — zero disruptions to VVS service accounts (confirmed by Bobbi & Caroline)",
    "Bobbi Dumphy (DOC) & Caroline Koerner (MNIT) both confirmed 'no new issues' following Control Hub upgrade",
    "Dev resources are onboarded and actively working — 'net positive' signal from Mike Matthews",
    "Houston: 'Leaps and bounds from where we were a month ago' — first time PM feels optimistic",
]
for i, item in enumerate(wins_mn):
    add_text_box(slide, f"✓  {item}", right_x + Inches(0.15), Inches(4.45) + i * Inches(0.3),
                 right_w - Inches(0.25), Inches(0.28), font_size=9.5, color=GREEN_OK)

# Key Stakeholders
add_rect(slide, right_x, Inches(6.1), right_w, Inches(0.8), LIGHT_BLUE)
add_text_box(slide, "Key Contacts:  Bobbi Dumphy (DOC — Virtual Court Mgr)  ·  Caroline Koerner (MNIT — Systems)  ·  Sara Larson (MNIT)  ·  Mike Reinhart (MNIT Enterprise)  ·  Vaughn Scheunemann (MNIT)",
             right_x + Inches(0.1), Inches(6.15), right_w - Inches(0.15), Inches(0.65),
             font_size=9, color=TEXT_DARK)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 6 — ACCOUNT STATUS: CDCR & OHIO
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, LIGHT_GRAY)
add_header_bar(slide, "Account Status: CDCR  &  Ohio",
               "California Department of Corrections & Rehabilitation  |  Ohio Corrections")
add_footer(slide, 6)

# ── CDCR (left half) ──────────────────────────────────────────────────────────
cx = Inches(0.2)
cw = Inches(6.3)

add_rect(slide, cx, Inches(1.3), cw, Inches(5.7), WHITE)
add_rect(slide, cx, Inches(1.3), cw, Inches(0.45), CISCO_BLUE)
add_text_box(slide, "CDCR — California Dept of Corrections & Rehabilitation",
             cx + Inches(0.1), Inches(1.33), cw - Inches(0.15), Inches(0.38),
             font_size=11, bold=True, color=WHITE)

status_badge(slide, "QUICK WINS DONE", cx + Inches(0.15), Inches(1.85), Inches(2.0), Inches(0.3), GREEN_OK)
status_badge(slide, "RELEASE PENDING", cx + Inches(2.25), Inches(1.85), Inches(2.0), Inches(0.3), AMBER)

cdcr_items = [
    ("Priority Focus", "CDCR was the FIRST account addressed in the bug fix sprint — quick wins have been completed per Ryan Schaefer."),
    ("Fix Status", "Engineering has resolved the high-priority CDCR issues. Fix package is ready and has not yet been deployed — release coordination in progress."),
    ("Customer Engagement", "CDCR is actively replying to PM emails — strong re-engagement signal. Houston: 'They seem to be pretty active replying back.'"),
    ("Release Timeline", "Deployment to be coordinated this week. Priority issue: release must not break other account environments — batched release strategy."),
    ("Sentiment", "Customer was disengaged for several months. Re-engagement is a positive signal. Fix delivery will be the critical moment to rebuild trust."),
    ("Next Step", "Houston and Ryan to coordinate CDCR release schedule. Jesse (exec) is on all related email threads."),
]
for i, (label, text) in enumerate(cdcr_items):
    y = Inches(2.28) + i * Inches(0.7)
    add_text_box(slide, label + ":", cx + Inches(0.15), y, Inches(1.3), Inches(0.28),
                 font_size=9.5, bold=True, color=CISCO_DARK)
    add_text_box(slide, text, cx + Inches(1.5), y, cw - Inches(1.65), Inches(0.62),
                 font_size=9.5, color=TEXT_DARK)

# ── Ohio (right half) ─────────────────────────────────────────────────────────
ox = Inches(6.82)
ow = Inches(6.3)

add_rect(slide, ox, Inches(1.3), ow, Inches(5.7), WHITE)
add_rect(slide, ox, Inches(1.3), ow, Inches(0.45), CISCO_DARK)
add_text_box(slide, "Ohio — Corrections Virtual Visitation",
             ox + Inches(0.1), Inches(1.33), ow - Inches(0.15), Inches(0.38),
             font_size=11, bold=True, color=WHITE)

status_badge(slide, "QUEUED — NEXT WEEK", ox + Inches(0.15), Inches(1.85), Inches(2.4), Inches(0.3), CISCO_TEAL)

ohio_items = [
    ("Priority Order", "Ohio is third in the sequenced bug fix sprint: CDCR → MNIT/MN DOC → Ohio. Ryan Schaefer confirmed Ohio will be addressed early next week."),
    ("Scope", "Ohio has 'housing AH' items and other issues that are part of the 15-item bug backlog. Ryan noted these are being set aside temporarily while critical accounts are cleared."),
    ("Approach", "Same triage model: identify critical/quick-win issues, batch-fix with Ricky, coordinate release deployment with Muba (back Monday)."),
    ("Expectation", "Ryan: 'Hopefully all three of those big customers will have an increase by next week.' Ohio is included in this target."),
    ("Stakeholder Note", "Gary Depretta (exec) is named in Jesse's email threads — indicating Ohio has exec-level visibility. This is a high-visibility account."),
    ("Next Step", "Ryan's team targets Ohio beginning early week of Nov 3. Houston to notify Ohio stakeholders proactively."),
]
for i, (label, text) in enumerate(ohio_items):
    y = Inches(2.28) + i * Inches(0.7)
    add_text_box(slide, label + ":", ox + Inches(0.15), y, Inches(1.45), Inches(0.28),
                 font_size=9.5, bold=True, color=CISCO_DARK)
    add_text_box(slide, text, ox + Inches(1.65), y, ow - Inches(1.8), Inches(0.62),
                 font_size=9.5, color=TEXT_DARK)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 7 — KUDOS, WINS & PATH FORWARD
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, LIGHT_GRAY)
add_header_bar(slide, "Kudos, Wins & Path Forward",
               "Recognizing the team's efforts — and what comes next")
add_footer(slide, 7)

# ── Shout-outs grid ───────────────────────────────────────────────────────────
shoutouts = [
    ("Ryan Schaefer", "Engineering Leadership", "Rebuilt the DevOps pipeline and test environment from scratch while managing a one-engineer dev team. Ryan personally championed stabilization and is architecting the phased fix strategy. His commitment is what's turning this around."),
    ("Ricky\n(VVS Dev Engineer)", "Bug Fix Execution", "Onboarded in under 2 weeks and is already executing batch bug fixes against one of the platform's most complex accounts (MN DOC). Ricky is the engine behind the current sprint — dedicated 100% to VVS."),
    ("Praveen Matanam", "Documentation & Operations", "Authored the Control Hub/MFA advisory that enabled the clean MFA deployment. Joined a critical call on India's national holiday. Consistently bridges engineering with customer communication."),
    ("Mike Matthews", "Technical Leadership", "Secured authorization for new dev resources, coordinated manual sync routines during the Sept 25 incident, and manages support-side coordination across all accounts."),
    ("Bobbi Dumphy &\nCaroline Koerner", "Customer Partnership", "MN DOC and MNIT's primary points of contact have remained constructively engaged through months of instability — enabling daily stand-ups that keep issues surfaced and progress visible."),
    ("Houston Downs", "Program Management", "Maintained daily customer cadence through a difficult period, escalated transparently, and is now the connective tissue between customers, delivery, and engineering leadership."),
]

card_w = Inches(4.1)
card_h = Inches(2.3)
gap = Inches(0.18)
start_x = Inches(0.2)
start_y = Inches(1.3)

for i, (name, role, desc) in enumerate(shoutouts):
    row = i // 3
    col = i % 3
    cx = start_x + col * (card_w + gap)
    cy = start_y + row * (card_h + gap)
    add_rect(slide, cx, cy, card_w, card_h, WHITE)
    add_rect(slide, cx, cy, card_w, Inches(0.38), CISCO_BLUE)
    add_text_box(slide, name, cx + Inches(0.1), cy + Inches(0.04), card_w - Inches(0.15),
                 Inches(0.32), font_size=10.5, bold=True, color=WHITE)
    add_text_box(slide, role, cx + Inches(0.1), cy + Inches(0.4), card_w - Inches(0.15),
                 Inches(0.24), font_size=9, bold=False, color=CISCO_TEAL, italic=True)
    add_text_box(slide, desc, cx + Inches(0.1), cy + Inches(0.68), card_w - Inches(0.18),
                 card_h - Inches(0.72), font_size=9, color=TEXT_DARK)


# ── Bottom wins bar ───────────────────────────────────────────────────────────
add_rect(slide, Inches(0.2), Inches(6.98), Inches(12.93), Inches(0.38), GREEN_OK)
wins = "✓ Platform stable 34+ days  ·  ✓ MFA deployed clean  ·  ✓ CDCR quick wins done  ·  ✓ Dev pipeline rebuilt  ·  ✓ Test env restored  ·  ✓ Phased release strategy in place"
add_text_box(slide, wins, Inches(0.35), Inches(7.0), Inches(12.7), Inches(0.32),
             font_size=9.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE 8 — NEXT STEPS & CLOSING
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, LIGHT_GRAY)
add_header_bar(slide, "Next Steps & Closing Notes", "Immediate priorities and commitments through November 7, 2026")
add_footer(slide, 8)

# Two-column layout
col_w2 = Inches(6.1)
col_gap = Inches(0.3)
col_y2  = Inches(1.35)
col_h2  = Inches(5.5)

# Left: Immediate actions (this week)
add_rect(slide, Inches(0.2), col_y2, col_w2, col_h2, WHITE)
add_rect(slide, Inches(0.2), col_y2, col_w2, Inches(0.42), CISCO_DARK)
add_text_box(slide, "THIS WEEK (Oct 29 – Nov 1)", Inches(0.3), col_y2 + Inches(0.06),
             col_w2 - Inches(0.15), Inches(0.32), font_size=11, bold=True, color=WHITE)

this_week = [
    ("Houston", "Share Ryan's update with MN DOC at Friday (Oct 31) morning stand-up"),
    ("Ricky", "MN DOC dashboard issue — active fix in progress (Oct 29)"),
    ("Houston / Ryan", "Coordinate and execute CDCR release deployment"),
    ("Muba (returns Mon)", "Assume deployment approval responsibilities — free Ricky for code fixes"),
    ("Mike Matthews", "Investigate Jenkins CI/CD failure root cause with Ryan's team"),
    ("Houston", "Review updated user guide for applied fixes (education dashboard records)"),
    ("Houston", "Re-raise bug visibility request with resource back in office"),
    ("Praveen", "Finalize bug tracking spreadsheet and version naming convention"),
]
for i, (owner, action) in enumerate(this_week):
    y = col_y2 + Inches(0.5) + i * Inches(0.58)
    add_rect(slide, Inches(0.28), y, Inches(1.2), Inches(0.28), CISCO_BLUE)
    add_text_box(slide, owner, Inches(0.28), y, Inches(1.2), Inches(0.28),
                 font_size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text_box(slide, action, Inches(1.55), y, col_w2 - Inches(1.45), Inches(0.4),
                 font_size=9.5, color=TEXT_DARK)

# Right: Next week + forward
rx = Inches(0.2) + col_w2 + col_gap
add_rect(slide, rx, col_y2, col_w2, col_h2, WHITE)
add_rect(slide, rx, col_y2, col_w2, Inches(0.42), HEADER_BLUE)
add_text_box(slide, "NEXT WEEK & BEYOND (Nov 3+)", rx + Inches(0.1), col_y2 + Inches(0.06),
             col_w2 - Inches(0.15), Inches(0.32), font_size=11, bold=True, color=WHITE)

next_week = [
    ("Ryan / Ricky", "Begin Ohio bug fix sprint — early next week (target: before Nov 7)"),
    ("Houston", "Notify Ohio stakeholders proactively with fix timeline"),
    ("Ryan's team", "Reconnect with Muba on reprioritized issue list across all accounts"),
    ("Houston", "Establish every-other-day status update cadence with all accounts"),
    ("All", "Phased bug fix releases with user-guide release notes — all three accounts"),
    ("Houston", "Ongoing maintenance of Monday/Tuesday-only change windows (no Fri–Sun)"),
    ("Jesse (Exec)", "Internal exec briefing — Gary Depretta (Ohio) meeting for VVS future direction"),
    ("Houston/Ryan", "Monitor for any new incidents; Jenkins CI/CD watch ongoing"),
]
for i, (owner, action) in enumerate(next_week):
    y = col_y2 + Inches(0.5) + i * Inches(0.58)
    add_rect(slide, rx + Inches(0.08), y, Inches(1.3), Inches(0.28), HEADER_BLUE)
    add_text_box(slide, owner, rx + Inches(0.08), y, Inches(1.3), Inches(0.28),
                 font_size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text_box(slide, action, rx + Inches(1.5), y, col_w2 - Inches(1.45), Inches(0.4),
                 font_size=9.5, color=TEXT_DARK)


# ── Save ──────────────────────────────────────────────────────────────────────
out_path = os.path.join(os.path.expanduser("~"), "Desktop", "VVS_Executive_Update_2026-10-29.pptx")
prs.save(out_path)
print(f"Saved: {out_path}")
