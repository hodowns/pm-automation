"""
Generate COH WxCC Kickoff Presentation
PID: F59190 | TX-COH WxCC-F59190
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from lxml import etree
import copy, os

# ── Brand colors ─────────────────────────────────────────────────────────────
NAVY       = RGBColor(0x1B, 0x2A, 0x3B)
BLUE       = RGBColor(0x15, 0x65, 0xA0)
BLUE_LIGHT = RGBColor(0x04, 0x9F, 0xD9)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
OFFWHITE   = RGBColor(0xF5, 0xF7, 0xFA)
GRAY       = RGBColor(0x6B, 0x6B, 0x6B)
DARK       = RGBColor(0x1F, 0x2D, 0x3D)
ACCENT     = RGBColor(0xEB, 0xF2, 0xFA)
GREEN      = RGBColor(0x2E, 0x86, 0x48)
ORANGE     = RGBColor(0xD4, 0x6B, 0x08)

# Phase colors
PH_COLORS = [
    RGBColor(0x15, 0x65, 0xA0),  # Ph1 - blue
    RGBColor(0x04, 0x9F, 0xD9),  # Ph2 - light blue
    RGBColor(0x1B, 0x7A, 0x5E),  # Ph3 - teal
    RGBColor(0x6A, 0x5A, 0xCD),  # Ph4 - purple
    RGBColor(0xD4, 0x6B, 0x08),  # Ph5 - orange
    RGBColor(0xC0, 0x39, 0x2B),  # Ph6 - red
    RGBColor(0x2E, 0x86, 0x48),  # Ph7 - green
]

W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

BLANK = prs.slide_layouts[6]  # blank layout

# ── Low-level helpers ─────────────────────────────────────────────────────────
def rgb_hex(c: RGBColor) -> str:
    return f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"

def solid_fill(shape, color: RGBColor):
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_rect(slide, l, t, w, h, color: RGBColor):
    shp = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    solid_fill(shp, color)
    shp.line.fill.background()
    return shp

def add_textbox(slide, l, t, w, h, text, size=18, bold=False, color=DARK,
                align=PP_ALIGN.LEFT, italic=False, wrap=True):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tb.word_wrap = wrap
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Aptos"
    return tb

def add_para(tf, text, size=14, bold=False, color=DARK,
             align=PP_ALIGN.LEFT, space_before=0, indent=False, italic=False):
    p = tf.add_paragraph()
    p.alignment = align
    if indent:
        p.level = 1
    if space_before:
        p.space_before = Pt(space_before)
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Aptos"
    return p

def footer_bar(slide, text="City of Houston — Webex Contact Center Migration  |  PID: F59190  |  CONFIDENTIAL"):
    add_rect(slide, 0, 7.1, 13.33, 0.38, NAVY)
    add_textbox(slide, 0.2, 7.12, 12, 0.32, text,
                size=9, color=WHITE, align=PP_ALIGN.LEFT)

def slide_header(slide, title, subtitle=None):
    add_rect(slide, 0, 0, 13.33, 1.1, BLUE)
    add_textbox(slide, 0.4, 0.1, 12.5, 0.72, title,
                size=24, bold=True, color=WHITE)
    if subtitle:
        add_textbox(slide, 0.4, 0.75, 12.5, 0.36, subtitle,
                    size=13, color=ACCENT, italic=True)

def section_label(slide, l, t, w, h, text, color=BLUE):
    add_rect(slide, l, t, w, h, color)
    add_textbox(slide, l+0.08, t+0.03, w-0.16, h-0.06, text,
                size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

def bullet_box(slide, l, t, w, h, title, bullets,
               title_color=BLUE, bg_color=None, bsize=13, tsize=13):
    if bg_color:
        add_rect(slide, l, t, w, h, bg_color)
    # title bar
    add_rect(slide, l, t, w, 0.35, title_color)
    add_textbox(slide, l+0.1, t+0.04, w-0.2, 0.28, title,
                size=tsize, bold=True, color=WHITE)
    tb = slide.shapes.add_textbox(
        Inches(l+0.1), Inches(t+0.4), Inches(w-0.2), Inches(h-0.5))
    tb.word_wrap = True
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for b in bullets:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(2)
        run = p.add_run()
        run.text = f"• {b}"
        run.font.size = Pt(bsize)
        run.font.color.rgb = DARK
        run.font.name = "Aptos"
    return tb

def add_table(slide, l, t, w, h, headers, rows,
              header_color=BLUE, alt_color=ACCENT, font_size=11):
    ncols = len(headers)
    nrows = len(rows) + 1
    tbl = slide.shapes.add_table(nrows, ncols,
        Inches(l), Inches(t), Inches(w), Inches(h)).table
    col_w = w / ncols
    for i in range(ncols):
        tbl.columns[i].width = Inches(col_w)
    # header row
    for ci, hdr in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.text = hdr
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.runs[0]
        run.font.size = Pt(font_size)
        run.font.bold = True
        run.font.color.rgb = WHITE
        run.font.name = "Aptos"
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_color
    # data rows
    for ri, row_data in enumerate(rows):
        bg = alt_color if ri % 2 == 0 else WHITE
        for ci, val in enumerate(row_data):
            cell = tbl.cell(ri+1, ci)
            cell.text = str(val) if val else ""
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
            run = p.runs[0] if p.runs else p.add_run()
            run.font.size = Pt(font_size-1)
            run.font.color.rgb = DARK
            run.font.name = "Aptos"
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
    return tbl

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 1 — TITLE
# ═══════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, NAVY)          # full bg
add_rect(sl, 0, 0, 0.25, 7.5, BLUE_LIGHT)     # left accent bar
add_rect(sl, 0, 5.9, 13.33, 1.6, BLUE)        # bottom bar

add_textbox(sl, 0.55, 0.7, 12, 0.5,
            "KICKOFF PRESENTATION", size=14, bold=True, color=BLUE_LIGHT,
            align=PP_ALIGN.LEFT)
add_textbox(sl, 0.55, 1.2, 12, 1.5,
            "City of Houston", size=42, bold=True, color=WHITE)
add_textbox(sl, 0.55, 2.6, 12, 0.9,
            "Webex Contact Center Migration", size=30, bold=False, color=ACCENT)
add_textbox(sl, 0.55, 3.6, 12, 0.5,
            "WxCCE → WxCC  |  Cloud Contact Center Deployment", size=16, color=GRAY, italic=True)

add_textbox(sl, 0.55, 5.95, 4, 0.35, "PID: F59190",
            size=12, bold=True, color=WHITE)
add_textbox(sl, 4.5, 5.95, 4.5, 0.35, "Partner: Netsync",
            size=12, color=WHITE)
add_textbox(sl, 9, 5.95, 4, 0.35, "Date: [MM/DD/YYYY]",
            size=12, color=ACCENT, align=PP_ALIGN.RIGHT)
add_textbox(sl, 0.55, 6.45, 12, 0.35,
            "Cisco Customer Experience  |  SLED United CX Delivery Team",
            size=11, color=ACCENT, italic=True)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 2 — AGENDA
# ═══════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, OFFWHITE)
slide_header(sl, "Agenda")
footer_bar(sl)

items = [
    ("01", "Introductions",                 "Meet the Cisco CX, Netsync, and City of Houston teams"),
    ("02", "Project Overview & Scope",      "Objectives, what's in scope, and what's not"),
    ("03", "Key Integrations & Open Items", "Current environment integrations and decisions needed"),
    ("04", "Project Phases & Timeline",     "7-phase delivery plan from Design through Closure"),
    ("05", "Roles & Responsibilities",      "Who owns what across Cisco CX, Netsync, and CoH"),
    ("06", "Communication & Cadence",       "Meeting schedule, SharePoint, and how we stay aligned"),
    ("07", "Customer Action Items",         "Pre-kickoff materials we're asking CoH to gather"),
    ("08", "Next Steps & Q&A",              "Immediate actions and open questions"),
]
for i, (num, title, desc) in enumerate(items):
    row = i % 4
    col = i // 4
    l = 0.5 + col * 6.5
    t = 1.3 + row * 1.45
    add_rect(sl, l, t, 6.1, 1.25, WHITE)
    add_rect(sl, l, t, 0.65, 1.25, BLUE)
    add_textbox(sl, l+0.07, t+0.28, 0.52, 0.6, num,
                size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(sl, l+0.75, t+0.05, 5.2, 0.38, title,
                size=14, bold=True, color=BLUE)
    add_textbox(sl, l+0.75, t+0.45, 5.2, 0.65, desc,
                size=11, color=GRAY)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 3 — TODAY'S OBJECTIVES
# ═══════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, OFFWHITE)
slide_header(sl, "Today's Objectives")
footer_bar(sl)

objectives = [
    ("🤝", "Align on Scope",
     "Confirm what's in scope, what's not, and resolve any open questions on the WSA proposal"),
    ("👥", "Introduce the Team",
     "Establish working relationships across Cisco CX, Netsync, and City of Houston"),
    ("📅", "Lock the Plan",
     "Walk through the 7-phase delivery timeline and confirm key milestone dates"),
    ("📋", "Assign Actions",
     "Leave today knowing exactly who needs to do what before Phase 1 begins"),
]
for i, (icon, title, body) in enumerate(objectives):
    l = 0.5 + (i % 2) * 6.4
    t = 1.35 + (i // 2) * 2.7
    add_rect(sl, l, t, 6.1, 2.4, WHITE)
    add_rect(sl, l, t, 6.1, 0.08, BLUE)           # top accent
    add_textbox(sl, l+0.2, t+0.15, 0.7, 0.7, icon, size=28)
    add_textbox(sl, l+1.0, t+0.12, 4.8, 0.5, title,
                size=16, bold=True, color=BLUE)
    add_textbox(sl, l+1.0, t+0.65, 4.8, 1.5, body,
                size=12, color=DARK)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 4 — MEET THE TEAM
# ═══════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, OFFWHITE)
slide_header(sl, "Meet the Team")
footer_bar(sl)

teams = [
    (BLUE,       "Cisco CX",          [
        "Project Manager:  Houston Downs",
        "CXM:  [TBD]",
        "Lead Engineer:  [TBD]",
        "",
        "hodowns@cisco.com",
    ]),
    (RGBColor(0x1B,0x7A,0x5E), "Netsync", [
        "Project Lead:  [TBD]",
        "Technical Lead:  [TBD]",
        "",
        "Partner — responsible for WFM,",
        "CRM, GW configs & UAT support",
    ]),
    (NAVY,       "City of Houston",   [
        "Executive Sponsor:  [TBD]",
        "Project Manager:  [TBD]",
        "311 Lead:  [TBD]",
        "Courts (CRTS) Lead:  [TBD]",
        "HPW Lead:  [TBD]",
        "Network/GW Engineer:  [TBD]",
    ]),
]
for i, (color, name, members) in enumerate(teams):
    l = 0.5 + i * 4.25
    add_rect(sl, l, 1.25, 3.95, 5.6, WHITE)
    add_rect(sl, l, 1.25, 3.95, 0.5, color)
    add_textbox(sl, l+0.12, 1.3, 3.7, 0.42, name,
                size=16, bold=True, color=WHITE)
    tb = sl.shapes.add_textbox(Inches(l+0.2), Inches(1.9),
                                Inches(3.55), Inches(4.8))
    tb.word_wrap = True
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for m in members:
        if first:
            p = tf.paragraphs[0]; first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(4)
        run = p.add_run()
        run.text = m
        run.font.size = Pt(13)
        run.font.color.rgb = GRAY if (not m or m.endswith(')') or '@' in m) else DARK
        run.font.italic = '@' in m or m == ""
        run.font.name = "Aptos"

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 5 — PROJECT OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, OFFWHITE)
slide_header(sl, "Project Overview")
footer_bar(sl)

# Left: key facts table
facts = [
    ("Customer",         "City of Houston (CoH)"),
    ("PID",              "F59190"),
    ("Partner",          "Netsync"),
    ("Migration",        "WxCCE → Webex Contact Center (Cloud)"),
    ("Agents",           "Up to 500 concurrent"),
    ("Call Flow Scripts","102 total across 4 lines of business"),
    ("AI Agent",         "CRTS Courts VRU (replaces CVP/Nuance)"),
    ("Migration Events", "4 LOB-wave migration windows"),
    ("Target Start",     "November 2, 2026"),
    ("Target Completion","June 11, 2027"),
    ("SharePoint",       "TX-COH WxCC-F59190"),
]
add_rect(sl, 0.4, 1.25, 5.8, 5.65, WHITE)
t = 1.35
for label, val in facts:
    bg = ACCENT if facts.index((label,val)) % 2 == 0 else WHITE
    add_rect(sl, 0.4, t, 5.8, 0.46, bg)
    add_textbox(sl, 0.55, t+0.07, 2.0, 0.34, label,
                size=11, bold=True, color=BLUE)
    add_textbox(sl, 2.6, t+0.07, 3.5, 0.34, val,
                size=11, color=DARK)
    t += 0.46

# Right: LOBs
add_rect(sl, 6.7, 1.25, 6.3, 0.5, BLUE)
add_textbox(sl, 6.9, 1.3, 6.0, 0.38, "Lines of Business in Scope",
            size=14, bold=True, color=WHITE)

lobs = [
    ("311",     "Call center, TTY/NexTalk, multi-language (EN/ES)"),
    ("Courts",  "CRTS AI Agent (replaces CVP/Nuance); outbound campaigns"),
    ("HPW",     "Houston Public Works; payment IVR (WU/ACI REST)"),
    ("Enterprise", "CAS, HPC, Parks, Callback, Enterprise routing"),
]
for i, (lob, desc) in enumerate(lobs):
    t = 1.9 + i * 1.15
    add_rect(sl, 6.7, t, 6.3, 1.05, WHITE)
    add_rect(sl, 6.7, t, 0.08, 1.05, PH_COLORS[i])
    add_textbox(sl, 6.85, t+0.06, 1.3, 0.38, lob,
                size=14, bold=True, color=BLUE)
    add_textbox(sl, 6.85, t+0.48, 5.9, 0.5, desc,
                size=11, color=GRAY)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 6 — SCOPE: WHAT WE'RE DELIVERING
# ═══════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, OFFWHITE)
slide_header(sl, "Scope of Work", "What Cisco CX is Delivering")
footer_bar(sl)

scope_cols = [
    ("Platform & Configuration", BLUE, [
        "WxCC Tenant setup — Large Package",
        "500 agent/supervisor profiles",
        "Skill groups, teams, AUX & wrap-up codes",
        "Desktop profiles per LOB (native)",
        "Standard out-of-box reports (109)",
        "Up to 5 custom Analyzer dashboards",
    ]),
    ("Call Flows & AI", RGBColor(0x04,0x9F,0xD9), [
        "102 call flow scripts migrated to Flow Designer",
        "311: DTMF IVR, TTY/NexTalk routing",
        "HPW: payment IVR (SS3/SS4/SS5)",
        "CRTS: Webex AI Agent (replaces CVP/Nuance)",
        "Enterprise/Parks/CAS/HPC call flows",
        "English + Spanish (es-mx, es-us) prompts",
    ]),
    ("Integrations & Testing", RGBColor(0x1B,0x7A,0x5E), [
        "1 out-of-box CRM connector",
        "WFM integration (Calabrio / custom)",
        "UC/PSTN integration validation",
        "Functional system testing",
        "UAT support + defect triage",
        "Regression testing",
    ]),
    ("Migration & Close", RGBColor(0x6A,0x5A,0xCD), [
        "4 LOB-wave migration windows",
        "72-hr hypercare per migration",
        "Admin training (2 hrs)",
        "Agent/Supervisor training (1.5 hrs)",
        "Reporting training (2 hrs)",
        "2 KT Workshops + 3-day post-impl. support",
    ]),
]
for i, (title, color, bullets) in enumerate(scope_cols):
    l = 0.35 + i * 3.2
    bullet_box(sl, l, 1.2, 3.05, 5.7, title, bullets,
               title_color=color, bg_color=WHITE, bsize=11)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 7 — SCOPE EXCLUSIONS
# ═══════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, OFFWHITE)
slide_header(sl, "Out of Scope", "Items Not Covered by the Current SOW — Change Request Required")
footer_bar(sl)

excl_left = [
    ("Desktop & CTI",        "Custom desktop / CTI customization (Ingenious CTI replacement OOS)"),
    ("Channels",             "Chat and Email channels"),
    ("Reporting",            "Custom reports & visualizations beyond 5 Analyzer dashboards"),
    ("Recording",            "Call recording requirements beyond native 30-day Webex storage"),
    ("WFM",                  "WFM replacement beyond Calabrio connector scope"),
    ("API / Web Dev",        "Custom REST API development or web services"),
]
excl_right = [
    ("Gateway Configs",      "GW configuration execution — CoH/Netsync responsibility; Cisco consulting only"),
    ("Number Porting",       "PSTN number porting — customer/partner responsibility"),
    ("Hardware / CPE",       "On-premises server/CPE configurations"),
    ("Testing",              "Stress, load, and performance testing"),
    ("Speech / AI (Other)",  "AI-enabled call flows beyond CRTS Courts; no speech except AI Agent"),
    ("Scope Changes",        "Any increased volume, complexity, or new configs require a Change Request"),
]
add_rect(sl, 0.4, 1.25, 12.5, 0.38, BLUE)
add_textbox(sl, 0.55, 1.3, 3.0, 0.3, "Category", size=11, bold=True, color=WHITE)
add_textbox(sl, 3.7, 1.3, 5.5, 0.3, "Exclusion Detail", size=11, bold=True, color=WHITE)
add_textbox(sl, 7.1, 1.3, 0.5, 0.3, "|", size=11, bold=True, color=WHITE)
add_textbox(sl, 7.3, 1.3, 2.0, 0.3, "Category", size=11, bold=True, color=WHITE)
add_textbox(sl, 9.4, 1.3, 3.3, 0.3, "Exclusion Detail", size=11, bold=True, color=WHITE)

for i in range(6):
    t = 1.7 + i * 0.82
    bg = ACCENT if i % 2 == 0 else WHITE
    add_rect(sl, 0.4, t, 6.4, 0.78, bg)
    add_textbox(sl, 0.55, t+0.07, 2.8, 0.62, excl_left[i][0],
                size=11, bold=True, color=BLUE)
    add_textbox(sl, 3.45, t+0.07, 3.25, 0.62, excl_left[i][1],
                size=10, color=DARK)
    add_rect(sl, 6.95, t, 6.0, 0.78, bg)
    add_textbox(sl, 7.05, t+0.07, 2.2, 0.62, excl_right[i][0],
                size=11, bold=True, color=BLUE)
    add_textbox(sl, 9.35, t+0.07, 3.5, 0.62, excl_right[i][1],
                size=10, color=DARK)

add_rect(sl, 0.4, 6.7, 12.5, 0.32, RGBColor(0xFF,0xF2,0xCC))
add_textbox(sl, 0.6, 6.72, 12, 0.28,
            "⚠   Any item above requires a formal Change Request — scope changes will be assessed for schedule and cost impact.",
            size=10, bold=True, color=ORANGE)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 8 — KEY INTEGRATIONS
# ═══════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, OFFWHITE)
slide_header(sl, "Key Integrations & Open Items",
             "Decisions needed before or during Phase 1 Design")
footer_bar(sl)

int_headers = ["Integration", "Current (WxCCE)", "WxCC Plan", "Owner", "Status"]
int_rows = [
    ["WFM",             "Teleopti",           "Calabrio or custom connector",              "Netsync",       "⚠ Demo needed"],
    ["Call Recording",  "Eleveo",             "Calabrio OR native 30-day OR future WFO",   "Cisco Acct/Netsync", "⚠ Decision needed"],
    ["Agent Desktop",   "Ingenious CTI",      "Native WxCC desktop (no CTI replacement)",  "Netsync",       "OOS — confirm"],
    ["ASR/TTS",         "Nuance (via CVP)",   "WxCC AI Agent (CRTS only); native TTS rest","Cisco CX",      "✅ IN SCOPE"],
    ["CRM",             "[TBD]",              "1 out-of-box CRM connector",                "Netsync",       "⚠ Confirm CRM"],
    ["Custom Dashboard","SquareO",            "WxCC Analyzer dashboards (up to 5 custom)", "Netsync",       "OOS beyond 5"],
    ["Payment IVR",     "Payment middleware", "REST API (WU/ACI) — proxy or public API",   "Netsync/CoH",   "⚠ API access needed"],
    ["PSTN/GW",         "SIP trunks (ATT)",   "CCP carrier; GW configs CoH/Netsync",       "CoH/Netsync",   "CoH action item"],
    ["TTY/NexTalk",     "SimpliciTTY",        "Existing NexTalk routing preserved",         "Netsync/CoH",   "✅ Confirmed plan"],
]
add_table(sl, 0.35, 1.2, 12.6, 5.75, int_headers, int_rows,
          header_color=BLUE, font_size=11)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 9 — PROJECT PHASES OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, OFFWHITE)
slide_header(sl, "Project Phases & Timeline")
footer_bar(sl)

phases = [
    ("1", "DESIGN &\nDISCOVERY",   "Nov 2 –\nDec 18, 2026",  PH_COLORS[0],
     ["Kickoff & A2Q", "Design Workshop", "SDD Draft + Sign-off"]),
    ("2", "BUILD",                  "Dec 21, 2026 –\nMar 12, 2027", PH_COLORS[1],
     ["Tenant Setup", "Call Flows (102)", "CRTS AI Agent", "Integrations"]),
    ("3", "VALIDATE",               "Mar 15 –\nApr 9, 2027",  PH_COLORS[2],
     ["System Testing", "Customer UAT", "Defect Resolution"]),
    ("4", "TRAINING",               "Apr 12–22,\n2027",       PH_COLORS[3],
     ["Admin Training", "Agent/Sup Training", "KT Workshops"]),
    ("5", "MIGRATION\nPLANNING",   "Mar 22 –\nApr 23, 2027", PH_COLORS[4],
     ["Migration Runbook", "DID Planning", "Go/No-Go Checkpoint"]),
    ("6", "CUTOVER",                "Apr 26 –\nMay 28, 2027", PH_COLORS[5],
     ["4 LOB Migrations", "72-hr Hypercare ea.", "Cutover Sign-off"]),
    ("7", "CLOSE",                  "Jun 1–11,\n2027",         PH_COLORS[6],
     ["Post-Impl. Support", "Issue Resolution", "Day 2 Handoff"]),
]

for i, (num, name, dates, color, bullets) in enumerate(phases):
    l = 0.35 + i * 1.82
    # Phase box
    add_rect(sl, l, 1.25, 1.65, 0.62, color)
    add_textbox(sl, l+0.05, 1.28, 0.38, 0.56, num,
                size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(sl, l+0.45, 1.3, 1.15, 0.56, name,
                size=10, bold=True, color=WHITE)
    # Date band
    add_rect(sl, l, 1.87, 1.65, 0.52, RGBColor(0xEB,0xF2,0xFA))
    add_textbox(sl, l+0.05, 1.9, 1.55, 0.46, dates,
                size=9, color=BLUE, align=PP_ALIGN.CENTER)
    # Bullet content
    add_rect(sl, l, 2.45, 1.65, 4.4, WHITE)
    tb = sl.shapes.add_textbox(
        Inches(l+0.1), Inches(2.55), Inches(1.45), Inches(4.2))
    tb.word_wrap = True
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for b in bullets:
        if first:
            p = tf.paragraphs[0]; first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(4)
        run = p.add_run()
        run.text = f"• {b}"
        run.font.size = Pt(10)
        run.font.color.rgb = DARK
        run.font.name = "Aptos"

# Connector arrows
for i in range(6):
    l = 1.95 + i * 1.82
    add_textbox(sl, l, 1.48, 0.25, 0.3, "▶",
                size=12, color=BLUE, align=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 10 — ROLES & RESPONSIBILITIES
# ═══════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, OFFWHITE)
slide_header(sl, "Roles & Responsibilities",
             "R = Responsible  |  A = Accountable  |  C = Consulted  |  I = Informed")
footer_bar(sl)

raci_headers = ["Phase / Activity", "Cisco CX", "Netsync", "CoH"]
raci_rows = [
    ["A2Q / Requirements Gathering",       "C",   "R/A", "C"],
    ["Solution Design Document (SDD)",     "R/A", "C",   "I"],
    ["Customer Sign-Off on SDD",           "I",   "I",   "R/A"],
    ["WxCC Tenant Setup & Configuration",  "R/A", "I",   "I"],
    ["Call Flow Development (102 scripts)","R/A", "C",   "I"],
    ["CRTS AI Agent Development",          "R/A", "I",   "C"],
    ["CRM Integration",                    "C",   "R/A", "C"],
    ["WFM Integration",                    "C",   "R/A", "C"],
    ["Gateway Configuration",              "C",   "R",   "A"],
    ["Functional System Testing",          "R/A", "C",   "I"],
    ["Customer UAT",                       "C",   "R",   "A"],
    ["Migration Plan / Runbook",           "R/A", "C",   "C"],
    ["Number Porting",                     "I",   "R",   "A"],
    ["LOB Cutover Execution",              "R/A", "R",   "C"],
    ["72-hr Hypercare",                    "C",   "R/A", "I"],
    ["Training Delivery",                  "R/A", "I",   "C"],
    ["Day 2 Handoff / Project Closure",    "R/A", "C",   "A"],
]
add_table(sl, 0.35, 1.2, 12.6, 5.75, raci_headers, raci_rows,
          header_color=BLUE, font_size=11)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 11 — COMMUNICATION & CADENCE
# ═══════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, OFFWHITE)
slide_header(sl, "Communication & Cadence")
footer_bar(sl)

meetings = [
    ("Weekly Status Call",       "Weekly",      "Cisco CX PM + Netsync + CoH PM",        "[Day/Time TBD]"),
    ("Technical Working Sessions","As Needed",   "Cisco CX Engineers + CoH LOB Leads",    "During Build phase"),
    ("Design Workshop",          "One-Time",     "Full team + CoH stakeholders",           "Week of Nov 23, 2026"),
    ("Migration Planning Reviews","Bi-Weekly",   "Cisco CX + Netsync + CoH PM/Net Eng",   "Phase 5 (Mar–Apr 2027)"),
    ("Pre-Migration Go/No-Go",   "Per Migration","All parties",                            "Before each cutover event"),
    ("Steering Committee",       "Monthly",      "Cisco CXM + CoH Executive Sponsor",     "[Day/Time TBD]"),
]
add_table(sl, 0.35, 1.25, 12.6, 3.8,
          ["Meeting", "Frequency", "Attendees", "When"],
          meetings, header_color=BLUE, font_size=11)

# Bottom info boxes
boxes = [
    (BLUE,  "SharePoint",   "TX-COH WxCC-F59190\nSLED United CX Delivery Team\nAll project documents filed here"),
    (RGBColor(0x1B,0x7A,0x5E), "Meeting Minutes",
                            "Distributed within 24 hrs of each call\nAction items tracked and followed up\nWebex recap emails filed to SharePoint"),
    (RGBColor(0x6A,0x5A,0xCD), "Issue / Risk Management",
                            "RAID Log maintained in PMO Workbook\nCritical issues escalated same-day\nChange Requests documented and approved before work begins"),
    (NAVY,  "Tools",        "Webex for meetings & messaging\nSharePoint / OneDrive for docs\nPMO Workbook F59190 for tracking"),
]
for i, (color, title, body) in enumerate(boxes):
    l = 0.35 + i * 3.2
    add_rect(sl, l, 5.25, 3.05, 1.62, WHITE)
    add_rect(sl, l, 5.25, 3.05, 0.35, color)
    add_textbox(sl, l+0.1, 5.27, 2.85, 0.3, title,
                size=12, bold=True, color=WHITE)
    add_textbox(sl, l+0.1, 5.65, 2.85, 1.1, body,
                size=10, color=DARK)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 12 — CUSTOMER ACTION ITEMS
# ═══════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, OFFWHITE)
slide_header(sl, "Customer Action Items",
             "Items CoH is asked to gather before Phase 1 begins")
footer_bar(sl)

action_cols = [
    ("Current Environment", BLUE, [
        "Active script inventory (102 scripts)\nwith LOB owner per script",
        "Current WxCCE topology/network diagram",
        "Firewall/port documentation",
    ]),
    ("Integrations", RGBColor(0x04,0x9F,0xD9), [
        "WFM (Teleopti) contract end date;\nopen to Calabrio demo?",
        "Recording (Eleveo) contract + retention requirements",
        "CRM system + agent screen-pop data requirements",
        "SquareO dashboard screenshots + LOB users",
        "Payment API docs (WU/ACI — public or proxy?)",
        "Ingenious CTI data fields + agent desktop screenshots",
    ]),
    ("Agent & Skills", RGBColor(0x1B,0x7A,0x5E), [
        "Agent count per LOB (confirm ~500)",
        "Current skill groups, teams, supervisor structure",
        "Full AUX and wrap-up code list",
        "LOB-specific desktop profile requirements",
    ]),
    ("Telephony & Access", RGBColor(0x6A,0x5A,0xCD), [
        "Full DID inventory (inbound, agent, admin)",
        "Number porting intent (reuse or new?)",
        "PSTN carrier contacts (AT&T 311 trunk, others)",
        "WxCCE read access for Cisco engineer",
        "LOB technical points of contact for Design Workshop",
    ]),
]
for i, (title, color, bullets) in enumerate(action_cols):
    l = 0.35 + i * 3.2
    bullet_box(sl, l, 1.2, 3.05, 5.72, title, bullets,
               title_color=color, bg_color=WHITE, bsize=10.5)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 13 — IMMEDIATE NEXT STEPS
# ═══════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, OFFWHITE)
slide_header(sl, "Immediate Next Steps")
footer_bar(sl)

next_steps = [
    (BLUE,  "Cisco CX",  [
        ("Engineer Assignment",    "Cisco CX to assign Lead Engineer to F59190",     "ASAP"),
        ("PMO Workbook",           "PMO Workbook F59190 filed to SharePoint — live", "Complete ✅"),
        ("Weekly Cadence",         "Confirm weekly status call day/time with CoH",    "This week"),
        ("Design Workshop Date",   "Propose Design Workshop date (target: Nov 23)",   "This week"),
    ]),
    (RGBColor(0x1B,0x7A,0x5E), "Netsync", [
        ("A2Q Kickoff",            "Begin A2Q assessment with CoH",                  "Nov 3"),
        ("Integration Review",     "Review WFM, recording, CRM requirements",        "Nov 3–13"),
        ("GW Config Discussion",   "Align with CoH Network team on GW plan",          "Phase 1"),
    ]),
    (NAVY,  "City of Houston", [
        ("Gather Action Items",    "Compile pre-kickoff materials (see prior slide)", "ASAP"),
        ("LOB Contacts",           "Confirm technical POC for each LOB",              "This week"),
        ("WxCCE Access",           "Provide Cisco read access to current environment","Before engineer assigned"),
        ("Confirm Start Date",     "Confirm Nov 2, 2026 as Phase 1 start",            "This week"),
    ]),
]
for i, (color, owner, steps) in enumerate(next_steps):
    l = 0.4 + i * 4.3
    add_rect(sl, l, 1.25, 4.1, 0.45, color)
    add_textbox(sl, l+0.15, 1.28, 3.8, 0.38, owner,
                size=15, bold=True, color=WHITE)
    for j, (item, detail, when) in enumerate(steps):
        t = 1.8 + j * 1.25
        bg = ACCENT if j % 2 == 0 else WHITE
        add_rect(sl, l, t, 4.1, 1.15, bg)
        add_rect(sl, l, t, 0.06, 1.15, color)
        add_textbox(sl, l+0.15, t+0.05, 3.75, 0.38, item,
                    size=12, bold=True, color=BLUE)
        add_textbox(sl, l+0.15, t+0.45, 2.9, 0.6, detail,
                    size=10, color=DARK)
        add_textbox(sl, l+3.05, t+0.45, 0.9, 0.4, when,
                    size=9, bold=True, color=color, align=PP_ALIGN.RIGHT)

# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 14 — THANK YOU / Q&A
# ═══════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, NAVY)
add_rect(sl, 0, 0, 0.35, 7.5, BLUE_LIGHT)
add_rect(sl, 0, 5.7, 13.33, 1.8, BLUE)

add_textbox(sl, 0.7, 1.0, 12, 0.7,
            "Questions & Discussion", size=14, bold=True, color=BLUE_LIGHT)
add_textbox(sl, 0.7, 1.6, 12, 1.4,
            "Thank You!", size=48, bold=True, color=WHITE)
add_textbox(sl, 0.7, 3.15, 12, 0.6,
            "We're excited to get this project off the ground with the City of Houston team.",
            size=16, color=ACCENT, italic=True)
add_textbox(sl, 0.7, 3.85, 12, 0.45,
            "Please share any open questions — we'll capture them as action items before we close.",
            size=13, color=GRAY)

add_textbox(sl, 0.7, 5.78, 4, 0.4, "Houston Downs",
            size=13, bold=True, color=WHITE)
add_textbox(sl, 0.7, 6.15, 4, 0.35, "hodowns@cisco.com",
            size=12, color=ACCENT, italic=True)
add_textbox(sl, 0.7, 6.5, 4, 0.35, "Project Manager, Cisco CX",
            size=11, color=ACCENT)
add_textbox(sl, 5.5, 5.78, 7.4, 1.1,
            "Cisco Customer Experience  |  SLED United CX Delivery Team\nPID: F59190  |  TX-COH WxCC-F59190",
            size=11, color=ACCENT, align=PP_ALIGN.RIGHT, italic=True)

# ═══════════════════════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════════════════════
out_dir = r'C:\Users\hodowns\Cisco\SLED United CX Delivery Team - CUSTOMER PROJECTS\TX\TX-COH WxCC-F59190'
out_path = os.path.join(out_dir, 'COH_WxCC_Kickoff_F59190.pptx')
prs.save(out_path)
print(f"Saved: {out_path}")
print(f"Slides: {len(prs.slides)}")
