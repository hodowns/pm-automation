from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

BLUE = RGBColor(0x15, 0x65, 0xA0)
DARK = RGBColor(0x1A, 0x24, 0x33)
MUTED = RGBColor(0x5B, 0x67, 0x78)
SOFT = RGBColor(0xF3, 0xF6, 0xFA)
LINE = RGBColor(0xD9, 0xDE, 0xE6)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Aptos"

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
s = prs.slides.add_slide(prs.slide_layouts[6])


def text(x, y, w, h, t, size, bold=False, color=DARK, align=PP_ALIGN.LEFT):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = t
    r.font.size, r.font.bold, r.font.name = Pt(size), bold, FONT
    r.font.color.rgb = color
    return tb


text(0.5, 0.3, 12.3, 0.6, "FY27 Reverse Plan Dashboard: one view of whether revenue lands in the quarter", 26, True)

# lead-in band
band = s.shapes.add_shape(1, Inches(0.5), Inches(1.05), Inches(12.33), Inches(0.85))
band.fill.solid(); band.fill.fore_color.rgb = SOFT
band.line.fill.background()
bar = s.shapes.add_shape(1, Inches(0.5), Inches(1.05), Inches(0.08), Inches(0.85))
bar.fill.solid(); bar.fill.fore_color.rgb = BLUE; bar.line.fill.background()
text(0.7, 1.1, 12.0, 0.75,
     "Revenue is recognized only when a milestone is complete and its MCC reaches RAO before the quarter deadline. "
     "PMs update their plans every Friday and the dashboard reads them directly, so Fiscal Friday reviews use current data with no manual compilation.",
     13, color=DARK)

rows = [
    ("Tab", "Question it answers", "What it adds"),
    ("Overview", "What needs action now?", "Full task-level detail, risks and mitigations. Where PMs act."),
    ("By Market", "How is each market doing?", "Revenue and on-track / at-risk / blocked by market; shows markets with no plans loaded."),
    ("Portfolio", "How is this PM or project doing?", "Filter by PM, market or project; plain-English summary with key blocker and next deadline."),
    ("History", "Are plans kept current?", "Dated log of who created or updated each plan. Evidence of Friday discipline."),
    ("Calendar", "What lands when vs. finance deadlines?", "Plan dates over the FY27 fiscal calendar: MCC deadlines, 5-day rule, OP/compliance, holidays."),
    ("Leadership", "What is the one-page answer?", "Printable pipeline, close forecast, by-market table, Needs Attention list, next 14 days."),
    ("Data Audit", "Can we trust the numbers?", "Automated checks for missing PID/PM/revenue/MCC date, duplicates, overdue tasks, unflagged slips."),
]
tbl = s.shapes.add_table(len(rows), 3, Inches(0.5), Inches(2.1), Inches(12.33), Inches(4.4)).table
for i, w in enumerate((1.7, 3.6, 7.03)):
    tbl.columns[i].width = Inches(w)
for r, row in enumerate(rows):
    for c, val in enumerate(row):
        cell = tbl.cell(r, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = BLUE if r == 0 else WHITE
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_top = cell.margin_bottom = Inches(0.04)
        p = cell.text_frame.paragraphs[0]
        run = p.add_run(); run.text = val
        run.font.name = FONT
        run.font.size = Pt(12 if r else 12.5)
        run.font.bold = (r == 0) or (c == 0)
        run.font.color.rgb = WHITE if r == 0 else DARK

text(0.5, 6.7, 12.3, 0.5,
     "Act: Overview   |   Slice: By Market, Portfolio   |   Look ahead/back: Calendar, History   |   Decide: Leadership   |   Trust: Data Audit",
     12, True, BLUE, PP_ALIGN.CENTER)

prs.save(r"C:\Users\hodowns\Claude Projects VS CODE\reverse_plan_dashboard_leadership_slide.pptx")
print("saved")
