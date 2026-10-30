"""
Generate COH WxCC Project Plan Excel workbook.
PID: F59190 | TX-COH WxCC-F59190
"""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os

# ── Colors ──────────────────────────────────────────────────────────────────
CISCO_BLUE       = "1565A0"   # header
CISCO_BLUE_MED   = "2E75B6"   # phase rows
CISCO_BLUE_LIGHT = "BDD7EE"   # phase row alt
ALT_ROW          = "EBF2FA"   # task alt row
WHITE            = "FFFFFF"
DARK_TEXT        = "1F2D3D"
HEADER_TEXT      = "FFFFFF"

# ── Helpers ──────────────────────────────────────────────────────────────────
def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color=DARK_TEXT, size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name="Aptos")

def border(style="thin"):
    s = Side(style=style, color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def align(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def apply_header(ws, row, col, value, width_hint=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = fill(CISCO_BLUE)
    cell.font = font(bold=True, color=HEADER_TEXT, size=10)
    cell.alignment = align(h="center")
    cell.border = border()
    return cell

def apply_phase(ws, row, col_count, label):
    for c in range(1, col_count + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill(CISCO_BLUE_MED)
        cell.font = font(bold=True, color=HEADER_TEXT, size=10)
        cell.alignment = align(h="left")
        cell.border = border()
    ws.cell(row=row, column=1, value=label)

def apply_task(ws, row, values, alt=False):
    bg = ALT_ROW if alt else WHITE
    for c, val in enumerate(values, start=1):
        cell = ws.cell(row=row, column=c, value=val)
        cell.fill = fill(bg)
        cell.font = font()
        cell.border = border()
        if isinstance(val, datetime):
            cell.number_format = "MM/DD/YYYY"
            cell.alignment = align(h="center", wrap=False)
        elif c == 1:  # Task #
            cell.alignment = align(h="center", wrap=False)
        elif c in (8, 11, 12):  # numeric cols
            cell.alignment = align(h="center", wrap=False)
        else:
            cell.alignment = align()

# ── Task data ────────────────────────────────────────────────────────────────
# Columns: Task#, Task Name, Description, Deliverables, Assignee,
#          Plan Start, Plan Finish, Plan Duration, Act Start, Act Finish,
#          Act Duration, Status, % Complete

NA = None  # blank date placeholder

tasks = [
    # PHASE 1 marker (no data cols)
    ("PH1", None),

    ("1.1", ["1.1", "Project Kickoff Meeting",
     "Introduce team, align on objectives and SOW, establish cadence and communication plan, confirm roles and responsibilities",
     "Meeting Minutes; Project Charter; Communication Plan",
     "Houston Downs / Cisco CX / Netsync / CoH",
     datetime(2026,11,2), datetime(2026,11,2), 1, NA, NA, NA, "Not Started", 0]),

    ("1.2", ["1.2", "A2Q – Assessment to Quote",
     "Gather requirements, review current WxCCE environment, complete requirements questionnaire, conduct stakeholder interviews",
     "Requirements Questionnaire; A2Q Report",
     "Netsync",
     datetime(2026,11,3), datetime(2026,11,13), 9, NA, NA, NA, "Not Started", 0]),

    ("1.3", ["1.3", "Review Existing Documentation",
     "Review current WxCCE call flow scripts (102 total), integrations (WFM, CRM, recording), and network/PSTN architecture",
     "Documentation Review Summary; Script Inventory",
     "Cisco CX / Netsync",
     datetime(2026,11,3), datetime(2026,11,13), 9, NA, NA, NA, "Not Started", 0]),

    ("1.4", ["1.4", "Gap Analysis",
     "Identify gaps between current WxCCE capabilities and planned WxCC deployment; document integrations at risk, risks, and dependencies for Integrations tab (WFM, Recording, CRM, Custom Dashboard, PSTN)",
     "Gap Analysis Report",
     "Cisco CX / Netsync",
     datetime(2026,11,16), datetime(2026,11,20), 5, NA, NA, NA, "Not Started", 0]),

    ("1.5", ["1.5", "Remote Design Workshop (1 Day)",
     "Conduct one-day remote Design Workshop with CoH stakeholders; align on WxCC architecture, call flow design, AI Agent scope, integration approach, and migration strategy",
     "Workshop Agenda & Notes",
     "Cisco CX / Netsync / CoH",
     datetime(2026,11,23), datetime(2026,11,23), 1, NA, NA, NA, "Not Started", 0]),

    ("1.6", ["1.6", "Draft Solution Design Document (SDD)",
     "Document WxCC tenant architecture, 102 call flows (311, HPW, CRTS, Enterprise), AI Agent for CRTS, CRM/WFM/PSTN integrations, migration steps, DID strategy, agent desktop profiles",
     "Solution Design Document (Draft)",
     "Cisco CX / Netsync",
     datetime(2026,11,24), datetime(2026,12,4), 9, NA, NA, NA, "Not Started", 0]),

    ("1.7", ["1.7", "SDD Review with Customer",
     "Present SDD to CoH stakeholders for review and feedback; conduct review sessions per LOB (311, Courts, HPW, Enterprise); incorporate revisions",
     "Revised SDD with CoH Comments",
     "Cisco CX / Netsync / CoH",
     datetime(2026,12,7), datetime(2026,12,11), 5, NA, NA, NA, "Not Started", 0]),

    ("1.8", ["1.8", "Customer Sign-Off on SDD",
     "Obtain written approval of final SDD from CoH; signed SDD becomes the implementation baseline",
     "Signed / Approved SDD",
     "CoH",
     datetime(2026,12,14), datetime(2026,12,18), 5, NA, NA, NA, "Not Started", 0]),

    # PHASE 2
    ("PH2", None),

    ("2.1", ["2.1", "WxCC Tenant Setup & Configuration (Large Package)",
     "Set up and configure Webex Contact Center tenant (Large Package); configure org settings, PSTN trunk integration, entry points, queues",
     "Configured WxCC Tenant",
     "Cisco CX",
     datetime(2026,12,21), datetime(2027,1,9), 15, NA, NA, NA, "Not Started", 0]),

    ("2.2", ["2.2", "Onboard Agents & Supervisors (up to 500)",
     "Create agent and supervisor profiles; configure skill groups, teams, AUX codes, wrap-up codes, agent desk settings, and desktop profiles per LOB",
     "Agent/Supervisor Profiles Configuration",
     "Cisco CX",
     datetime(2027,1,12), datetime(2027,1,23), 10, NA, NA, NA, "Not Started", 0]),

    ("2.3", ["2.3", "Courts (CRTS) AI Agent Development",
     "Build Webex AI Agent for Courts CRTS VRU call flows (replaces CVP/Nuance); includes 7 CRTS outbound scripts and self-service CSMART inbound flow; VXML-to-Flow conversion",
     "AI Agent Flows – CRTS (7 scripts)",
     "Cisco CX",
     datetime(2027,1,12), datetime(2027,1,30), 15, NA, NA, NA, "Not Started", 0]),

    ("2.4", ["2.4", "311 Call Flow Development",
     "Migrate 311 LOB call flows from CVP to WxCC Flow Designer; DTMF IVR up to 3 levels depth; includes TTY/NexTalk routing (forward to 832-582-2181, queue to 311_TTD_ATT_SG skill); English and Spanish",
     "311 Call Flows in WxCC Flow Designer",
     "Cisco CX",
     datetime(2027,1,26), datetime(2027,2,13), 15, NA, NA, NA, "Not Started", 0]),

    ("2.5", ["2.5", "HPW (Houston Public Works) Call Flow Development",
     "Migrate HPW call flows including HPW Payment IVR (SS3/SS4/SS5); configure English and Spanish (es-mx, es-us) prompts; Western Union/ACI payment REST API integration",
     "HPW Call Flows in WxCC; HPW Payment IVR",
     "Cisco CX",
     datetime(2027,1,26), datetime(2027,2,13), 15, NA, NA, NA, "Not Started", 0]),

    ("2.6", ["2.6", "Enterprise / Other LOB Call Flow Development",
     "Migrate remaining LOB flows: CAS, HPC, Enterprise, Parks IVR, Callback flows (CallbackEntry, CallbackEntryHHD, EWT, HuCallback); all DTMF-based per scope",
     "Enterprise/Other Call Flows in WxCC",
     "Cisco CX",
     datetime(2027,2,1), datetime(2027,2,19), 15, NA, NA, NA, "Not Started", 0]),

    ("2.7", ["2.7", "Native Agent Desktop Configuration",
     "Configure WxCC native Webex Agent desktop with desktop variables and profiles per team; no CTI or custom desktop plugins (Ingenious CTI OOS per scope exclusions)",
     "Desktop Configuration",
     "Cisco CX",
     datetime(2027,2,9), datetime(2027,2,19), 9, NA, NA, NA, "Not Started", 0]),

    ("2.8", ["2.8", "CRM Integration",
     "Configure one out-of-box CRM connector (SNOW or SFDC) for native data integration; coordinate with Netsync on agent data display requirements",
     "CRM Integration Configuration",
     "Netsync",
     datetime(2027,1,26), datetime(2027,2,13), 15, NA, NA, NA, "Not Started", 0]),

    ("2.9", ["2.9", "WFM Integration",
     "Configure Workforce Management integration (Calabrio or custom connector); covers 311 and HPW/CAS user groups; Teleopti replacement discussion with CoH",
     "WFM Integration Configuration",
     "Netsync",
     datetime(2027,1,26), datetime(2027,2,13), 15, NA, NA, NA, "Not Started", 0]),

    ("2.10", ["2.10", "Standard Reports & Dashboards",
     "Configure WxCC standard out-of-box reports (109 available); provide up to 5 custom dashboard visualizations in Analyzer; knowledge transfer on reporting concepts",
     "Standard Reports; Up to 5 Custom Dashboards",
     "Cisco CX",
     datetime(2027,2,16), datetime(2027,2,26), 9, NA, NA, NA, "Not Started", 0]),

    ("2.11", ["2.11", "UC/PSTN Integration Validation",
     "Validate call delivery from enterprise/PSTN to WxCC cloud; test agent site connectivity; validate local GW configs (311, Courts, Enterprise, Agent cluster) — GW config execution is CoH/Netsync responsibility",
     "Integration Validation Report",
     "Cisco CX / CoH / Netsync",
     datetime(2027,2,23), datetime(2027,3,5), 9, NA, NA, NA, "Not Started", 0]),

    ("2.12", ["2.12", "Build Review Checkpoint / Go-to-Test Sign-Off",
     "Internal review of full build completeness; all LOBs configured; confirm all components ready for testing; obtain CoH go-to-test sign-off",
     "Build Checklist; Go-to-Test Sign-Off",
     "Cisco CX / Netsync / CoH",
     datetime(2027,3,8), datetime(2027,3,12), 5, NA, NA, NA, "Not Started", 0]),

    # PHASE 3
    ("PH3", None),

    ("3.1", ["3.1", "Test Plan Development",
     "Develop System Acceptance Test Plan (SATP) with test cases for all LOBs and features; document expected results, pass/fail criteria, test tool requirements",
     "System Acceptance Test Plan (SATP)",
     "Cisco CX",
     datetime(2027,3,15), datetime(2027,3,19), 5, NA, NA, NA, "Not Started", 0]),

    ("3.2", ["3.2", "Functional System Testing (Cisco CX)",
     "Execute functional tests on WxCC solution: all call flows, AI Agent (CRTS), desktop, integrations, reporting; document results against SATP",
     "Functional Test Results",
     "Cisco CX",
     datetime(2027,3,22), datetime(2027,3,26), 5, NA, NA, NA, "Not Started", 0]),

    ("3.3", ["3.3", "Customer UAT (User Acceptance Testing)",
     "CoH and Netsync execute UAT; test all LOBs (311, Courts, HPW, Enterprise); validate agent and supervisor workflows; per-LOB sign-off",
     "UAT Test Results; Per-LOB Sign-Off Sheets",
     "Netsync / CoH",
     datetime(2027,3,29), datetime(2027,4,2), 5, NA, NA, NA, "Not Started", 0]),

    ("3.4", ["3.4", "Defect Triage & Resolution",
     "Review all logged UAT defects; confirm reproducibility; assign severity/priority; resolve issues and retest fixes",
     "Defect Log; Resolution Report",
     "Cisco CX",
     datetime(2027,3,29), datetime(2027,4,9), 10, NA, NA, NA, "Not Started", 0]),

    ("3.5", ["3.5", "Regression Testing",
     "Re-run full regression test suite to confirm defect fixes have not introduced new issues; final validation pass before cutover",
     "Final Regression Test Results",
     "Cisco CX",
     datetime(2027,4,7), datetime(2027,4,9), 3, NA, NA, NA, "Not Started", 0]),

    ("3.6", ["3.6", "UAT Sign-Off",
     "Obtain written sign-off from CoH confirming successful UAT; baseline cleared for Migration Planning and Cutover",
     "Signed UAT Acceptance",
     "CoH",
     datetime(2027,4,9), datetime(2027,4,9), 1, NA, NA, NA, "Not Started", 0]),

    # PHASE 4
    ("PH4", None),

    ("4.1", ["4.1", "Agent/Supervisor Training (1–1.5 hrs)",
     "Train agents and supervisors on WxCC native desktop; cover agent workflow, supervisor monitoring; train-the-trainer model; up to 10 participants per session",
     "Training Completion; Attendance Record",
     "Cisco CX / WSA",
     datetime(2027,4,12), datetime(2027,4,12), 1, NA, NA, NA, "Not Started", 0]),

    ("4.2", ["4.2", "Reporting Training (2 hrs)",
     "Training on WxCC Analyzer: reporting basics, navigating 109 standard reports, creating custom reports from stock templates",
     "Training Completion; Attendance Record",
     "Cisco CX / WSA",
     datetime(2027,4,13), datetime(2027,4,13), 1, NA, NA, NA, "Not Started", 0]),

    ("4.3", ["4.3", "Administrator Training (2 hrs)",
     "Train WxCC platform administrators: admin best practices, platform management, RBAC, flow/config changes, business hours overrides",
     "Training Completion; Attendance Record",
     "Cisco CX / WSA",
     datetime(2027,4,14), datetime(2027,4,14), 1, NA, NA, NA, "Not Started", 0]),

    ("4.4", ["4.4", "Knowledge Transfer Workshop – Agent/Supervisor & Reporting",
     "Formal remote KT Workshop (1 business day, up to 10 participants): Agent/Supervisor workflows, reporting tools, analyzer dashboards",
     "KT Materials; KT Completion Record",
     "Cisco CX",
     datetime(2027,4,21), datetime(2027,4,21), 1, NA, NA, NA, "Not Started", 0]),

    ("4.5", ["4.5", "Knowledge Transfer Workshop – Administration & Day 2 Handover",
     "Formal remote KT Workshop (1 business day, up to 10 participants): Administrator role, platform management, Day 2 support model handover",
     "KT Materials; Day 2 Handover Package",
     "Cisco CX",
     datetime(2027,4,22), datetime(2027,4,22), 1, NA, NA, NA, "Not Started", 0]),

    # PHASE 5
    ("PH5", None),

    ("5.1", ["5.1", "Develop Migration Plan / Runbook",
     "Develop detailed migration runbook: LOB wave groupings, step-by-step cutover procedures, rollback procedures, communication plan, change window requests",
     "Migration Plan / Runbook",
     "Cisco CX",
     datetime(2027,3,22), datetime(2027,4,2), 10, NA, NA, NA, "Not Started", 0]),

    ("5.2", ["5.2", "DID / Number Planning",
     "Plan new DIDs for agents on Webex Calling/WebRTC; confirm inbound routing number strategy (reuse or new); plan PSTN number porting schedule — customer/partner responsible for porting",
     "DID Assignment Plan; Number Porting Schedule",
     "CoH / Netsync",
     datetime(2027,3,22), datetime(2027,4,2), 10, NA, NA, NA, "Not Started", 0]),

    ("5.3", ["5.3", "Gateway Configuration Planning",
     "Document GW config changes for local GW sites (311, Courts, Enterprise, Agent cluster) to route inbound numbers to WxCC CCP; CoH/Netsync execute with Cisco consulting",
     "Gateway Configuration Plan",
     "CoH / Netsync (Cisco consulting)",
     datetime(2027,3,29), datetime(2027,4,9), 10, NA, NA, NA, "Not Started", 0]),

    ("5.4", ["5.4", "LOB Migration Grouping & Cutover Schedule",
     "Define 4 migration LOB waves; assign cutover dates and change windows; confirm per-LOB sign-off checkpoints; align with Netsync and CoH stakeholders",
     "Cutover Schedule; LOB Wave Assignments",
     "Houston Downs / Cisco CX / CoH",
     datetime(2027,4,7), datetime(2027,4,11), 5, NA, NA, NA, "Not Started", 0]),

    ("5.5", ["5.5", "Pre-Migration Go/No-Go Checkpoint",
     "Final review of all migration prerequisites (SATP signed, GW configs complete, agents trained, DIDs assigned); obtain go/no-go sign-off from all parties",
     "Go/No-Go Sign-Off",
     "Cisco CX / Netsync / CoH",
     datetime(2027,4,22), datetime(2027,4,23), 2, NA, NA, NA, "Not Started", 0]),

    # PHASE 6
    ("PH6", None),

    ("6.1", ["6.1", "Migration 1 – LOB Group 1",
     "Execute first LOB wave migration from WxCCE to WxCC: move inbound numbers via CCP; agents log into WxCC with new DIDs; per-LOB sign-off at completion",
     "Migrated LOB Group 1",
     "Cisco CX / WSA",
     datetime(2027,4,26), datetime(2027,4,26), 1, NA, NA, NA, "Not Started", 0]),

    ("6.2", ["6.2", "Day 1 Hypercare – Migration 1 (72 hrs)",
     "Netsync/CX provide post-migration hypercare: monitor WxCC stability, resolve escalated issues, confirm LOB Group 1 operating normally on new platform",
     "Hypercare Support Log – Group 1",
     "Netsync / CX",
     datetime(2027,4,27), datetime(2027,4,29), 3, NA, NA, NA, "Not Started", 0]),

    ("6.3", ["6.3", "Migration 2 – LOB Group 2",
     "Execute second LOB wave migration from WxCCE to WxCC; move all numbers/DIDs for Group 2 via CCP; confirm per-LOB sign-off",
     "Migrated LOB Group 2",
     "Cisco CX / WSA",
     datetime(2027,5,4), datetime(2027,5,4), 1, NA, NA, NA, "Not Started", 0]),

    ("6.4", ["6.4", "Day 1 Hypercare – Migration 2 (72 hrs)",
     "Post-migration hypercare support for LOB Group 2 during first 72 hours on new platform",
     "Hypercare Support Log – Group 2",
     "Netsync / CX",
     datetime(2027,5,5), datetime(2027,5,7), 3, NA, NA, NA, "Not Started", 0]),

    ("6.5", ["6.5", "Migration 3 – LOB Group 3",
     "Execute third LOB wave migration; move all numbers/DIDs for Group 3 via CCP; confirm per-LOB sign-off",
     "Migrated LOB Group 3",
     "Cisco CX / WSA",
     datetime(2027,5,11), datetime(2027,5,11), 1, NA, NA, NA, "Not Started", 0]),

    ("6.6", ["6.6", "Day 1 Hypercare – Migration 3 (72 hrs)",
     "Post-migration hypercare support for LOB Group 3 during first 72 hours on new platform",
     "Hypercare Support Log – Group 3",
     "Netsync",
     datetime(2027,5,12), datetime(2027,5,14), 3, NA, NA, NA, "Not Started", 0]),

    ("6.7", ["6.7", "Migration 4 – LOB Group 4 (Final Migration)",
     "Execute final LOB wave migration; full cutover of all remaining agents from WxCCE to WxCC; WxCCE decommission planning begins",
     "Migrated LOB Group 4; Full Cutover Confirmation",
     "Cisco CX / WSA",
     datetime(2027,5,18), datetime(2027,5,18), 1, NA, NA, NA, "Not Started", 0]),

    ("6.8", ["6.8", "Day 1 Hypercare – Migration 4 (72 hrs)",
     "Post-migration hypercare for final LOB group; confirm full production stability across all 500 agents on WxCC",
     "Hypercare Support Log – Group 4",
     "Netsync",
     datetime(2027,5,19), datetime(2027,5,21), 3, NA, NA, NA, "Not Started", 0]),

    ("6.9", ["6.9", "Cutover Complete Sign-Off",
     "Confirm all LOBs successfully migrated; document any outstanding items; obtain formal cutover completion sign-off from CoH",
     "Cutover Complete Sign-Off",
     "CoH",
     datetime(2027,5,26), datetime(2027,5,28), 3, NA, NA, NA, "Not Started", 0]),

    # PHASE 7
    ("PH7", None),

    ("7.1", ["7.1", "Post-Implementation Support (3 Business Days)",
     "One Cisco remote resource provides consultative troubleshooting support; commences next business day after final migration; up to 3 business days or 1 calendar week",
     "Post-Implementation Support Log",
     "Cisco CX",
     datetime(2027,6,1), datetime(2027,6,3), 3, NA, NA, NA, "Not Started", 0]),

    ("7.2", ["7.2", "Outstanding Issues Resolution",
     "Document and resolve any remaining post-go-live issues; update defect/issue log; confirm all critical issues resolved before formal project closure",
     "Issue Resolution Report; Closed Issue Log",
     "Cisco CX / Netsync / CoH",
     datetime(2027,6,4), datetime(2027,6,9), 4, NA, NA, NA, "Not Started", 0]),

    ("7.3", ["7.3", "Final Project Handoff & Closure",
     "Formal transition to Day 2 support model (WSA); deliver all project artifacts (SDD, test results, KT materials, migration runbook); obtain final project closure sign-off from CoH",
     "Handoff Package; Project Closure Report",
     "Cisco CX / WSA / Houston Downs",
     datetime(2027,6,10), datetime(2027,6,11), 2, NA, NA, NA, "Not Started", 0]),
]

PHASE_LABELS = {
    "PH1": "PHASE 1 — DESIGN & DISCOVERY  (Nov 2, 2026 – Dec 18, 2026)",
    "PH2": "PHASE 2 — BUILD  (Dec 21, 2026 – Mar 12, 2027)",
    "PH3": "PHASE 3 — VALIDATE  (Mar 15, 2027 – Apr 9, 2027)",
    "PH4": "PHASE 4 — TRAINING  (Apr 12, 2027 – Apr 22, 2027)",
    "PH5": "PHASE 5 — MIGRATION PLANNING  (Mar 22, 2027 – Apr 23, 2027)",
    "PH6": "PHASE 6 — CUTOVER  (Apr 26, 2027 – May 28, 2027)",
    "PH7": "PHASE 7 — POST-IMPLEMENTATION SUPPORT & CLOSE  (Jun 1, 2027 – Jun 11, 2027)",
}

# ── Build workbook ───────────────────────────────────────────────────────────
wb = openpyxl.Workbook()

# ── Sheet 1: Project Information ─────────────────────────────────────────────
pi = wb.active
pi.title = "Project Information"
pi.sheet_view.showGridLines = False

def info_title(ws, row, text):
    cell = ws.cell(row=row, column=2, value=text)
    cell.fill = fill(CISCO_BLUE)
    cell.font = font(bold=True, color=HEADER_TEXT, size=12)
    cell.alignment = align(h="left")
    ws.row_dimensions[row].height = 22

def info_row(ws, row, label, value="", highlight=False):
    lc = ws.cell(row=row, column=2, value=label)
    vc = ws.cell(row=row, column=3, value=value)
    bg = ALT_ROW if highlight else WHITE
    for c in [lc, vc]:
        c.fill = fill(bg)
        c.border = border()
    lc.font = font(bold=True)
    lc.alignment = align(h="left", wrap=False)
    vc.font = font()
    vc.alignment = align(h="left", wrap=False)
    ws.row_dimensions[row].height = 16

pi.column_dimensions["A"].width = 2
pi.column_dimensions["B"].width = 30
pi.column_dimensions["C"].width = 45

pi.cell(row=1, column=2, value="City of Houston — Webex Contact Center Migration").font = Font(bold=True, color=CISCO_BLUE, size=16, name="Aptos")
pi.row_dimensions[1].height = 28

info_title(pi, 3, "Project Information")
info_row(pi, 4, "Project Name", "City of Houston Webex Contact Center (WxCC) Migration", True)
info_row(pi, 5, "PID", "F59190")
info_row(pi, 6, "Customer", "City of Houston (CoH)", True)
info_row(pi, 7, "Customer Experience Manager (CXM)", "")
info_row(pi, 8, "Project Manager", "Houston Downs", True)
info_row(pi, 9, "Partner", "Netsync")
info_row(pi, 10, "Account Manager", "", True)
info_row(pi, 11, "Service Sales", "")
info_row(pi, 12, "Project Start Date", "11/02/2026", True)
info_row(pi, 13, "Project End Date", "06/11/2027")
info_row(pi, 14, "SharePoint Folder", "TX-COH WxCC-F59190", True)
info_row(pi, 15, "SOW / Proposal", "WSA Detailed Proposal – COH WxCC CX Setup Assist Services")

info_title(pi, 17, "Scope Summary")
info_row(pi, 18, "Total Agents", "Up to 500 concurrent agents", True)
info_row(pi, 19, "Call Flow Scripts", "102 total (90 WxCC + 12 CVP including CRTS AI Agent conversion)")
info_row(pi, 20, "Lines of Business", "311, Courts (CRTS), HPW, CAS, HPC, Enterprise, Parks", True)
info_row(pi, 21, "Migration Events", "4 migration windows (LOB-wave based)")
info_row(pi, 22, "AI Agent", "CRTS Courts VRU (replaces CVP/Nuance) — IN SCOPE", True)
info_row(pi, 23, "CRM Integration", "1 out-of-box CRM connector (Netsync)")
info_row(pi, 24, "WFM Integration", "Calabrio or custom connector (Netsync)", True)
info_row(pi, 25, "Custom Reporting", "109 standard reports + up to 5 custom visualizations")
info_row(pi, 26, "Training Model", "Train-the-Trainer; Admin (2hr), Agent/Sup (1.5hr), Reporting (2hr)", True)

info_title(pi, 28, "Key Exclusions")
excl = [
    "WFM / Call Recording customization beyond native Webex storage",
    "Native Desktop only — desktop/CTI customization is OOS",
    "Gateway Configurations (CoH/Netsync responsibility; Cisco consulting only)",
    "Chat and Email channels",
    "Stress, load, and performance testing",
    "Custom reporting & visualizations beyond 5 dashboards",
    "On-premises CPE configurations",
    "Number porting (customer/partner responsibility)",
]
for i, e in enumerate(excl, start=29):
    info_row(pi, i, f"  • {e}", "", i % 2 == 0)

# ── Sheet 2: Project Plan ─────────────────────────────────────────────────────
pp = wb.create_sheet("Project Plan")
pp.sheet_view.showGridLines = False
pp.freeze_panes = "A3"

# Column definitions: (header, width)
COLS = [
    ("Task #",              9),
    ("Task Name",          36),
    ("Description",        52),
    ("Deliverables",       32),
    ("Assignee",           28),
    ("Planned\nStart Date", 14),
    ("Planned\nFinish Date",14),
    ("Planned\nDuration",   11),
    ("Actual\nStart Date",  14),
    ("Actual\nFinish Date", 14),
    ("Actual\nDuration",    11),
    ("Status",             14),
    ("% Complete",         11),
]

# Title banner
pp.merge_cells("A1:M1")
title_cell = pp.cell(row=1, column=1, value="City of Houston — Webex Contact Center Migration  |  Project Plan  |  PID: F59190")
title_cell.fill = fill(CISCO_BLUE)
title_cell.font = Font(bold=True, color=HEADER_TEXT, size=13, name="Aptos")
title_cell.alignment = Alignment(horizontal="center", vertical="center")
pp.row_dimensions[1].height = 26

# Column widths & headers
for ci, (hdr, w) in enumerate(COLS, start=1):
    pp.column_dimensions[get_column_letter(ci)].width = w
    apply_header(pp, 2, ci, hdr)
pp.row_dimensions[2].height = 28

# Write rows
row = 3
alt = False
for key, data in tasks:
    if data is None:
        # Phase header row
        apply_phase(pp, row, len(COLS), PHASE_LABELS[key])
        pp.merge_cells(f"A{row}:M{row}")
        pp.row_dimensions[row].height = 20
        row += 1
        alt = False
    else:
        apply_task(pp, row, data, alt)
        pp.row_dimensions[row].height = 52
        row += 1
        alt = not alt

# Auto-filter on header row
pp.auto_filter.ref = f"A2:{get_column_letter(len(COLS))}{row-1}"

# ── Sheet 3: Contact List (skeleton) ─────────────────────────────────────────
cl = wb.create_sheet("Contact List")
cl.sheet_view.showGridLines = False

cl.merge_cells("A1:F1")
ct = cl.cell(row=1, column=1, value="City of Houston WxCC  |  Contact List  |  PID: F59190")
ct.fill = fill(CISCO_BLUE)
ct.font = Font(bold=True, color=HEADER_TEXT, size=12, name="Aptos")
ct.alignment = Alignment(horizontal="center", vertical="center")
cl.row_dimensions[1].height = 24

contact_headers = ["Name", "Organization", "Role", "Email", "Phone", "Notes"]
contact_widths  = [22, 18, 26, 32, 16, 28]
for ci, (h, w) in enumerate(zip(contact_headers, contact_widths), start=1):
    apply_header(cl, 2, ci, h)
    cl.column_dimensions[get_column_letter(ci)].width = w

contacts = [
    ("Houston Downs",    "Cisco CX",  "Project Manager (PM)",              "hodowns@cisco.com",   "", ""),
    ("",                 "Cisco CX",  "Customer Experience Manager (CXM)", "",                    "", ""),
    ("",                 "Cisco CX",  "Lead Technical Engineer",           "",                    "", ""),
    ("Alec Franco",      "Cisco CX",  "Cat Center SME",                    "alecfran@cisco.com",  "", "F51119 ref"),
    ("",                 "Netsync",   "Netsync Project Lead",              "",                    "", ""),
    ("",                 "Netsync",   "Netsync Technical Lead",            "",                    "", ""),
    ("",                 "CoH",       "CoH Project Sponsor",               "",                    "", ""),
    ("",                 "CoH",       "CoH Project Manager",               "",                    "", ""),
    ("",                 "CoH",       "311 Operations Lead",               "",                    "", ""),
    ("",                 "CoH",       "Courts (CRTS) IT Lead",             "",                    "", ""),
    ("",                 "CoH",       "HPW IT Lead",                       "",                    "", ""),
    ("",                 "CoH",       "Network/GW Engineer",               "",                    "", ""),
]
for i, row_data in enumerate(contacts, start=3):
    bg = ALT_ROW if i % 2 == 0 else WHITE
    for ci, val in enumerate(row_data, start=1):
        cell = cl.cell(row=i, column=ci, value=val)
        cell.fill = fill(bg)
        cell.font = font()
        cell.border = border()
        cell.alignment = align(h="left", wrap=False)

# ── Save ─────────────────────────────────────────────────────────────────────
out_path = os.path.join(os.path.expanduser("~"), "Downloads",
                        "PMO_Workbook_TX_COH_WxCC_F59190.xlsx")
wb.save(out_path)
print(f"Saved: {out_path}")
