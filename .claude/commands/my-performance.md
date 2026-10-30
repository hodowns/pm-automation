# /my-performance

You are helping Houston Downs (hodowns@cisco.com), a Cisco US PS Program Manager, complete his weekly check-in. Follow these steps exactly, in order.

---

## Profile

**Direct manager:** Greg Hopper (grhopper@cisco.com)
**Additional managers:**
- Jonathan Nissao — Texas projects
- Manoj Sathyaraj
- Becky Livergood — Northwest projects

**Active projects:**
| PID | Customer / Project |
|---|---|
| F51119 | City of Houston — LCS |
| E55085 | City of Houston — WAN Security |
| F29426 | Jacksonville Sheriff — VVS SAS |
| C19602 | Ada County — VVS JMS Integration |
| F24373 | UChicago Medicine — LCS 2Yr |
| F18212 | UChicago Medicine — SAS/LCS |
| F48852 | Maryland DHS — CAT ISE SAS |
| F40696 | NYS Senate — SAS |
| F46100 | State of Wisconsin — VVS SAS |
| C16507 | MN DOC — VVS |
| F43673 | Hennepin County — CC AST |
| F49302 | GDIT — Firewall SAS |
| F49303 | GDIT — Firewall SAS II |

**Goals & FY27 Priority Alignment:**
1. **Weekly Check-In Habit** — Complete My Performance weekly check-ins at or above the 80% organizational target throughout Q1–Q2 FY27, submitting top priorities each week linked to my active goals. → *All Priorities*
2. **AI-Driven PM Workflow Automation** — Deploy and actively use an integrated suite of AI-assisted tools (meeting recaps, action item tracking, automated nudges, financial reporting, schedule coordination) to reduce manual PM overhead by an estimated 3+ hours/week, and share the toolkit with at least 3 teammates by Q2 FY27 (January 2027). → *Leverage AI to Enhance Work*
3. **Project Delivery Health** — Maintain Green or Yellow project health status on all active PIDs throughout Q1–Q2 FY27, with 100% of project closures completed within 30 days of contract end and zero unplanned customer escalations. → *Drive Durable Growth*
4. **Customer Communication Cadence** — Establish and maintain a consistent weekly or bi-weekly customer touchpoint cadence across all active projects, with meeting minutes distributed within 24 hours of each meeting and all action items tracked and followed up within 5 business days. → *Simplify for Customer Outcomes*
5. **Professional Development** — Complete at least 2 Cisco-endorsed training courses or certifications relevant to project management or AI-enabled work by the end of Q2 FY27 (January 2027), applying learnings to at least one active project initiative. → *Leverage AI to Enhance Work*

**Output folder:** `C:\Users\hodowns\OneDrive - Cisco\Desktop\Check Ins\`

---

## Step 0 — Chrome CDP Setup (One-Time Per Session)

Before anything else, verify Chrome is running with remote debugging. If the user hasn't done this yet this session, show them the launch command:

**Windows PowerShell:**
```powershell
& "C:\Program Files\Google\Chrome\Application\chrome.exe" `
    --remote-debugging-port=9222 `
    --user-data-dir="$env:USERPROFILE\chrome-dev-profile"
```

In the opened Chrome window, navigate to `https://teamspace.cisco.com` and complete Cisco SSO + Duo if prompted. Once the dashboard loads, Playwright can take over via CDP — no further auth needed.

> Skip this step if Chrome is already open with `--remote-debugging-port=9222` from an earlier run this session.

---

## Step 1 — Pull Data

Before asking any questions, silently gather context from the last 7 days:

1. **Outlook calendar** — meetings Houston organized or attended (via win32com)
2. **Outlook email** — sent items and inbox referencing any PID or customer name above. Also scan for emails with subjects containing: "all hands", "all-hands", "FY27 strategy", "SLED all", "CX all hands", "cockpit", "AI enablement" — these contain Cisco-wide strategic goals to carry into the Goals section
3. **Webex** — pull recent messages broadly using the token at `C:\Users\hodowns\webex-mcp-oauth\tokens.json`
4. **Previous check-ins** — read the most recent file in `C:\Users\hodowns\OneDrive - Cisco\Desktop\Check Ins\`
5. **AI Champions (Claude Code Champions) meeting notes** — primary source for improvement insights and strategic themes. Read the most recent 2–3 files from:
   `C:\Users\hodowns\Cisco\SLED United CX Delivery Team - CUSTOMER PROJECTS\XXXX Innovation\AI Champions\Meeting Minutes\`
   Also check: `C:\Users\hodowns\Cisco\SLED United CX Delivery Team - CUSTOMER PROJECTS\_unfiled\` for Flight School and PMO-wide sessions
6. **Project meeting notes** — read the most recent notes from:
   - F51119: `C:\Users\hodowns\Cisco\SLED United CX Delivery Team - CUSTOMER PROJECTS\TX\TX-COH LCS-S- F51119\Meeting Minutes\`
   - F24373: `C:\Users\hodowns\Cisco\SLED United CX Delivery Team - CUSTOMER PROJECTS\IL\IL-S- UCMC LCS 2Yr F24373\Meeting Notes\`
   - C16507: `C:\Users\hodowns\Cisco\SLED United CX Delivery Team - CUSTOMER PROJECTS\MN\MN - MN DOC VVS -T- C16507\Meeting Minutes\`
   Extract improvement examples and actionable priorities from these notes.
7. **PMO Leadership call notes** — check `C:\Users\hodowns\OneDrive - Cisco\Desktop\Thursdays PMO Leadership call.docx` for organizational priorities that may come from all-hands or leadership sessions

Identify which manager(s) Houston connected with this week (calendar or email) — this will be used in the output.

---

## Step 2 — Infer Answers from Data

Using the data pulled in Step 1, **automatically determine** all five answers below. Do not ask Houston interactively — derive everything from meeting notes, emails, and transcripts.

1. **Top priorities this week** — The 2–3 highest-impact things accomplished or moved forward, based on meeting notes, sent emails, and action items closed.
2. **Blockers or risks** — Open action items, canceled/deferred decisions, technical unknowns, or escalation signals found in notes.
3. **Customer highlights** — Notable wins, milestone events, or concerns surfaced in external-facing sessions or email threads.
4. **AI / innovation work** — Cockpit/Claude Code sessions led or attended, skills tested or shipped, estimated time savings from automations.
5. **Next week's focus** — Committed follow-ups and due-date action items from this week's meetings; carry forward any unresolved items from the previous check-in.

After inferring all five, present a brief **"Here's what I found"** summary (one line per answer) and ask Houston to confirm or correct before building the document. This is the only interactive pause in the skill.

---

## Step 3 — Build the Check-In Document

**Narrative voice rules — apply to every sentence:**
- Lead every accomplishment with a strong action verb: *Drove*, *Delivered*, *Secured*, *Accelerated*, *Resolved*, *Established*, *Coordinated*, *Facilitated*, *Advanced*, *Executed*
- Connect every activity to a business outcome: time saved, risk eliminated, revenue protected, relationship strengthened, delivery accelerated
- Elevate routine PM work into strategic language: attending a meeting → *led governance touchpoint*; sending follow-up → *maintained delivery accountability*; fixing a config issue → *eliminated critical network stability risk*
- Every item must trace back to one of the 5 FY27 goals explicitly
- Tone: confident, executive-facing, portfolio-level — reads like someone managing a high-velocity book of business effortlessly

Using the data pulled in Step 1 and answers from Step 2, generate the check-in document in this format:

```
# Weekly Check-In — [Month Day, Year]
**Houston Downs** | US PS Program Manager
**Manager(s) connected with this week:** [auto-detected from calendar/email]

---

## This Week's Highlights
[2–3 sentence narrative synthesizing accomplishments, customer interactions, and meeting activity]

---

## Priority Updates

### [Goal 1: Weekly Check-In Habit] — All Priorities
[Note if check-in submitted this week. Running compliance rate if calculable from previous check-ins.]

### [Goal 2: AI-Driven PM Workflow Automation] — Leverage AI to Enhance Work
[Automations used or improved this week. Teammates shown the toolkit. Estimated time saved.]

### [Goal 3: Project Delivery Health] — Drive Durable Growth
[Project status updates, any closures completed, escalations or risks flagged.]

### [Goal 4: Customer Communication Cadence] — Simplify for Customer Outcomes
[Meeting minutes sent, action item follow-ups issued, touchpoints completed this week.]

### [Goal 5: Professional Development] — Leverage AI to Enhance Work
[Trainings completed or in progress. "No activity this week." if nothing.]

---

## Project Pulse
[One line per active PID that had activity this week. Skip silent ones.]
• F51119 — [one-line status]
• [PID] — [one-line status]

---

## Blockers / Risks
[From Q2. "None flagged." if skipped.]

---

## Next Week's Focus
• [Priority 1]
• [Priority 2]
• [Priority 3 if given]

---

## Rolling Items from Last Week
[Any priorities from the previous check-in that didn't appear as completed this week — carry them forward here.]

---
*Generated by /my-performance | Claude Code*
```

---

## Step 4 — Flag Unaligned Priorities

If any activity from this week doesn't fit cleanly under one of the four goals, flag it at the end of the document and suggest a new goal to capture it going forward.

---

## Step 5 — Save

Save the completed document as `YYYY-MM-DD-checkin.md` in `C:\Users\hodowns\OneDrive - Cisco\Desktop\Check Ins\`.

Confirm the filename and path after saving.

---

## Step 6 — Write JSON Payload for Playwright

Write `C:\Users\hodowns\Downloads\performance_data.json` as a proper JSON file (not embedded in markdown). This is what `my_performance_checkin.py` reads first.

Synthesize each field from the conversation data, applying the hyper-productive narrative voice from Step 3:

- `weekly_accomplishments` — key accomplishments + customer highlights (Q1 + Q3); bullet symbols (•) with semicolons; **max 490 chars; no literal newlines**
- `improvement_insight` — the **strongest 1–2 items** found in transcripts that answer "What's one improvement you made last week—for yourself, your team, or your customers?" Pull from AI Champions calls first (files in `XXXX Innovation\AI Champions\Meeting Minutes\`), then F51119/F24373/C16507 meeting notes. Frame using executive language. **Max 490 chars; no newlines.** Examples of strong improvement items: "Drove GitHub Champions membership from 22→42 in one week through structured Champion-led outreach, accelerating PMO AI adoption timeline"; "Delivered live issue→PR→merge Claude demo at 50+ PMO session, establishing GitHub as the standard workflow for skill contributions"; "Identified Haiku model as drop-in replacement for routine tasks at 50% of Sonnet cost, reducing PMO token spend without quality degradation"
- `top_priorities_next_week` — Q5 priorities, numbered; **max 490 chars**
- `customer_impact_value` — one strong sentence from Q3
- `blockers_and_support_needed` — Q2, or `"None at this time."` if skipped
- `internal_leadership_contributions` — Q4, or `"No activity this week."` if skipped
- `new_priorities` — array of 4–6 priority objects extracted from transcripts/emails, each mapped to the best-fit FY27 goal. Use judgment on priority ranking.

For `new_goals`: extract 2–4 broad, organization-level strategic goals from PMO Leadership call notes, AI Champions sessions, Flight School notes, or all-hands emails found in Step 1. These are **Cisco-wide priorities, not personal goals** — executive strategy language, broad scope. If no explicit all-hands source is found, derive from the themes in recent AI Champions and PMO leadership sessions. Include only goals not already in the system.

**Reference goals found in Champions/Flight School content (use as starting point, refine from latest session notes):**
- *Scale AI-Led Delivery Across US PS PMO* — Embed AI-assisted workflows (Claude Code, Cockpit) into standard delivery practice for all Program Managers, targeting measurable reduction in administrative overhead and increased capacity for customer-facing work
- *Drive Customer Outcome Velocity* — Compress time-to-value for all active SLED PS engagements through proactive health metrics, structured QBR cadences, and outcome tracking tied to contract milestones  
- *PMO Enablement & Capability Building* — Establish a sustainable knowledge-sharing infrastructure (Champions program, training dashboards, shared workflow libraries) that scales AI adoption beyond early adopters to the full PMO population

```json
{
  "employee_name": "Houston Downs",
  "employee_email": "hodowns@cisco.com",
  "role": "Customer Program Manager",
  "organization": "SLED West PMO | Cisco Customer Experience (CX)",
  "reporting_period": "Week Ending [date from check-in]",
  "new_goals": [
    { "title": "...", "description": "One-sentence strategic objective — broad, org-level, executive tone" },
    { "title": "...", "description": "..." }
  ],
  "fields": {
    "weekly_accomplishments": "...",
    "improvement_insight": "...",
    "top_priorities_next_week": "...",
    "customer_impact_value": "...",
    "blockers_and_support_needed": "...",
    "internal_leadership_contributions": "...",
    "new_priorities": [
      { "text": "...", "goal": "Goal title here" },
      { "text": "...", "goal": "Goal title here" }
    ]
  }
}
```

Confirm: `Payload written → C:\Users\hodowns\Downloads\performance_data.json`

---

## Step 7 — Run Playwright Automation

Confirm Chrome CDP is running (Step 0), then offer:

> "Payload written. Ready to auto-fill TeamSpace via Playwright?"

If Houston confirms, run:

```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\my_performance_checkin.py"
```

The script connects to Chrome on port 9222, navigates to `teamspace.cisco.com`, then executes the full wizard:

**Step 1 — Reflect:**
1. Selects **Chat/Email** for "When you communicated with your leader last week, which was most helpful?"
2. Fills the improvement textarea from `improvement_insight`
3. Resolves all last week's priorities (checks each checkbox)
4. Clicks **Save & continue**

**Step 2 — Plan:**
1. Clicks **Add Priority** for each item in `new_priorities` (4–6 items), maps each to its goal
2. Clicks **Save & continue**

**Step 3 — Review & Submit:**
1. Clicks **Submit** (full submission — not draft)

Watch the output for `✓`/`✗` per field. Screenshots saved to the script directory. If CDP fails, re-run Step 0.
