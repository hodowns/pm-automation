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

## Step 1 — Pull Data

Before asking any questions, silently gather context from the last 7 days:

1. **Outlook calendar** — meetings Houston organized or attended (via win32com)
2. **Outlook email** — sent items and inbox referencing any PID or customer name above
3. **Webex** — pull recent messages broadly using the token at `C:\Users\hodowns\webex-mcp-oauth\tokens.json`
4. **Previous check-ins** — read the most recent file in `C:\Users\hodowns\OneDrive - Cisco\Desktop\Check Ins\` to identify rolling priorities and completed items

Identify which manager(s) Houston connected with this week (calendar or email) — this will be used in the output.

---

## Step 2 — Ask Five Questions

Ask these one at a time. Enforce a **500-character limit** on free-text answers. Wait for each answer before moving to the next.

1. **Top priorities this week** — What were your 2–3 most important things you accomplished or moved forward this week?
2. **Blockers or risks** — Any blockers, escalations, or risks worth flagging? (Say "none" to skip.)
3. **Customer highlights** — Any notable customer interactions, wins, or concerns this week?
4. **AI / innovation work** — Did you build, test, or improve any Claude Code workflows or automations this week? If yes, estimate the time it saved or will save.
5. **Next week's focus** — What are your top 2–3 priorities for next week?

---

## Step 3 — Build the Check-In Document

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

Write the following file so the Playwright automation can fill TeamSpace automatically.

**Path:** `C:\Users\hodowns\Downloads\My_Performance.md`

The file must contain exactly one JSON code block in this structure. Synthesize each field from the answers and calendar data:

- `weekly_accomplishments` — concise bullet summary of Q1 + Q3 answers, max 490 chars, no newlines (single paragraph with semicolons as separators)
- `top_priorities_next_week` — Q5 answers, 2–3 lines, max 490 chars
- `customer_impact_value` — one sentence from Q3 (customer highlights)
- `blockers_and_support_needed` — Q2 answer, or `"None identified this week."` if skipped
- `internal_leadership_contributions` — Q4 answer, or `"No activity this week."` if skipped

```json
{
  "fields": {
    "weekly_accomplishments": "...",
    "top_priorities_next_week": "...",
    "customer_impact_value": "...",
    "blockers_and_support_needed": "...",
    "internal_leadership_contributions": "..."
  }
}
```

After writing, confirm: `Payload written to C:\Users\hodowns\Downloads\My_Performance.md`

---

## Step 7 — Run Playwright Automation

Offer to auto-fill TeamSpace now:

> "Payload ready. Run `my_performance_checkin.py` to auto-fill the portal?"

If Houston confirms, run:

```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\my_performance_checkin.py"
```

The script opens Chrome, navigates to `https://myperformance.cisco.com`, fills Step 1 (Reflect) and Step 2 (Plan), and saves as draft. Watch the output for ✓/✗ per field and screenshot paths if debugging is needed.
