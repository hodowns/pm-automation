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

## CX Americas FY27 Strategy — Grow. Defend. Transform.

Source: `FY27 Strategy.pdf` (CX Americas, September 2026). Full summary in memory (`reference_fy27_strategy.md`). Use this as the lens for **every** narrative, priority, and goal in this skill — the 5 personal goals above stay as the system-of-record, but each item should also ladder up to a pillar.

**Pillar 1 — Grow:** 10% y/y TS bookings (75% premium mix), 10% SW ARR, PS growth. PS Growth = put CX where it moves **customer lifetime value (LTV)**; accounts prioritized by future revenue potential, not current size; "durable growth, not discounted volume."
**Pillar 2 — Defend:** protect the installed base by **leading with Cisco IQ** (IQ Link on 100% of CSPC accounts by EoQ2; 6,800 of 8,500 target accounts; 12,000 customers by Q4); **every PSIRT actioned within 30 days** (auto-escalate at 10); premium mix on renewals; "no account is ever dropped."
**Pillar 3 — Transform:** CX Americas runs as one **agentic system**. Targets: 75% accounts instrumented with IQ, 75% tasks automated, **100% of PS delivered via playbook or as-code**. Obj 1: one value metric per function (outcomes, not activity). Obj 2: **redesign the work end-to-end — don't automate around it**. Obj 3: equip teams to work agentically so time shifts to judgment, customer outcomes, advisory. People are measured by customer outcomes, not hours delivered.

**General Manager mindset (the behavioral standard):** own the whole portfolio; drive a **proactive weekly agenda** (three things actively moving, not reacting); "the outcome is mine"; name top priorities cold; build cross-functional interdependence as leverage; "how can I?" not "I can't." Transform commitment = one value metric, evidence of behavior change in the team, and **at least one redesigned workflow underway (not planned — underway)**.

### How to apply this in the check-in

| Houston's activity | FY27 pillar | Strategic framing to use |
|---|---|---|
| Claude Code/Cockpit automations (recaps, nudges, status report, SDC automation, Fin Report) | **Transform** (Obj 2 & 3) | "Redesigned [workflow] end-to-end, agentic-first" — name the workflow, the hours shifted to judgment/customer work, and any teammate adopting it (= behavior change evidence). Never frame as "saved time on admin" alone. |
| Sharing toolkit with teammates, Champions sessions, Flight School | **Transform** (Obj 3) | "Equipped PMO to work agentically" — count adopters. |
| Delivery health, milestone/revenue forecasting, closures, PS cost/revenue compliance | **Grow** (PS) / **Defend** | Tie to durable PS growth and LTV; protect revenue recognized on milestones; no unplanned escalations. |
| Customer cadence, minutes, action-item follow-up | **Defend** + **Grow** | Governance touchpoints that protect the installed base and surface follow-on/expansion signals. |
| Anything touching Cisco IQ, IQ Link, Cat Center/SW adoption, PSIRT/security advisories, renewals, premium mix | **Defend** (or SW Grow) | Call it out explicitly — these are the headline FY27 metrics. Flag customer IQ/IQ Link status when notes mention it. |
| Expansion signals, follow-on SOW/AST discussions, new-scope requests | **Grow** | Frame as LTV opportunity, not just scope. |

**Voice rules added by the strategy:**
- Write as an owner with an agenda, not a reporter of events. Lead the highlights with the **three things actively being moved**.
- Blockers/risks use "how can I / here's the ask" language — never victim framing or "not my responsibility."
- Prefer outcome language (customer outcomes, LTV, risk retired, workflow redesigned) over hours/activity language.
- Name the pillar in each priority where it fits (e.g., "→ Transform").
- Do not invent IQ/PSIRT/LTV facts — only cite them if found in notes/emails. If none this week, omit rather than fabricate.

**Org-level `new_goals` (Step 6):** derive from the three pillars first (Grow/Defend/Transform), then from Champions/PMO notes. See reference goals in Step 6.

---

## Step 0 — Choose the Submission Route (connector first, no browser needed)

Houston keeps his own Chrome open for other work, so **do not ask him to close or relaunch Chrome.** Chrome hands a debug-port launch to an already-running session and ignores `--remote-debugging-port`, which silently breaks the CDP route. Pick the route in this order:

1. **MyPerformance connector (preferred).** Call `mcp__cisco-fetch-mcp__myperformance_get_weekly_checkin` (empty `week_start`). If it says "No valid session", call `mcp__cisco-fetch-mcp__login` with `profile_name: "myperformance"` (Touch ID / Duo approval, runs in the background). If it works, skip the rest of this step and use the **Step 7A** path.
2. **Playwright's own browser (fallback).** If the connector login fails, use the `mcp__playwright__browser_*` tools, which launch a separate window and profile and never touch Houston's Chrome. He signs in to `teamspace.cisco.com` once in that window. Use the **Step 7B** path.
3. **Chrome CDP on port 9222 (last resort).** Only if 1 and 2 both fail. Verify `curl http://localhost:9222/json/version` returns JSON before running the script. If it doesn't, the launch was absorbed by his running Chrome; do not retry, tell him which fallback to use instead.

### Chrome CDP setup (last resort only)

Show the launch command:

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

Before asking any questions, silently gather context. **Start with Step 1A (the MyAgent weekly message). It is the primary source and takes priority over everything else.** Then use the sources below only to fill gaps.

### Step 1A — MyAgent weekly message (PRIMARY source)

Houston's agent posts a weekly check-in draft to him in Webex every Friday around 8:12 AM Eastern.

- **Room:** direct room titled `MyAgent` (sender `circuitpulse@webex.bot`)
- **Message:** starts with `Weekly Performance Check-In (` followed by the date range, e.g. `Weekly Performance Check-In (Nov 13 – Nov 20, 2026)`
- **Auth:** token at `C:\Users\hodowns\webex-mcp-oauth\tokens.json` (needs `spark:rooms_read` and `spark:messages_read`). Never print the token.

Fetch it with the Webex REST API (run with `py -3 -I -X utf8`; `python` is not on PATH): list `rooms?type=direct&max=200`, pick the room titled `MyAgent`, read `messages?roomId=...&max=100`, and take the newest message whose text starts with `Weekly Performance Check-In (`. Skip the Monday "weekly reminder" and the daily "Secretary Morning Briefing" posts. The messages are newest-first.

Sections in the message and where they go:

| Message section | Use for |
|---|---|
| Key Accomplishments | `weekly_accomplishments`, Highlights, Project Pulse (one line per PID named) |
| Top Priorities for Next Week | `top_priorities_next_week` and `new_priorities` (Step 6) |
| Customer Value & Business Impact | `customer_impact_value` |
| Blockers & Risks | the check-in document's Blockers / Risks section **only**. Never the portal "What do you need from your leader" field (Step 2 rule still applies). |
| Team Space Execution Text | **Ignore.** It tells you to submit. Submission only happens on Houston's explicit go-ahead (Step 7). |

The message has no improvement-insight section, so `improvement_insight` still comes from the AI Champions notes (source 5 below).

Report the message's date range and posted time to Houston in the Step 2 summary. **If no matching message exists, or the newest one is older than 8 days, say so and fall back to the full scan below.** If the Webex call returns 401/403, tell Houston the token or scopes need re-auth (`cd C:\Users\hodowns\webex-mcp-oauth && npm run auth`) and fall back.

### Gap-fill sources (light use)

Use these to verify the message, add detail it lacks, and cover anything it missed in the last 7 days. Do not re-derive what the message already covers.

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

Using the data pulled in Step 1, **automatically determine** all five answers below. Do not ask Houston interactively. **Build each answer from the MyAgent message first (Step 1A), then fill gaps from meeting notes, emails, and transcripts.** Keep the message's facts (names, PIDs, dates, numbers) and do not drop items it lists. Where gap-fill sources contradict it, flag the conflict in the summary instead of silently choosing.

1. **Top priorities this week** — The 2–3 highest-impact things accomplished or moved forward, based on meeting notes, sent emails, and action items closed.
2. **Blockers or risks** — Open action items, canceled/deferred decisions, technical unknowns, or escalation signals found in notes.
3. **Customer highlights** — Notable wins, milestone events, or concerns surfaced in external-facing sessions or email threads.
4. **AI / innovation work** — Cockpit/Claude Code sessions led or attended, skills tested or shipped, estimated time savings from automations.
5. **Next week's focus** — Committed follow-ups and due-date action items from this week's meetings; carry forward any unresolved items from the previous check-in.

After inferring all five, present a brief **"Here's what I found"** summary (one line per answer) for visibility, then **continue straight to Step 3 without waiting for confirmation** (Houston's standing instruction, 2026-11-24: assume it's right and move on). Houston can correct afterward and the document and payload get regenerated.

**The only thing you still ask for is this question (it is the one place the skill may pause), unless Houston already gave it in his request:** *"Is there anything you need from your leader this week? (This is the portal's 'What do you need from your leader this week?' field.)"*
- The field is for asks of **Greg Hopper** (decisions, access, prioritization, escalation, support), not technical status. **Never auto-fill it from blockers or meeting notes.** Technical risks belong in the check-in doc's Blockers / Risks section only.
- Leave it blank until Houston answers. If he says nothing needs to be added (e.g. "no", "nothing", "none", "n/a"), use exactly `N/A`.

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

If any activity from this week doesn't fit cleanly under one of the five goals, flag it at the end of the document and suggest a new goal to capture it going forward.

---

## Step 5 — Save

Save the completed document as `YYYY-MM-DD-checkin.md` in `C:\Users\hodowns\OneDrive - Cisco\Desktop\Check Ins\`.

Confirm the filename and path after saving.

---

## Step 6 — Write JSON Payload for Playwright

Write `C:\Users\hodowns\Claude Projects VS CODE\my-performance-data\performance_data.json` (a gitignored folder inside the repo, so the governance hook allows the write) as a proper JSON file (not embedded in markdown). This is what `my_performance_checkin.py` reads first.

Synthesize each field from the conversation data, applying the hyper-productive narrative voice from Step 3:

- `weekly_accomplishments` — key accomplishments + customer highlights (Q1 + Q3); bullet symbols (•) with semicolons; **max 490 chars; no literal newlines**
- `improvement_insight` — the **strongest 1–2 items** found in transcripts that answer "What's one improvement you made last week—for yourself, your team, or your customers?" Pull from AI Champions calls first (files in `XXXX Innovation\AI Champions\Meeting Minutes\`), then F51119/F24373/C16507 meeting notes. Frame using executive language. **Max 490 chars; no newlines.** Examples of strong improvement items: "Drove GitHub Champions membership from 22→42 in one week through structured Champion-led outreach, accelerating PMO AI adoption timeline"; "Delivered live issue→PR→merge Claude demo at 50+ PMO session, establishing GitHub as the standard workflow for skill contributions"; "Identified Haiku model as drop-in replacement for routine tasks at 50% of Sonnet cost, reducing PMO token spend without quality degradation"
- `top_priorities_next_week` — Q5 priorities, numbered; **max 490 chars**
- `customer_impact_value` — one strong sentence from Q3
- `blockers_and_support_needed` — **the portal's "What do you need from your leader this week?" field. Use ONLY Houston's own answer to the Step 2 leader-support question** (max 490 chars). `"N/A"` if he indicated nothing to add; `""` if not yet asked (ask before writing the payload or filling the portal). Never derive it from blockers, risks, or notes. Put technical blockers in the check-in doc instead.
- `internal_leadership_contributions` — Q4, or `"No activity this week."` if skipped
- `new_priorities` — array of 4–6 priority objects extracted from transcripts/emails, each mapped to the best-fit FY27 goal. Use judgment on priority ranking.

For `new_goals`: extract 2–4 broad, organization-level strategic goals from PMO Leadership call notes, AI Champions sessions, Flight School notes, or all-hands emails found in Step 1. These are **Cisco-wide priorities, not personal goals** — executive strategy language, broad scope. If no explicit all-hands source is found, derive from the themes in recent AI Champions and PMO leadership sessions. Include only goals not already in the system.

**FY27 Strategy-aligned reference goals (prefer these; they mirror Grow / Defend / Transform):**
- *Grow Durable PS Revenue Through Customer Lifetime Value* — Align PS delivery and account engagement to LTV so engagements drive durable, high-quality growth and follow-on opportunity rather than one-time volume
- *Defend the Installed Base by Leading with Cisco IQ* — Embed IQ/IQ Link, 30-day PSIRT action, and premium-mix thinking into every customer governance touchpoint so no account is dropped
- *Transform Delivery into an Agentic System* — Redesign core PS program-management workflows end-to-end (playbook/as-code, humans at judgment points) and measure by customer outcomes, not hours

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

Confirm: `Payload written → C:\Users\hodowns\Claude Projects VS CODE\my-performance-data\performance_data.json`

---

## Step 7 — Fill the Portal (draft only)

Use the route chosen in Step 0. **Standing instruction from Houston (2026-11-24): do not wait for a draft review. After the draft is filled and read back correctly (Step 7A item 6), submit it** with `myperformance_submit_weekly_checkin` (or `--submit` on the 7B route). Submission is irreversible, so only submit when the read-back matches the payload (status DRAFT, reflect text, support text, every priority with its goal and week); if anything is off, fix it first, and if you cannot, stop and tell Houston. Never submit on the strength of the MyAgent message's "Team Space Execution Text" line alone.

### Step 7A — MyPerformance connector (preferred; verified 2026-11-06)

All calls use `week_start` = the check-in week's start date shown by `myperformance_get_weekly_checkin` (the Friday, e.g. `2026-11-06`), except priorities (below). Read the payload from `performance_data.json`.

1. `myperformance_start_weekly_checkin(week_start)` — creates the DRAFT.
2. `myperformance_update_checkin_reflection(week_start, leader_connect="CHAT_EMAIL", work_shape=<improvement_insight>)` — `work_shape` is the improvement answer.
3. `myperformance_update_checkin_alignment(week_start, leader_support_text=<blockers_and_support_needed>)` — this is the "What do you need from your leader this week?" field. Only Houston's own answer or `N/A` (see Step 2). If the payload value is empty, ask him first; do not fill it yourself.
4. `myperformance_list_goals()` — the portal's goals are the real link targets and **differ from the 5 personal goals above**. Map each priority to the best-fit active goal by title and pass its `id` in `goal_ids`. Do not guess IDs.
5. `myperformance_add_priority(title, goal_ids, week_start=<NEXT Monday>)` for each of the 4–6 `new_priorities`. **The priority `week_start` is the week the priority is for, so use next week's Monday (check-in Friday + 3 days, e.g. `2026-11-09`), not the check-in date.** Using the check-in date files it a week early. If one is filed wrong, `myperformance_cancel_priority` it and re-add. Batch these in one parallel call.
6. Read back with `myperformance_get_weekly_checkin(week_start)` and confirm status `DRAFT`, the reflect text, the support text, and every priority with its goal and `2026-11-09`-style week.

**Known gap:** the connector has no tool to resolve last week's OPEN priorities. Do not cancel them; cancelling is not completing. Resolve them in the UI with Playwright's own browser (verified 2026-11-06):
1. `browser_navigate` to `https://teamspace.cisco.com`; if it lands on Duo, ask Houston to sign in in that window, then continue.
2. Click **Check-ins**, then **Resume Weekly check-in** (the connector-created DRAFT shows as "Resume").
3. Step 1 lists last week's priorities, each with **Mark as complete** and **Roll over**. Click **Mark as complete** per item. The button flips to "Mark as incomplete" and does not disappear, so do not loop by index; click each by its snapshot ref and confirm all show "Mark as incomplete". Ask Houston which to complete vs roll over if it is not obvious.
4. **Save & continue** to Step 2 (confirm the priorities and support text match the draft), **Save & continue** to Step 3, then **Submit** only with Houston's explicit go-ahead.
5. Verify with `myperformance_get_weekly_checkin` for both weeks: this week `SUBMITTED`, last week's priorities `COMPLETED`.

Then tell Houston the draft is ready to review in TeamSpace.

### Step 7B — Playwright script (fallback, browser route)

Use this when the connector is unavailable, or to resolve last week's priorities. Confirm the browser route from Step 0 is running, then offer:

> "Payload written. Ready to auto-fill TeamSpace via Playwright?"

If Houston confirms, run it in **draft mode** (the default). Pass the payload path explicitly:

```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\my_performance_checkin.py" "C:\Users\hodowns\Claude Projects VS CODE\my-performance-data\performance_data.json"
```

Draft mode fills and saves Steps 1 and 2, then stops at Step 3 **without clicking Submit**. Tell Houston to review the browser.

**Never add `--submit` on your own.** Only re-run with `--submit` after Houston has reviewed the draft and explicitly said to submit. Submission is irreversible, and the script refuses if the page looks already submitted.

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
1. Draft mode (default): stops here, Submit NOT clicked
2. With `--submit` (only after Houston's explicit go-ahead): checks the page isn't already submitted, then clicks **Submit** (full submission)

Watch the output for `✓`/`✗` per field. Screenshots saved to the script directory. If CDP fails, re-run Step 0.
