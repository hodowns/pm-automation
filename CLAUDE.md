# Claude Projects VS CODE — Workflow Triggers

## Request a Resource in SDC
**Trigger:** "Request a Resource" / "Resource Request" / "Submit RR"

Before running, ask if not specified: **project-level or task-level?**

```
py -3 -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\submit_resource_request.py" \
  --pid [PID] --role [Role] --work-description "[desc]" \
  --start-date YYYY-MM-DD --end-date YYYY-MM-DD --hours [N] \
  --architecture [arch] --technology "[tech]" --grade [8|10|11|12] \
  --milestone "[Milestone Name]" \
  [--suggested-resource "Full Name"] \
  [--task-level]   # add for task-level; omit for project-level
```

Add `--dry-run` to stop at the review page without submitting.
Work Description is auto-generated from `Scoping/` folder when possible.

**Hardcoded defaults (change in script if org differs):**
- Resource Allocation Reason: `Assign new resource`
- Onsite Requirement: `No`
- Citizenship Requirement: `No`
- Time Zone Requirement: `No`
- Security Clearance: `No`
- Start Date Flexibility: `Ontime`
- Scheduling Mode (task-level): `Scheduled`

## Cost Forecast
**Trigger:** "Cost Forecast"

Run the Oracle Projects bulk cost forecast submission:
```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\op_cost_forecast.py"
```
Opens Edge, waits for SSO login, then automatically: Select_All_Shown → Submit → Yes.
After running, verify Last Submitted resets to 0 in Oracle Projects.

## Cost Revenue Compliance
**Trigger:** "Cost Revenue Compliance" / "Cost and Revenue Compliance"

Follow the workflow in `pm-agent-system/agents/cost-revenue-compliance-agent.md`.
Replaces Oracle Projects cost forecast (op_cost_forecast.py retired 2026-10-25).
Navigates to each active SDC project and executes per-project compliance steps.

**Step 1 — Fetch active PIDs:**
Call `sfdc_get_my_projects(phase="Active")` and extract all PIDs.

**Step 2 — Run (discovery mode first):**
```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\cost_revenue_compliance.py" [PID1] [PID2] ...
```
Output: `Desktop\SDC Compliance Discovery\discovery_report.html`

**Step 3 — Run automation mode (after selectors confirmed):**
```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\cost_revenue_compliance.py" --run [PID1] [PID2] ...
```

## Secretary Agent
**Trigger:** "Secretary Agent"

Run the weekly action item report:
```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\secretary_agent.py"
```
Scans Webex recap emails, meeting notes, and Outlook for Houston's action items.
Output: `Desktop\Secretary Agent\action-items.html`

## Weekly Summary
**Trigger:** "Weekly Summary [PID]"

```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\weekly_summary.py" [PID]
```

## Profile Report
**Trigger:** "Profile report [PID]"

```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\profile_report.py" [PID]
```

## End of Day Recap
**Trigger:** "send end of day recap" / "EOD recap"

```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\End of Day Recap\send_eod_recap.py"
```
Add `--yesterday` flag for previous day's meetings.

## Simple Recap
**Trigger:** "simple recap for [meeting name]"

```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\Simple Meeting Recap\send_recap.py" "[meeting name keywords]"
```

## Schedule Booker
**Trigger:** "Schedule Booker"

Follow the workflow in `pm-agent-system/agents/schedule-booker-agent.md`.
Scans inbox for all `Re: Schedule Request:` replies, detects confirmed option, checks if already on calendar, and opens a draft recurring or one-time invite for review.
```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\schedule_booker.py"
```

## Schedule Requester
**Trigger:** "Schedule Requester"

Follow the workflow in `pm-agent-system/agents/schedule-requester-agent.md`.
Prompts for meeting name, PID, customer recipients, Cisco SME(s), duration (default 45 min — remind if not stated), and one-time vs recurring.
Subject format: `Schedule Request: [Meeting Name] [EXTERNAL] [PID]` — required for future reply-detection workflow.

## Nudge Agent
**Trigger:** "Nudge"

Run the action item nudge drafter:
```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\nudge_agent.py"
```
Scans meeting notes (.md) and Webex recap emails for other people's open action items.
Opens one Outlook draft per owner — concise table of their outstanding items.
Add `--days 30` to widen the lookback window beyond the default (since last run).

## Training Check
**Trigger:** "what trainings are due?"

```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\training_check.py"
```
Scans Outlook inbox (last 60 days) for training obligations. Ignores trainings Houston is delivering to customers. Cross-references completion emails so finished trainings are not listed as due.
Output: Outlook draft to hodowns@cisco.com — table of Training Name | Link | Due Date.
Also runs automatically every 50 days via Windows Task Scheduler.

## Close PID
**Trigger:** "Close [PID]" or "Close PID [PID]"

```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\Close a PID Gen 3\close_project.py" [PID]
```
Opens Chrome (SSO auto-login after first run), searches SDC for the PID, adds Lessons Learned, sets Oracle Project Status to "Delivery Close", runs the Project Closure Checklist, checks all manual tasks, and submits to PCT.
Multiple PIDs supported: `close_project.py F31938 F26710`

## Fin Report
**Trigger:** "Fin Report"

Multi-step workflow — do NOT run the script directly without args (it uses input() which breaks in Claude):

**Step 1 — Find CSVs.** List recent CSVs in Downloads:
```
python -u -X utf8 -c "
import glob, os
d = os.path.expanduser('~/Downloads')
csvs = sorted(glob.glob(d + '/*.csv'), key=os.path.getmtime, reverse=True)[:8]
for i, f in enumerate(csvs): print(i, os.path.basename(f))
"
```
Present the list and use AskUserQuestion to confirm which file is GPPR and which is Expenditure.

**Step 2 — Build Excel + open DocCentral.**
```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\fin_report.py" --gppr "<GPPR_PATH>" --exp "<EXP_PATH>"
```
Script builds the Excel file, opens DocCentral in the browser, then exits cleanly.
Tell the user: "DocCentral is open. Upload the file, copy the direct link, and paste it here."

**Step 3 — Create Outlook draft.** Once the user pastes the DocCentral link:
```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\fin_report.py" --draft-only --link "<URL>"
```

**Unattended / Task Scheduler mode** (future — pending Tableau access):
```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\fin_report.py" --auto
```

## Meeting Minutes
**Trigger:** "meeting minutes"

Default meeting: **Weekly CoH LCS Status Review [EXTERNAL] F51119** (use this when no meeting name is specified).

Follow the minutes-minion workflow in `pm-agent-system/agents/minutes-minion.md`.

## Agenda Generator
**Trigger:** "Create Agenda: [meeting name]" / "agenda for [meeting]" / "build agenda for [meeting]"

Follow the agenda-generator workflow in `pm-agent-system/agents/agenda-generator.md`.
Scans recent emails, last meeting notes, and open action items; builds a categorized agenda filtered by attendee specialization; opens the next calendar occurrence pre-populated for review.

## Status Report
**Trigger:** "Status Report [PID]" / "Status Report"

Follow the workflow in `pm-agent-system/agents/status-report-agent.md`.
Scans Outlook and Webex for account email traffic, generates an interactive HTML status report with all technology track slides.
```
py -3 -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\status_report.py" [PID]
```
Default PID is F51119 (City of Houston LCS). Add `--since YYYY-MM-DD` to override scan window. Add `--no-scan` to skip email/Webex and generate from previous state only.
Output: `OneDrive\Desktop\Status Reports\F51119_status_YYYY-MM-DD.html` (opens in browser automatically).

**GitHub Pages collaboration (DM can view + edit the same live report):**
- First time only: `status_report.py F51119 --no-scan --push --deploy-pages --gh-org cx-usps-auto --gh-repo status-reports`
  (also run `--gh-save-token --gh-token ghp_xxx` once to store the PAT)
- Every week with `--push`: also commits `data/F51119_state.json` to the repo
- DM opens `https://wwwin-github.cisco.com/pages/cx-usps-auto/status-report-presentation/` and enters their own PAT once
- Edits auto-save to GitHub; other person's browser polls every 30s and shows a refresh banner
