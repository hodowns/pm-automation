# Weekly Check-In Skill — Setup Guide

Set up a personalized `/weekly-checkin` skill in Claude Code that automatically drafts your weekly check-in from your email, calendar, and project data — aligned to your goals and Cisco US PS FY26 priorities.

**Setup takes about 10 minutes. Running it each week takes about 5.**

---

## Before You Start

Make sure you have:

- Claude Code installed and open
- Outlook connected to Claude Code (MCP)
- Webex connected to Claude Code (MCP) — optional but recommended

If you are not sure whether Outlook or Webex are connected, type `! gh` in Claude Code or ask your Claude Code admin.

---

## How to Set Up

Copy everything in the box below and paste it into Claude Code as your first message. Claude will guide you through the rest — no technical knowledge needed.

---

```
I want to set up a personalized weekly check-in skill. Please guide me through the setup one step at a time. Wait for my answer before moving to the next question.

Start by briefly explaining what this skill will do, then ask me the following questions in order:

1. DATA SOURCES
Ask me which of these I have access to and want Claude to pull from each week:
- Outlook email (inbox and sent items)
- Outlook calendar
- Webex team spaces (ask me to share the space name or link for each one I want included)
- Project documents or status reports saved on my computer (ask for the folder path)

2. MY MANAGER
Ask: "Who is your direct manager? First and last name."

3. MY PROJECTS
Ask: "What are your active projects or customer accounts right now? Just list the names."

4. MY GOALS
Explain that goals come in two types:
- Business Outcome: results I'm accountable for delivering
- Professional Development: growth that expands my impact

Tell me goals should be SMART (specific, measurable, attainable, relevant, timebound — 1 to 2 quarter horizon) and under 500 characters each.

Ask if I have existing goals to use, or if I want help drafting them. If I want help, ask about my role and responsibilities and suggest 2-4 goals for me to review.

5. FY26 ALIGNMENT
Once my goals are confirmed, map each one to the Cisco US PS FY26 priority it best supports and show me the mapping:
- Drive Durable Growth
- Simplify for Customer Outcomes
- Leverage AI to Enhance Work

Ask me to confirm the mapping looks right before moving on.

6. OUTPUT LOCATION
Ask where I want my check-in files saved. Suggest creating a check-ins/ folder in my current working directory and offer to set it up.

Once I have answered all six questions, build my skill:
- Create the check-ins/ folder
- Create a /weekly-checkin command at .claude/commands/weekly-checkin.md

The skill should pull from my confirmed data sources, auto-detect how I connected with my manager each week, ask me five focused questions, enforce a 500-character limit on free-text answers, tag each priority to one of my goals, flag anything that does not fit an existing goal and suggest a new one, and save the output as YYYY-MM-DD-checkin.md in my check-ins folder.

After building, confirm what was created and tell me how to run it.
```

---

## What Gets Built

After setup, Claude creates:

- **`check-ins/`** — folder where your weekly documents are saved
- **`.claude/commands/weekly-checkin.md`** — your personal skill file

Run it any time by typing `/weekly-checkin` in Claude Code.

---

## Tips

- The more data sources you connect, the less you have to type each week
- Your check-in gets smarter over time — Claude reads previous check-ins to track which priorities are complete vs. rolling over
- If a weekly priority does not fit one of your goals, Claude will flag it and suggest a new goal rather than forcing a bad fit
- You can edit `.claude/commands/weekly-checkin.md` at any time to update your goals, add new data sources, or change the output format

---

*Created by Ruthie Brown, Customer Project Specialist, Cisco US PS.*
*Questions? Find me on Webex.*
