# Claude-How-To
Automated workflows for SLED/USPS

## Business Problem
A Program Manager running a portfolio of SLED/USPS customer projects repeats the same administrative work across several systems: submitting SDC resource requests, keeping project cost and revenue forecasts compliant, writing and filing meeting minutes and recaps, chasing action items, booking meetings, building weekly status and financial reports, closing finished projects, and completing the weekly performance check-in. Today that is done by hand across SDC, Oracle Projects, Outlook, Webex and SharePoint.

## Why
The tooling these tasks depend on is changing: Oracle Projects closes on 2026-10-25 and SDC now handles forecasting, which retired the earlier Oracle cost-forecast script and required a replacement. The workload also scales with the number of active projects, so manual handling grows with the portfolio.

## How
Each workflow is a Python script, a Claude Code skill or agent definition, or both, started by a trigger phrase documented in `CLAUDE.md` (for example "Cost Revenue Compliance" or "meeting minutes"). Browser-based systems such as SDC and the Teamspace portal are driven with Playwright; Outlook, Webex and SDC data are read through their integrations. Outputs are Outlook drafts, HTML reports, and filed documents, so the work can be reviewed before it goes out.

## Not This
- **Not a replacement for SDC, Outlook, Webex, or Oracle Projects.** It reads from and drives those systems; they stay the systems of record.
- **Not fully autonomous.** A person starts each workflow, usually with a trigger phrase documented in `CLAUDE.md`.
- **Not (yet) a supported team product.** It was built around one Program Manager's priorities and is designed with broader use cases in mind, with the aim of becoming a scalable solution for all PMs over time. Until then, expect defaults specific to that PM's portfolio, such as org names and hardcoded fields.

## Impact
Unquantified. Time savings have been estimated informally but not measured, so no figure is claimed here.

## Details

| Facet | Value |
|---|---|
| **Status** | `active` |
| **Owner** | hodowns@cisco.com |
| **Related** | *none* |

Note: the My Performance weekly check-in workflow (`.claude/commands/my-performance.md` and `my_performance_checkin.py`) is still `experimental`.
