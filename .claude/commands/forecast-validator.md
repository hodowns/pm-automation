---
description: Forecast Validator — trigger "validate forecast". For every active SDC project, open the Forecasting tab and click Validate Forecast
---

# Forecast Validator

Trigger: **"validate forecast"**.
Replaces the retired Oracle Projects cost forecast (`op_cost_forecast.py`, OP closed 2026-10-25).

## What it does, per project

1. Open the project record in SDC.
2. Go to the **Forecasting** tab (it is named "Forecasting", not "Forecast").
3. Click the blue **Validate Forecast** button.
4. Confirm the green banner reads "Forecast was reviewed by Houston Downs on <today>."

## How to run it (Playwright MCP — no script needed)

1. Call `sfdc_get_my_projects(phase="Active")`. Keep each project's `pid` and `sfid`.
2. With the Playwright MCP, go to `https://csone.lightning.force.com/lightning/r/pse__Proj__c/<sfid>/view`.
   If it lands on a Duo page, ask Houston to approve the Duo prompt, then wait ~20 seconds.
3. Run the loop with `browser_run_code_unsafe`, 4-6 projects per call (a full run in one call takes too long). Per project:
   - `page.goto(url, {waitUntil:'domcontentloaded'})`
   - click `getByRole('tab', {name:'Forecasting', exact:true}).locator('visible=true').first()`
   - if `getByRole('button', {name:'Validate Forecast'}).locator('visible=true').first()` appears, click it
   - wait ~5s, then read the visible text matching `/reviewed by/i` as the confirmation
   - wrap each project in try/catch so one failure doesn't stop the batch
4. Report a table: PID | Result. Results are: validated today / already validated today / no button / no Forecasting tab.

## Known cases

- **No Validate Forecast button, banner already dated today:** already validated earlier today. Fine.
- **No Validate Forecast button, banner dated an earlier day** (seen on F43673, reviewed Oct 26): SDC did not offer the button. Flag it to Houston; don't guess why.
- **No Forecasting tab** (seen on F59190, an Adjunct project with a $0.01 budget): nothing to validate. Flag it.
- Projects past their end date can still be Active in SDC (e.g. F33347 ended 2026-10-20). Mention them so Houston can close them.

## Notes

- This is a write action in SDC. Houston asked for it directly, so run it without a dry run, but always report exactly what was and wasn't clicked.
- The old script `cost_revenue_compliance.py` is read-only discovery. Its `--run` mode is still a stub; it also looked for a tab called "Forecast", which is why discovery found nothing there.
- Related: `pm-agent-system/agents/cost-revenue-compliance-agent.md`.
