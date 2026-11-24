---
description: Fix a Reverse-plans PR that conflicts on data/all.json. Usage: /auto-resolve 193
---

# Auto-resolve a Reverse-plans PR conflict

Argument: `$ARGUMENTS` is the PR number (e.g. `193` or `PR#193`; strip any `PR#` or `#`).
If no number is given, list open conflicting PRs and ask which to fix:
`GH_HOST=wwwin-github.cisco.com gh pr list --repo cx-usps-auto/Reverse-plans --state open --json number,title,mergeable`

Repo: `cx-usps-auto/Reverse-plans` on `wwwin-github.cisco.com`. Local clone: `C:\Users\hodowns\rp-work\Reverse-plans`.

## Why this exists
Every plan PR rewrites `data/all.json`, so when two are open the second conflicts. Accept-current/incoming/both in the web editor loses plans. `scripts/fix_pr.sh` merges main into the PR branch, rebuilds `all.json` from the plan files, and pushes.

## Steps

1. **Check the PR** is OPEN and that only plan files + `data/all.json` changed:
   `gh pr view N --repo cx-usps-auto/Reverse-plans --json state,mergeable,files`
   If it touches code (e.g. `index.html`), STOP and tell the user — the script refuses these and a human/author must rebase.

2. **Order check:** if other plan PRs conflicting on `all.json` are open, fix them one at a time; each must be merged before fixing the next. Tell the user the order.

3. **Make sure the clone exists and is current** (clone it if missing):
   ```
   cd /c/Users/hodowns/rp-work/Reverse-plans && git fetch -q origin main && git checkout -q origin/main
   ```
   (Clone: `git clone https://wwwin-github.cisco.com/cx-usps-auto/Reverse-plans.git` into `/c/Users/hodowns/rp-work`.)

4. **Run the script.** The PATH prefix is required: Git Bash's `python3` is a Windows Store stub; the real Python is in AppData.
   ```
   cd /c/Users/hodowns/rp-work/Reverse-plans && PATH="/c/Users/hodowns/AppData/Local/Python/bin:$PATH" GH_HOST=wwwin-github.cisco.com bash scripts/fix_pr.sh N
   ```
   Success ends with "Pushed. PR #N should be mergeable in a few seconds." The "CONFLICT ... data/all.json" line before it is expected.

5. **Verify** after ~20-30s:
   `GH_HOST=wwwin-github.cisco.com gh pr view N --repo cx-usps-auto/Reverse-plans --json state,mergeable,mergeStateStatus`
   Expect `MERGEABLE` / `CLEAN`. `UNKNOWN` means GitHub is still computing — wait and recheck.

6. **Report** the plan count from the script output and the PR link. The user clicks Merge (merging is blocked for Claude in this setup; do not try to work around it). Then offer to run the next PR.

## Failure cases
- "These files conflict too" → non-`all.json` conflict; stop, report the files.
- "fork does not allow maintainer edits" → ask the PR author to tick "Allow edits from maintainers".
- "Python was not found" → the PATH prefix was dropped; rerun with it.
- Do not paste commands for the user into Git Bash (pasting adds stray `~`); do the work from here.
