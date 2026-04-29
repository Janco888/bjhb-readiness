---
description: Build the weekly Job Readiness dashboard from SAP exports
argument-hint: "[--debug]"
allowed-tools:
  - Bash
  - Read
  - Glob
---

# /readiness — Build the Weekly Job Readiness Dashboard

You are executing the weekly readiness pipeline. Follow these steps exactly and do not skip any of them.

## Step 1: Check that input files exist

Check that the two required SAP exports are present in the `inputs/` folder, and note whether the optional PO file is present:

```bash
ls -la inputs/coois_components.xlsx inputs/mb52_stock.xlsx inputs/y00_zmpo.xlsx 2>&1
```

If `coois_components.xlsx` or `mb52_stock.xlsx` is missing:
- Tell the user which file is missing
- Remind them to export it from SAP (refer to `docs/SAP_EXPORT_CHECKLIST.md`)
- **STOP** — do not continue

If `y00_zmpo.xlsx` is missing:
- Note it to the user: "PO file not found — the dashboard will run without PO coverage columns. Export ZMPO and drop it in inputs/ to enable PO visibility."
- Continue — it is optional

## Step 2: Validate the inputs

Run the validation script to catch format issues early:

```bash
python scripts/validate_inputs.py
```

If validation fails:
- Read the error message carefully
- Report it to the user in plain language
- Suggest the most likely cause (usually a column layout issue or wrong date range in SAP)
- **STOP** — do not continue until inputs are fixed

## Step 3: Build the dashboard

Run the main build script:

```bash
python scripts/build_readiness.py $ARGUMENTS
```

Watch the output for:
- The READY / PARTIAL / NOT READY counts at the end
- Any warnings or errors in the log

## Step 4: Sanity-check the results

Compare this week's numbers against what's reasonable:

- **If READY = 0**: something is wrong. Either stock is genuinely empty, or the date filter caught nothing. Flag this to the user.
- **If NOT READY > 90% of total**: probably fine if the plant is busy, but check with user whether this matches reality.
- **If READY > 80% of total**: unusual. Could be valid (quiet week), or the input files might be from different dates. Flag to user.
- **If total jobs < 10 or > 500**: outside expected range, ask user to confirm the SAP export filters.

## Step 5: Report to the user

Give the user a concise summary:

```
✓ Dashboard built: outputs/Job_Readiness_Board_YYYY-MM-DD.xlsx

This week:
  🟢 READY:     X jobs can start now
  🟡 PARTIAL:   X jobs have small shortages
  🔴 NOT READY: X jobs blocked by missing parts

Top priority: <brief comment on what matters most this week>

Input files archived to inputs/archive/
```

Then suggest one concrete next action based on the numbers — e.g. "Recommend reviewing the top 5 NOT READY jobs with the foreman" or "The 14 READY jobs can be released to the workshop today."

## Important Rules

- **Never edit `scripts/build_readiness.py` during a /readiness run.** If the script is broken, tell the user and stop.
- **Never modify files in `inputs/archive/`**.
- **If the Python script produces unexpected output**, do not attempt to patch the code. Report the exact error to the user and stop.
- **If the user asks for a different view** (e.g. "can you show me just the red ones?"), offer to filter the Excel file but don't rebuild it.
