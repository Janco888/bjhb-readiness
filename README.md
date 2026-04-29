# BJHB Job Readiness Board

**One question, answered every Monday: can this job start today, or will it run short?**

## What This Does

Takes two SAP exports, simulates what would happen if Stores tried to pick every released production order against current warehouse stock in start-date order, and outputs a colour-coded Excel dashboard showing which jobs are ready to release to the workshop.

The important insight: **stock is shared across jobs**. If three jobs each need 40 units and there's only 100 in stock, one of them will be short — even though standard reports show 100 available for every job. This dashboard finds those hidden shortages.

## Quick Start

### First time setup

```bash
# 1. Clone or copy this folder somewhere on your Mac
cd ~/Projects  # or wherever
# (assume the folder is already here: bjhb_readiness)

# 2. Create a virtual environment and install dependencies
cd bjhb_readiness
python3 -m venv .venv
source .venv/bin/activate
pip install -e .

# 3. Verify installation
python scripts/validate_inputs.py --help
```

### Weekly Monday routine (5 minutes)

1. **Export from SAP** (see `docs/SAP_EXPORT_CHECKLIST.md` for exact clicks):
   - COOIS Components view → save as `coois_components.xlsx`
   - MB52 stock → save as `mb52_stock.xlsx`

2. **Drop both files into `inputs/`** — overwrite the previous ones

3. **Run it** — one of these:
   ```bash
   # Option A: from Claude Code
   /readiness

   # Option B: from terminal
   source .venv/bin/activate
   python scripts/build_readiness.py
   ```

4. **Open the dashboard** — it's in `outputs/` with today's date in the filename.

## What You'll See

The dashboard has 4 sheets:

| Sheet | Purpose |
|---|---|
| `READINESS_BOARD` | Main view — one row per job, traffic-light status |
| `COMPONENT_DETAIL` | Why each job got its colour — filter by MO for a specific job |
| `STOCK_LEDGER` | The receipts — proves the math when someone questions a result |
| `HOW_IT_WORKS` | Explanation embedded in the file for anyone else who opens it |

**Red (NOT READY)** = don't release to floor, escalate
**Yellow (PARTIAL)** = small shortage, planner decides
**Green (READY)** = safe to release

## Troubleshooting

| Problem | Likely cause | Fix |
|---|---|---|
| "File not found: inputs/coois_components.xlsx" | File named wrong or missing | Check filename is exactly `coois_components.xlsx` |
| "Missing required column: X" | SAP layout changed | Run with `--debug` flag to see what columns were found |
| "0 rows after filtering" | Wrong date range in SAP export | Check Start Date filter in COOIS — should cover today + 90 days |
| All jobs show NOT READY | MB52 export is empty or wrong plant | Reopen MB52 in SAP, verify plant 502 |
| All jobs show READY | COOIS export missing `Quantity withdrawn` column | Re-export COOIS with full component layout |

If the script fails, the error message will usually point to the problem. If it doesn't, run:

```bash
python scripts/validate_inputs.py --verbose
```

This prints exactly what it found in each file.

## Extending This

Want to add something? Some rules:

- **Don't modify `build_readiness.py`** for one-off queries — write a new script
- **Do add new scripts** for new views (PO tracker, PR tracker, planned-order forecast)
- **Reuse the input files** — don't re-export SAP for each new script
- **Keep each script single-purpose** — easier to debug and maintain

Ideas for future scripts (in order of usefulness):

1. `scripts/po_impact.py` — re-runs the readiness check assuming pending POs arrive on time
2. `scripts/planned_forward.py` — projects readiness 4 weeks out using planned orders
3. `scripts/supplier_scorecard.py` — which suppliers cause the most shortages

## Questions?

Check `CLAUDE.md` for the technical context, or ask Claude Code directly — it has the context loaded.
