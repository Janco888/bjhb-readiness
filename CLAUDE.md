# BJHB Job Readiness Board — Project Context for Claude Code

## Purpose

One question, answered weekly: **"For each released production order, can it start this week — yes or no?"**

This project replaces ad-hoc Excel work with a repeatable pipeline. Every Monday the planner drops three SAP exports into `inputs/`, runs `/readiness`, and gets a dashboard showing which jobs can run and which will be short on parts.

## The Core Logic — Virtual Pick Simulation

The key insight is that **MB52 stock is shared**. If MB52 shows 100 units of a material, and three jobs each need 40, the dashboard must recognise the third job will be short — even though standard reports show "100 in stock" for all three.

The algorithm:
1. Load all released MOs from COOIS, sorted by Start Date ascending
2. Load current unrestricted stock from MB52, and all the Purchase order information from y00_zmpo
3. For each MO in start-date order, simulate picking each component against the remaining stock pool, and what is to be delived based on the purchase order created
4. Deduct what gets picked. Later MOs inherit the depleted pool.
5. Classify each MO as READY / PARTIAL / NOT READY based on component outcomes, and POs that is created

This mirrors how Stores would actually pick if they processed all jobs in start-date order on a Monday morning.

## Project Structure

```
.
├── CLAUDE.md                    # This file
├── README.md                    # Human-readable instructions
├── pyproject.toml               # Python dependencies
├── .claude/
│   ├── commands/
│   │   └── readiness.md         # /readiness slash command
│   └── agents/
│       └── sap-validator.md     # Subagent for SAP data quality
├── scripts/
│   ├── build_readiness.py       # Main dashboard builder
│   └── validate_inputs.py       # Pre-flight SAP file check
├── inputs/                      # Drop SAP exports here
│   ├── coois_components.xlsx    # Expected filename
│   ├── mb52_stock.xlsx          # Expected filename
│   ├── y00_zmp0.xlsx            # Expected filename
│   └── archive/                 # Old inputs auto-archived after each run
├── outputs/                     # Generated dashboards (timestamped)
└── docs/
    └── SAP_EXPORT_CHECKLIST.md  # Printable SAP clicking guide
```

## Expected Input Files

Drop these two files into `inputs/` before running:

### `coois_components.xlsx`
- Source: SAP Transaction **COOIS**, Component view
- Required columns: `Order`, `Material`, `Material Description`, `Requirement Quantity`, `Quantity withdrawn`, `Procurement Type`, `Header Material Text`, `Header Basic Start Date`, `Header Basic Finish Date`, `Header SD order`
- Filter: Plant 502, Status REL (Released), exclude DLV and TECO
- Date range: today - 30 days → today + 90 days

### `mb52_stock.xlsx`
- Source: SAP Transaction **MB52**
- Hierarchical format (SAP standard): material row, then location+qty row per storage location
- Key column: Unrestricted stock (column index 5 in the raw export, zero-indexed)
- Filter: Plant 502, all storage locations

### `y00_zmp0.xlsx`
- Source: SAP Transaction **y00_zmp0.xlsx**
- Hierarchical format (SAP standard)
- Filter: Plant 502, all storage locations

## How to Run

```bash
# Manual
python scripts/validate_inputs.py && python scripts/build_readiness.py

# Via Claude Code
/readiness
```

## Conventions

- **Never modify files in `inputs/archive/`** — historical record
- **Outputs are timestamped** — never overwritten, so history is preserved
- **Stock parsing is fragile** — MB52 format varies between SAP installations; if extraction fails, check the first 10 rows of the file manually
- **Type E (internal) parts** are treated the same as Type F for the simulation, but they're flagged in the dashboard — workshop makes these, so "short" means the workshop must build them before the job runs

## Known Limitations (don't build around these without discussion)

- Simulation uses **Start Date** to rank picks. Real picking may differ if planner manually prioritises.
- Stock = **unrestricted only**. QI, Transit, and Blocked stock are ignored (intentional — they're not pickable today).
- The simulation assumes **no partial picks**. Real-world picking may split allocations across days.

## When Something Goes Wrong

1. Run `validate_inputs.py` first — catches 80% of issues
2. Check the console output from `build_readiness.py` — it prints row counts and status distributions
3. If the READY count looks wildly wrong (e.g., 0 READY or 100% READY), the input data is probably not what we expect
4. If XML validation fails on the output, the openpyxl version may have introduced a conflict — check `pyproject.toml` for pinned versions

## Not In Scope (Yet)

- Forward demand from planned orders (MD04)
- Lead time-based predictive warnings
- Multi-plant support (hardcoded to 502 currently)

If someone asks for any of the above, suggest building it as a separate script in `scripts/` rather than extending `build_readiness.py`.
