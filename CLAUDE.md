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
│   ├── y00_zmpo.xlsx            # Expected filename (note: letter 'o', not zero)
│   ├── me5a_prs.xlsx            # Optional: ME5A purchase requisitions
│   └── archive/                 # Old inputs moved here after each run
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

### `y00_zmpo.xlsx`
- Source: SAP Transaction **y00_zmpo** (note: letter 'o' at end, not zero — common typo)
- Flat table format with columns: `Material`, `Purchasing Document`, `PO-Quantity`, `GR-Quantity`, `Delivery Date`, `Name`
- Filter: Plant 502, open purchase orders only

### `me5a_prs.xlsx` (optional)
- Source: SAP Transaction **ME5A**, Purchase Requisitions
- When present, annotates short components with PR delivery dates
- Enables the "Supply Risk" column to consider both POs and PRs

## How to Run

```bash
# Manual
python scripts/validate_inputs.py && python scripts/build_readiness.py

# Via Claude Code
/readiness
```

## Conventions

- **Never modify files in `inputs/archive/`** — historical record
- **Archive moves, not copies** — files are moved out of `inputs/` after each run so stale files never carry over to the next week
- **Outputs are timestamped** — never overwritten, so history is preserved
- **Stock parsing is fragile** — MB52 format varies between SAP installations; if extraction fails, check the first 10 rows of the file manually
- **Type E (internal) parts** do NOT affect job readiness colour — they're tracked in the "Internal Short" column for workshop awareness only; readiness is scored on purchased/external parts only
- **Supply Risk column** shows PO IN TIME / PO LATE / NO PO for each short job — only populated when `y00_zmpo.xlsx` is present

## Streamlit Web UI (Alternative to CLI)

A full web interface is available for interactive use:

```bash
streamlit run app.py
```

The web UI adds: file upload, interactive job filtering, Plotly charts, production load by work centre, delay risk analysis, and shortage report download. It also supports the optional `me5a_prs.xlsx` and COOIS Operations files for richer analysis. The CLI (`/readiness`) produces the same core Excel dashboard without the interactive layer.

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

## Known Issues Found & Fixed (2026-05-14 multi-agent audit)

These were identified by a 4-agent code review and fixed in the same session. Record kept here so future Claude sessions don't re-introduce them.

### Bugs fixed in `build_readiness.py`
- **Archive used `shutil.copy2` instead of `shutil.move`** — inputs were never removed after archiving, so old files accumulated in `inputs/` and could pollute future runs. Fixed to use `shutil.move`.
- **`annotate_with_prs()` was never called in `main()`** — PR (ME5A) data was loaded in the Streamlit app only; CLI completely ignored it. Fixed: `main()` now looks for `me5a_prs.xlsx` and calls `annotate_with_prs()`.
- **Float precision accumulated in stock simulation** — after each pick, `remaining[mat]` was not rounded, causing tiny float errors (e.g., 0.00000001 treated as available stock). Fixed: round to 4 decimal places after each deduction.
- **`Days_to_Start` didn't normalize `Start_Date`** — if Start_Date had a time component, comparison against midnight `today` could be off by 1 day. Fixed: `.dt.normalize()` applied before subtraction.
- **`include_groups=False` requires pandas >= 2.2** — pyproject.toml allowed `pandas>=2.0.0`, which would crash on 2.0/2.1. Fixed: added version guard `if pd_ver >= (2, 2)` and pinned `pandas>=2.2.0,<3.0` in pyproject.toml.
- **`to_num()` silently swallowed parse errors** — non-numeric stock cells returned 0 with no warning. Fixed: parse_failures counter added; logged as WARNING after MB52 is loaded.
- **`Supply_Risk` column missing** — planner couldn't tell at a glance whether a PO would arrive before the job's start date. Fixed: new column added to READINESS_BOARD showing PO IN TIME / PO LATE / NO PO / "-".
- **`Total_Short_Qty` was computed but not shown** — the field existed in `df_jobs` but wasn't written to the Excel sheet. Fixed: now shown as column 16 on READINESS_BOARD.

### Bugs fixed in `validate_inputs.py`
- **Recommended columns not checked** — `Header SD order` and `MRP Type` are optional but silently missing changes simulation results. Fixed: warning printed if either is absent.
- **Blank `Material` codes not flagged** — build_readiness.py silently drops these rows; validator now warns with count.
- **Duplicate Order entries with different Start Dates** — breaks sort order in simulation. Validator now warns with list of affected orders.
- **Unexpected `Procurement Type` values** — unknown types are silently treated as external; validator now warns.
- **PO `Delivery Date` completeness not checked** — all-NA or all-unparseable dates would silently break Supply Risk. Validator now raises an error if no dates are parseable, and warns on partial failures.

### Documentation fixes
- `y00_zmp0.xlsx` (with zero) was wrong in CLAUDE.md and SAP_EXPORT_CHECKLIST.md; correct filename is `y00_zmpo.xlsx` (letter o). Code was always correct; docs were misleading.
- CLAUDE.md did not mention the Streamlit web UI (`app.py`) — added.
- CLAUDE.md did not mention `me5a_prs.xlsx` optional input — added.

### Dependency fixes in `pyproject.toml`
- `plotly` was missing from pyproject.toml (only in requirements.txt) — added with version bounds.
- All dependencies given upper-bound pins to prevent silent breaking changes on major upgrades.
- `requirements.txt` and `pyproject.toml` were inconsistent — pyproject.toml is now the single source of truth.
