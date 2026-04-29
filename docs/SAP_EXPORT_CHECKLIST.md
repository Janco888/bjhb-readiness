# SAP EXPORT CHECKLIST

Keep this at your desk. Follow it exactly every Monday. If SAP's behaviour changes, update this file — don't change the Python scripts.

---

## Export 1 of 2: COOIS Components

Save as: `inputs/coois_components.xlsx`

**Transaction:** `COOIS`

**Selection screen:**

| Field | Value |
|---|---|
| Plant | `502` |
| Order type | (leave blank — all types) |
| Order status | Tick `REL` (Released) ✓ — ensure `DLV` and `TECO` are **not** ticked |
| Basic start date | Today - 30 days → Today + 90 days |
| View | **Components** (radio button) |

**Before executing — layout check:**

Click `Settings → Layout → Change`. Ensure these columns are in the layout:

- [ ] Order
- [ ] Material
- [ ] Material Description
- [ ] Requirement Quantity
- [ ] Quantity withdrawn
- [ ] Procurement Type
- [ ] Header Material Text
- [ ] Header Basic Start Date
- [ ] Header Basic Finish Date
- [ ] Header SD order

**Execute** (F8)

**Export:**
- Menu: `List → Export → Spreadsheet`
- Format: `Excel .xlsx`
- Filename: `coois_components.xlsx`
- Save to: `inputs/` folder of this project

---

## Export 2 of 2: MB52 Stock

Save as: `inputs/mb52_stock.xlsx`

**Transaction:** `MB52`

**Selection screen:**

| Field | Value |
|---|---|
| Material | (leave blank — all materials) |
| Plant | `502` |
| Storage location | (leave blank — all locations) |
| Special stock indicator | (leave blank) |

**Execute** (F8)

**Export:**
- Menu: `List → Export → Spreadsheet`
- Format: `Excel .xlsx`
- Filename: `mb52_stock.xlsx`
- Save to: `inputs/` folder of this project

---

---

## Export 3 of 3: Open Purchase Orders (ZMPO) — Optional but Recommended

Save as: `inputs/y00_zmpo.xlsx`

**Transaction:** Custom Z-report — `ZMPO`

**Selection screen:**

| Field | Value |
|---|---|
| Plant | `502` |
| Open items only | Yes (or equivalent filter in the Z-report) |

**Required columns in layout:**

- [ ] Material
- [ ] Purchasing Document
- [ ] Item
- [ ] Name (Supplier name)
- [ ] PO-Quantity
- [ ] GR-Quantity
- [ ] Delivery Date

**Export:** same Excel export process as COOIS and MB52.

**Note:** Rows without a material code (free-text POs — e.g., drawings sent to laser cutters, consumables) are automatically filtered out by the pipeline. Only PO lines with a SAP material number are matched to the readiness dashboard.

**What this enables:** Short components on the dashboard will show the earliest incoming PO number, delivery date, and open quantity — so the planner can see at a glance whether a shortage will self-resolve before the job needs to start.

---

## After All Exports

1. Verify all three files are in `inputs/` and named exactly as above
2. Open each file briefly to confirm it has data (not just headers)
3. Run the dashboard:

```bash
# In Claude Code
/readiness

# Or terminal
python scripts/build_readiness.py
```

---

## Troubleshooting

### COOIS export is huge / takes forever

- Narrow the date range — try Today → Today + 60 days
- Ensure `Order Status = REL` is ticked (not all statuses)
- Remove any user-specific filter that might have crept in

### MB52 export has 0 rows

- You're probably filtering by material or storage location accidentally — clear all filters
- Verify plant is `502`

### Files save but Python complains about format

- SAP sometimes saves `.xlsx` as HTML-wrapped XML instead of real xlsx
- Fix: choose `Spreadsheet → In Existing Format` or `Spreadsheet → XLSX` specifically
- If still broken, save as `.xls` and Python will still read it (rename to `.xlsx` if needed)

### A column you need isn't in the layout

- `Settings → Layout → Change → Current`
- Drag the missing column from "Hidden Fields" to "Displayed Fields"
- `Save layout` so it persists next time

---

## When SAP Layouts Change

If IT pushes a SAP update and a column name or format changes:

1. **Don't patch the Python yet.** Run `/readiness` — the validator will report what's missing.
2. If the fix is adding a column to the SAP layout, update this checklist.
3. If the fix requires code changes, flag it to the maintainer — don't modify blindly.
