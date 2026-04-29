"""
Pre-flight validation for SAP input files.
Catches 80% of problems before build_readiness.py runs.

Usage:
    python scripts/validate_inputs.py
    python scripts/validate_inputs.py --verbose
    python scripts/validate_inputs.py --inputs-dir /custom/path
"""

import argparse
import sys
from pathlib import Path
from datetime import datetime, timedelta

import pandas as pd


REQUIRED_COOIS_COLUMNS = [
    "Order",
    "Material",
    "Material Description",
    "Requirement Quantity",
    "Quantity withdrawn",
    "Procurement Type",
    "Header Material Text",
    "Header Basic Start Date",
    "Header Basic Finish Date",
]

# MB52 is a hierarchical dump, not a flat table — we validate structure instead
MB52_MIN_ROWS = 100  # Sanity check: should have hundreds at minimum


class ValidationError(Exception):
    """Raised when an input file fails validation."""


def validate_coois(path: Path, verbose: bool = False) -> dict:
    """Check COOIS export has expected columns and reasonable data."""
    if not path.exists():
        raise ValidationError(f"File not found: {path}")

    try:
        df = pd.read_excel(path)
    except Exception as e:
        raise ValidationError(f"Could not read {path.name}: {e}")

    # Column check
    missing = [c for c in REQUIRED_COOIS_COLUMNS if c not in df.columns]
    if missing:
        raise ValidationError(
            f"{path.name} missing required columns: {missing}\n"
            f"Found columns: {list(df.columns)}\n"
            f"Ensure you ran COOIS in Component view with full layout."
        )

    # Row count check
    if len(df) == 0:
        raise ValidationError(f"{path.name} is empty — check SAP filter settings")

    # Date column sanity — should have recent dates
    try:
        df["Header Basic Start Date"] = pd.to_datetime(
            df["Header Basic Start Date"], errors="coerce"
        )
        recent = df["Header Basic Start Date"].dropna()
        if len(recent) == 0:
            raise ValidationError(
                f"{path.name} has no valid start dates — check date column format"
            )

        latest = recent.max()
        today = pd.Timestamp.today().normalize()
        if latest < today - timedelta(days=180):
            raise ValidationError(
                f"{path.name} latest start date is {latest.date()} — "
                f"export is more than 6 months old"
            )
    except Exception as e:
        raise ValidationError(f"Date validation failed for {path.name}: {e}")

    # Required qty sanity
    if df["Requirement Quantity"].isna().all():
        raise ValidationError(f"{path.name} has no requirement quantities")

    stats = {
        "file": path.name,
        "rows": len(df),
        "unique_orders": df["Order"].nunique(),
        "unique_materials": df["Material"].nunique(),
        "date_range": f"{recent.min().date()} → {recent.max().date()}",
        "procurement_types": df["Procurement Type"].value_counts().to_dict(),
    }

    if verbose:
        print(f"\n✓ {path.name}")
        for k, v in stats.items():
            print(f"    {k}: {v}")

    return stats


def validate_mb52(path: Path, verbose: bool = False) -> dict:
    """Check MB52 export has the hierarchical structure we expect."""
    if not path.exists():
        raise ValidationError(f"File not found: {path}")

    try:
        df = pd.read_excel(path, header=None)
    except Exception as e:
        raise ValidationError(f"Could not read {path.name}: {e}")

    if len(df) < MB52_MIN_ROWS:
        raise ValidationError(
            f"{path.name} has only {len(df)} rows — expected {MB52_MIN_ROWS}+. "
            f"Check that the export includes all materials."
        )

    # Try parsing a few material rows to verify structure
    # Standard MB52 structure: row i = material header, row i+1 = location+qty
    materials_found = 0
    for i in range(5, min(50, len(df))):
        v0 = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ""
        if v0 and v0 not in ["nan", "Material", "Locat"] and pd.notna(df.iloc[i, 9]):
            materials_found += 1

    if materials_found == 0:
        raise ValidationError(
            f"{path.name} structure not recognized — "
            f"expected hierarchical MB52 format with materials starting around row 5. "
            f"Check that you exported in the standard MB52 layout."
        )

    stats = {
        "file": path.name,
        "total_rows": len(df),
        "materials_detected_in_first_50_rows": materials_found,
    }

    if verbose:
        print(f"\n✓ {path.name}")
        for k, v in stats.items():
            print(f"    {k}: {v}")

    return stats


PO_REQUIRED_COLUMNS = [
    "Material", "Purchasing Document", "PO-Quantity", "GR-Quantity", "Delivery Date", "Name",
]


def validate_pos(path: Path, verbose: bool = False) -> dict:
    """Check ZMPO purchase order export has expected columns and reasonable data."""
    if not path.exists():
        raise ValidationError(f"File not found: {path}")

    try:
        df = pd.read_excel(path)
    except Exception as e:
        raise ValidationError(f"Could not read {path.name}: {e}")

    missing = [c for c in PO_REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValidationError(
            f"{path.name} missing required columns: {missing}\n"
            f"Found columns: {list(df.columns)}"
        )

    if len(df) == 0:
        raise ValidationError(f"{path.name} is empty")

    has_material = df["Material"].notna() & (df["Material"].astype(str).str.strip() != "")
    delivery_range = pd.to_datetime(df["Delivery Date"], errors="coerce").dropna()

    stats = {
        "file": path.name,
        "total_rows": len(df),
        "rows_with_material_code": int(has_material.sum()),
        "rows_without_material_code": int((~has_material).sum()),
        "delivery_date_range": f"{delivery_range.min().date()} → {delivery_range.max().date()}" if len(delivery_range) else "no dates",
    }

    if verbose:
        print(f"\n✓ {path.name}")
        for k, v in stats.items():
            print(f"    {k}: {v}")

    return stats


def main():
    parser = argparse.ArgumentParser(description="Validate SAP input files")
    parser.add_argument(
        "--inputs-dir",
        type=Path,
        default=Path(__file__).parent.parent / "inputs",
        help="Directory containing SAP exports",
    )
    parser.add_argument("--verbose", "-v", action="store_true", help="Show detailed stats")
    args = parser.parse_args()

    inputs_dir = args.inputs_dir.resolve()

    print(f"Validating inputs in: {inputs_dir}")
    print("=" * 60)

    coois_path = inputs_dir / "coois_components.xlsx"
    mb52_path = inputs_dir / "mb52_stock.xlsx"
    po_path = inputs_dir / "y00_zmpo.xlsx"

    errors = []

    try:
        coois_stats = validate_coois(coois_path, verbose=args.verbose)
        print(
            f"✓ COOIS:  {coois_stats['rows']} rows, "
            f"{coois_stats['unique_orders']} orders, "
            f"{coois_stats['unique_materials']} materials"
        )
    except ValidationError as e:
        errors.append(str(e))
        print(f"✗ COOIS failed validation")

    try:
        mb52_stats = validate_mb52(mb52_path, verbose=args.verbose)
        print(f"✓ MB52:   {mb52_stats['total_rows']} rows")
    except ValidationError as e:
        errors.append(str(e))
        print(f"✗ MB52 failed validation")

    if po_path.exists():
        try:
            po_stats = validate_pos(po_path, verbose=args.verbose)
            print(
                f"✓ PO:     {po_stats['total_rows']} rows "
                f"({po_stats['rows_with_material_code']} with material code, "
                f"{po_stats['rows_without_material_code']} free-text)"
            )
        except ValidationError as e:
            errors.append(str(e))
            print(f"✗ PO file failed validation")
    else:
        print(f"- PO:     y00_zmpo.xlsx not found — PO coverage columns will be blank")

    print("=" * 60)

    if errors:
        print("\nERRORS:\n")
        for err in errors:
            print(f"  • {err}\n")
        sys.exit(1)

    print("\nAll inputs look good — ready to run build_readiness.py\n")
    sys.exit(0)


if __name__ == "__main__":
    main()
