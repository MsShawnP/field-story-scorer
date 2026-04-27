"""
Generate sample xlsx files for testing field-story-scorer.

Creates two files in samples/input/:
  sample_sales.xlsx        — clean dataset with a range of field types
  sample_mixed_types.xlsx  — dataset with genuine mixed-type cells in one column,
                             designed to demonstrate the --strict-types flag

Usage:
    python generate_sample.py
"""

import random
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook

OUTPUT_DIR = Path(__file__).parent / "samples" / "input"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
RNG = np.random.default_rng(42)
random.seed(42)


def make_clean_dataset(n: int = 500) -> None:
    """Write sample_sales.xlsx — a varied but clean dataset."""
    regions = ["Northeast", "Southeast", "Midwest", "West", "Southwest"]
    categories = ["Widget A", "Widget B", "Widget C", "Gadget X", "Gadget Y"]
    statuses = ["Active", "Inactive", "Pending", "Closed"]

    df = pd.DataFrame({
        "customer_id":      [f"CUST{str(i).zfill(5)}" for i in range(1, n + 1)],
        "region":           [random.choice(regions) for _ in range(n)],
        "product_category": [random.choice(categories) for _ in range(n)],
        "order_status":     [random.choice(statuses) for _ in range(n)],
        "revenue":          RNG.lognormal(mean=5, sigma=1.5, size=n).round(2),
        "units_sold":       RNG.integers(1, 200, size=n),
        "discount_pct":     [random.choice([0, 0.05, 0.10, 0.15, 0.20, None]) for _ in range(n)],
        "sales_rep":        [f"Rep{random.randint(1, 25)}" for _ in range(n)],
        "order_date":       [
            (datetime(2023, 1, 1) + timedelta(days=random.randint(0, 365))).strftime("%Y-%m-%d")
            for _ in range(n)
        ],
        "days_to_close":    RNG.integers(1, 90, size=n).astype(float),
        "is_renewal":       RNG.choice([True, False], size=n),
        "customer_score":   RNG.normal(75, 15, size=n).clip(0, 100).round(1),
        "notes":            [random.choice(["Good client", "Needs follow-up", None, "VIP", None, None]) for _ in range(n)],
        "constant_col":     ["FIXED"] * n,
        "mostly_null":      [random.choice([None, None, None, None, "value"]) for _ in range(n)],
    })

    path = OUTPUT_DIR / "sample_sales.xlsx"
    df.to_excel(path, index=False, sheet_name="Sales")
    print(f"Created {path}  ({len(df)} rows × {len(df.columns)} columns)")


def make_mixed_type_dataset(n: int = 200) -> None:
    """
    Write sample_mixed_types.xlsx using openpyxl directly so that cells in
    'revenue_mixed' are genuine native types (float or str), not pandas-coerced.

    This is the canonical test file for --strict-types mode:
      - Standard run:      revenue_mixed scores near-identical to revenue
      - --strict-types run: revenue_mixed shows score penalty + type_mix breakdown
    """
    revenues = RNG.lognormal(mean=5, sigma=1.5, size=n).round(2).tolist()
    bad_rows = set(random.sample(range(n), 15))

    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["customer_id", "revenue", "revenue_mixed", "region"])

    regions = ["Northeast", "Southeast", "Midwest", "West", "Southwest"]
    for i in range(n):
        rev = float(revenues[i])
        mixed = "N/A" if i in bad_rows else rev
        ws.append([f"CUST{str(i + 1).zfill(5)}", rev, mixed, random.choice(regions)])

    path = OUTPUT_DIR / "sample_mixed_types.xlsx"
    wb.save(path)
    print(f"Created {path}  ({n} rows, 15 intentional string cells in revenue_mixed)")
    print()
    print("Demonstrate the difference:")
    print(f"  python scorer.py --input {path} --output-dir samples/output/")
    print(f"  python scorer.py --input {path} --output-dir samples/output/ --strict-types")


if __name__ == "__main__":
    make_clean_dataset()
    make_mixed_type_dataset()
