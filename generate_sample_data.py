"""
Generates 3 synthetic Excel files (current week / last week / same week last year)
so the Weekly Sales Dashboard can be demoed without real company data.

Usage:
    python generate_sample_data.py
Then upload the 3 files it creates (in ./sample_data/) into the Streamlit app.
"""

import numpy as np
import pandas as pd
import os

np.random.seed(42)

CATEGORIES = [
    ("Men", "Shoes"), ("Men", "Jackets"), ("Men", "Pants"),
    ("Women", "Tops"), ("Women", "Dresses"), ("Women", "Shoes"),
    ("Kids", "Shoes"), ("Kids", "Jackets"),
]

PLACEHOLDER_IMAGE = "https://via.placeholder.com/150"


def make_week(n_articles=25, sales_scale=1.0, seed_offset=0):
    rng = np.random.default_rng(42 + seed_offset)
    cg = [CATEGORIES[i % len(CATEGORIES)] for i in range(n_articles)]
    sold = (rng.poisson(lam=20, size=n_articles) * sales_scale).astype(int)
    return pd.DataFrame({
        "Sold Items After Return": sold,
        "Return Rate": np.round(rng.uniform(1, 12, size=n_articles), 1),
        "CG2": [c[0] for c in cg],
        "CG3": [c[1] for c in cg],
        "Supplier Article Name": [f"Article {i+1:03d}" for i in range(n_articles)],
        "Image Link": [PLACEHOLDER_IMAGE] * n_articles,
    })


if __name__ == "__main__":
    os.makedirs("sample_data", exist_ok=True)

    current = make_week(sales_scale=1.15, seed_offset=0)   # this week: up
    last = make_week(sales_scale=1.0, seed_offset=1)        # last week: baseline
    year_ago = make_week(sales_scale=0.9, seed_offset=2)    # same week last year: lower

    current.to_excel("sample_data/current_week.xlsx", index=False)
    last.to_excel("sample_data/last_week.xlsx", index=False)
    year_ago.to_excel("sample_data/same_week_last_year.xlsx", index=False)

    print("Sample files written to ./sample_data/:")
    print("  - current_week.xlsx")
    print("  - last_week.xlsx")
    print("  - same_week_last_year.xlsx")
