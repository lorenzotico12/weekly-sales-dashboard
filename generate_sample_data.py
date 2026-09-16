"""
Generates 3 synthetic Excel files (current week / last week / same week last year)
so the Weekly Sales Dashboard can be demoed without real company data.

These files intentionally look like a real-world wide export (~24 columns,
many irrelevant to the app) rather than a clean 6-column file, so uploading
them actually demonstrates the app's column normalization: it picks out
just the columns it needs and ignores the rest.

Usage:
    python generate_sample_data.py
Then upload the 3 files it creates (in ./sample_data/) into the Streamlit app.
"""

import numpy as np
import pandas as pd
import os
import datetime

np.random.seed(42)

CATEGORIES = [
    ("Men", "Shoes"), ("Men", "Jackets"), ("Men", "Pants"),
    ("Women", "Tops"), ("Women", "Dresses"), ("Women", "Shoes"),
    ("Kids", "Shoes"), ("Kids", "Jackets"),
]

PLACEHOLDER_IMAGE = "https://via.placeholder.com/150"
BRANDS = ["Northfield", "Aurelia", "Kestrel & Co", "Marlowe"]
COLORS = ["Black", "Navy", "Olive", "Cream", "Rust", "Grey"]
COUNTRIES = ["DE", "FR", "IT", "ES", "NL"]


def make_week(n_articles=25, sales_scale=1.0, seed_offset=0, week_label="202501"):
    rng = np.random.default_rng(42 + seed_offset)
    cg = [CATEGORIES[i % len(CATEGORIES)] for i in range(n_articles)]
    sold_after = (rng.poisson(lam=20, size=n_articles) * sales_scale).astype(int)
    return_rate = np.round(rng.uniform(1, 12, size=n_articles), 1)
    sold_before = (sold_after / (1 - return_rate / 100)).round().astype(int)

    # Realistic wide export: the app only needs 6 of these columns.
    # Column order deliberately mixed, like a real system export.
    return pd.DataFrame({
        "Supplier Name": [rng.choice(BRANDS) for _ in range(n_articles)],
        "Supplier Code": [f"SUP-{1000 + i}" for i in range(n_articles)],
        "Brand Code": [f"BR{rng.integers(10, 99)}" for _ in range(n_articles)],
        "Brand Name": [rng.choice(BRANDS) for _ in range(n_articles)],
        "Article Type": [c[1] for c in cg],
        "Sales Season": ["SS25" if week_label.startswith("2025") else "FW24"] * n_articles,
        "Department (CG1)": ["Apparel & Footwear"] * n_articles,
        "CG2": [c[0] for c in cg],
        "CG3": [c[1] for c in cg],
        "CG4": [rng.choice(["Casual", "Formal", "Sport"]) for _ in range(n_articles)],
        "CG5": [rng.choice(["Core", "Seasonal"]) for _ in range(n_articles)],
        "Article Activation Date": [
            (datetime.date(2025, 1, 1) - datetime.timedelta(days=int(rng.integers(0, 200)))).isoformat()
            for _ in range(n_articles)
        ],
        "Sustainable Flag": [rng.choice(["Y", "N"]) for _ in range(n_articles)],
        "Plus Size Flag": [rng.choice(["Y", "N"]) for _ in range(n_articles)],
        "Week (YYYYWW)": [week_label] * n_articles,
        "Country": [rng.choice(COUNTRIES) for _ in range(n_articles)],
        "Sold Items Before Return": sold_before,
        "Sold Items After Return": sold_after,
        "Return Rate": return_rate,
        "Returns Size Small": rng.integers(0, 5, size=n_articles),
        "Returns Size Big": rng.integers(0, 5, size=n_articles),
        "Returns Dislike": rng.integers(0, 8, size=n_articles),
        "Supplier Article Code": [f"ART-{i+1:04d}" for i in range(n_articles)],
        "Supplier Article Name": [f"Article {i+1:03d}" for i in range(n_articles)],
        "Supplier Color Code": [f"C{rng.integers(100, 999)}" for _ in range(n_articles)],
        "Supplier Color Description": [rng.choice(COLORS) for _ in range(n_articles)],
        "Image Link": [PLACEHOLDER_IMAGE] * n_articles,
    })


if __name__ == "__main__":
    os.makedirs("sample_data", exist_ok=True)

    current = make_week(sales_scale=1.15, seed_offset=0, week_label="202506")   # this week: up
    last = make_week(sales_scale=1.0, seed_offset=1, week_label="202505")        # last week: baseline
    year_ago = make_week(sales_scale=0.9, seed_offset=2, week_label="202406")    # same week last year: lower

    current.to_excel("sample_data/current_week.xlsx", index=False)
    last.to_excel("sample_data/last_week.xlsx", index=False)
    year_ago.to_excel("sample_data/same_week_last_year.xlsx", index=False)

    print("Sample files written to ./sample_data/ (24 columns each, mimicking a real export):")
    print("  - current_week.xlsx")
    print("  - last_week.xlsx")
    print("  - same_week_last_year.xlsx")
