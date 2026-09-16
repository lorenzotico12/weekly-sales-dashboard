# Weekly Sales Performance Dashboard

**[Try the live app →](https://weekly-sales-dashboard-nxkvukdlvqrq2xfgsbpgf7.streamlit.app/)**

A Streamlit app that takes three weekly sales exports (current week, last week,
same week last year), computes week-over-week and year-over-year KPIs, and
auto-generates a PowerPoint report — including an auto-populated "Top 5 Best
Sellers" slide with product images pulled from URLs.

Sample Excel files to test it with are in [`sample_data/`](sample_data/) — no
real company data needed (see below for details).

## What it demonstrates
- Data pipelines with **pandas** (column normalization, groupby aggregation, WoW/YoY deltas)
- Turning a manual reporting task into a **self-serve tool** (upload → click → download)
- Programmatic **PowerPoint generation** with `python-pptx` — a designed report (color palette,
  stat cards, dynamic image placement) rather than plain bullet points
- Threshold-based KPI status badges (on target / watch / off target) so the report reads as
  an assessment, not just a data dump — thresholds are configurable constants at the top of
  the script (`RETURN_RATE_TARGET`, `SALES_GROWTH_STRONG`, etc.)
- Input validation with clear, actionable error messages instead of raw stack traces
- A working, deployable **Streamlit** front end

## Run it locally
```bash
pip install -r requirements.txt
streamlit run weekly_sales_dashboard_app.py
```

## Try it without real company data
This repo includes a synthetic data generator so anyone (a recruiter, a
teammate) can demo the tool without needing real sales exports:
```bash
python generate_sample_data.py
```
This creates `sample_data/current_week.xlsx`, `last_week.xlsx`, and
`same_week_last_year.xlsx`. Upload all three into the app and click
**"Generate KPI Report and PowerPoint"**.

## Expected input format
Each Excel file needs these columns (case-insensitive):
| Column | Meaning |
|---|---|
| Sold Items After Return | units sold, net of returns |
| Return Rate | % returned |
| CG2 | top-level category (e.g. Men/Women/Kids) |
| CG3 | sub-category (e.g. Shoes/Jackets) |
| Supplier Article Name | product name |
| Image Link | URL to a product image |

## Known limitations
- Column names are hardcoded to one naming convention — uploads are
  validated with a clear error message if a required column is missing,
  but the app doesn't yet support alternate column-name mappings.
- If a product image fails to load, the report falls back to a styled
  placeholder card (by design, so one broken link doesn't crash the whole
  report) — but there's no logging of *which* images failed.
- No automated tests yet.
