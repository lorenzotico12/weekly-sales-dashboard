# Weekly Sales Performance Dashboard

A Streamlit app that takes three weekly sales exports (current week, last week,
same week last year), computes week-over-week and year-over-year KPIs, and
auto-generates a PowerPoint report — including an auto-populated "Top 5 Best
Sellers" slide with product images pulled from URLs.

## What it demonstrates
- Data pipelines with **pandas** (column normalization, groupby aggregation, WoW/YoY deltas)
- Turning a manual reporting task into a **self-serve tool** (upload → click → download)
- Programmatic **PowerPoint generation** with `python-pptx` — a designed report (color palette,
  stat cards, dynamic image placement) rather than plain bullet points
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
