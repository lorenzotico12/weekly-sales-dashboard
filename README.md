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
Each Excel file needs these columns (case-insensitive):
| Column | Meaning |
|---|---|
| Sold Items After Return | units sold, net of returns |
| Return Rate | % returned |
| CG2 | top-level category (e.g. Men/Women/Kids) |
| CG3 | sub-category (e.g. Shoes/Jackets) |
| Supplier Article Name | product name |
| Image Link | URL to a product image |

## Known limitations (be upfront about these if asked in an interview)
- Column names are still hardcoded to one naming convention — uploads are
  validated with a clear error message if a required column is missing,
  but the app doesn't yet support alternate column-name mappings.
- If a product image fails to load, the report falls back to a styled
  placeholder card (by design, so one broken link doesn't crash the whole
  report) — but there's no logging of *which* images failed.
- No automated tests yet.

---

## Putting this on your CV

**1. Get it on GitHub with this README.** A CV line means nothing without a
   link a recruiter can click. Create a public repo, push this code, and
   put the repo link on your CV/LinkedIn next to the project.

**2. Deploy a live demo (free, ~10 minutes).** Streamlit Community Cloud
   (share.streamlit.io) deploys directly from a GitHub repo — connect the
   repo, point it at `weekly_sales_dashboard_app.py`, and you get a public
   URL. This is the single highest-leverage thing you can do: "here's a
   link, try it yourself" beats any bullet point.

**3. Suggested CV bullet(s):**
   > Built and deployed a Streamlit tool that automates weekly sales
   > reporting — ingests raw Excel exports, computes WoW/YoY KPIs with
   > pandas, and generates a client-ready PowerPoint report (incl.
   > auto-fetched product images) in seconds instead of manually in Excel/PPT.

   If you added the NumPy step (see below), you can also say:
   > ...including vectorized (NumPy) outlier detection to flag anomalous
   > return-rate spikes.

**4. Be ready to talk about the *why*, not just the *what*.** The strongest
   answer to "tell me about this project" is the time/manual-work it
   replaced — e.g. "this used to be a person rebuilding this PPT by hand
   every Monday; this cut it to under a minute." If that's true for how you
   originally built it, use that framing.

**5. On the CV skill claim "Python (Pandas/NumPy)":** this script currently
   only uses pandas directly (NumPy runs underneath it, unseen). If you
   want the NumPy claim to hold up under a technical follow-up question,
   say the word and I'll add a small explicit NumPy component (e.g.
   z-score-based outlier flagging on return rates) — it's a natural fit
   and takes minutes to add.
