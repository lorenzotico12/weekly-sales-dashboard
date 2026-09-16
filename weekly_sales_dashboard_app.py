
import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

st.set_page_config(page_title="Weekly Sales Dashboard", layout="wide")
st.title("📈 Weekly Sales Performance Dashboard")

st.markdown("Upload your 3 Excel files: **Current Week**, **Last Week**, and **Same Week Last Year**.\n")
uploaded_files = st.file_uploader("Upload Excel Files", type=["xlsx"], accept_multiple_files=True)

# Normalize and map column names
def standardize_columns(df):
    df.columns = [col.strip().lower() for col in df.columns]
    col_map = {
        "sold itmens after return": "sold items after return",
        "sold items after return": "sold items after return",
        "return rate": "return rate",
        "cg2": "cg2",
        "cg3": "cg3",
        "supplier article name": "supplier article name",
        "image link": "image link"
    }
    return df.rename(columns={col: col_map.get(col, col) for col in df.columns})

def compute_kpis(current_df, last_df, year_df):
    current_sales = current_df["sold items after return"].sum()
    last_sales = last_df["sold items after return"].sum()
    year_sales = year_df["sold items after return"].sum()

    sales_wow = ((current_sales - last_sales) / last_sales) * 100 if last_sales else 0
    sales_yoy = ((current_sales - year_sales) / year_sales) * 100 if year_sales else 0

    current_return = current_df["return rate"].mean()
    last_return = last_df["return rate"].mean()
    year_return = year_df["return rate"].mean()

    gender_data = current_df.groupby("cg2")["sold items after return"].sum().sort_values(ascending=False)
    total = gender_data.sum()
    breakdown = [(str(g), int(c), (c / total) * 100 if total else 0) for g, c in gender_data.items()]

    return {
        "sales": int(current_sales),
        "sales_wow": sales_wow,
        "sales_yoy": sales_yoy,
        "return_rate": current_return,
        "return_wow": current_return - last_return,
        "return_yoy": current_return - year_return,
        "breakdown": breakdown,
    }

# ---------------------------------------------------------------------------
# PowerPoint design system ("Midnight Executive" palette)
# ---------------------------------------------------------------------------
NAVY = RGBColor(0x1E, 0x27, 0x61)
ICE = RGBColor(0xCA, 0xDC, 0xFC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x6E, 0x6E, 0x6E)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
RED = RGBColor(0xC6, 0x28, 0x28)

FONT_HEADER = "Cambria"
FONT_BODY = "Calibri"


def _delta_color(value):
    return GREEN if value >= 0 else RED


def _card(slide, left, top, width, height, fill=ICE):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    shape.shadow.inherit = False
    try:
        shape.adjustments[0] = 0.06
    except Exception:
        pass
    return shape


def _textbox(slide, left, top, width, height, text, size, color,
             bold=False, align=PP_ALIGN.LEFT, font=FONT_BODY):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    return box


def _multirun_textbox(slide, left, top, width, height, runs, align=PP_ALIGN.LEFT, font=FONT_BODY):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    for text, size, color, bold in runs:
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = font
    return box


def build_kpi_slide(prs, kpis):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = WHITE

    _textbox(slide, Inches(0.6), Inches(0.45), Inches(11.5), Inches(0.9),
             "Weekly KPI Summary", 38, NAVY, bold=True, font=FONT_HEADER)

    card_w = Inches(5.75)
    card_h = Inches(2.3)
    gap = Inches(0.35)
    top1 = Inches(1.7)
    left1 = Inches(0.6)
    left2 = left1 + card_w + gap
    pad = Inches(0.35)

    # Sales card
    _card(slide, left1, top1, card_w, card_h)
    _textbox(slide, left1 + pad, top1 + Inches(0.25), card_w - 2 * pad, Inches(0.35),
             "SALES (UNITS)", 14, NAVY, bold=True)
    _textbox(slide, left1 + pad, top1 + Inches(0.6), card_w - 2 * pad, Inches(1.0),
             f"{kpis['sales']:,}", 54, NAVY, bold=True, font=FONT_HEADER)
    _multirun_textbox(slide, left1 + pad, top1 + Inches(1.65), card_w - 2 * pad, Inches(0.5), [
        ("WoW  ", 14, GRAY, False),
        (f"{kpis['sales_wow']:+.1f}%", 14, _delta_color(kpis['sales_wow']), True),
        ("      YoY  ", 14, GRAY, False),
        (f"{kpis['sales_yoy']:+.1f}%", 14, _delta_color(kpis['sales_yoy']), True),
    ])

    # Return rate card
    _card(slide, left2, top1, card_w, card_h)
    _textbox(slide, left2 + pad, top1 + Inches(0.25), card_w - 2 * pad, Inches(0.35),
             "RETURN RATE", 14, NAVY, bold=True)
    _textbox(slide, left2 + pad, top1 + Inches(0.6), card_w - 2 * pad, Inches(1.0),
             f"{kpis['return_rate']:.1f}%", 54, NAVY, bold=True, font=FONT_HEADER)
    _multirun_textbox(slide, left2 + pad, top1 + Inches(1.65), card_w - 2 * pad, Inches(0.5), [
        ("WoW  ", 14, GRAY, False),
        (f"{kpis['return_wow']:+.1f} pts", 14, _delta_color(-kpis['return_wow']), True),
        ("      YoY  ", 14, GRAY, False),
        (f"{kpis['return_yoy']:+.1f} pts", 14, _delta_color(-kpis['return_yoy']), True),
    ])

    # Category breakdown cards
    top2 = top1 + card_h + Inches(0.35)
    breakdown = kpis["breakdown"]
    n = max(len(breakdown), 1)
    total_width = Inches(11.6)
    gap2 = Inches(0.3)
    card2_w = Emu(int((total_width - gap2 * (n - 1)) / n))
    card2_h = Inches(1.7)
    x = left1
    for gender, count, pct in breakdown:
        _card(slide, x, top2, card2_w, card2_h)
        _textbox(slide, x + Inches(0.25), top2 + Inches(0.2), card2_w - Inches(0.5), Inches(0.4),
                 gender.upper(), 14, NAVY, bold=True)
        _textbox(slide, x + Inches(0.25), top2 + Inches(0.55), card2_w - Inches(0.5), Inches(0.6),
                 f"{count:,}", 30, NAVY, bold=True, font=FONT_HEADER)
        _textbox(slide, x + Inches(0.25), top2 + Inches(1.2), card2_w - Inches(0.5), Inches(0.4),
                 f"{pct:.1f}% of sales", 12, GRAY)
        x = x + card2_w + gap2

    return slide


def _fetch_image_bytes(url):
    try:
        resp = requests.get(url, timeout=6, headers={"User-Agent": "Mozilla/5.0"})
        if resp.status_code == 200 and resp.content:
            return BytesIO(resp.content)
    except Exception:
        pass
    return None


def build_top5_slide(prs, df):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = WHITE

    _textbox(slide, Inches(0.6), Inches(0.45), Inches(12.1), Inches(0.75),
             "Top 5 Best Sellers — Current Week", 30, NAVY, bold=True, font=FONT_HEADER)

    top5 = df.sort_values("sold items after return", ascending=False).head(5)
    n = max(len(top5), 1)

    left0 = Inches(0.6)
    top0 = Inches(1.6)
    gap = Inches(0.25)
    total_w = Inches(12.1)
    card_w = Emu(int((total_w - gap * (n - 1)) / n))
    card_h = Inches(4.15)
    img_size = min(card_w - Inches(0.4), Inches(1.9))

    x = left0
    for _, row in top5.iterrows():
        _card(slide, x, top0, card_w, card_h)

        img_left = x + (card_w - img_size) / 2
        img_top = top0 + Inches(0.25)

        image_stream = _fetch_image_bytes(row["image link"])
        if image_stream is not None:
            slide.shapes.add_picture(image_stream, img_left, img_top, width=img_size, height=img_size)
        else:
            circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, img_left, img_top, img_size, img_size)
            circle.fill.solid()
            circle.fill.fore_color.rgb = NAVY
            circle.line.fill.background()
            circle.shadow.inherit = False
            tf = circle.text_frame
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = "No Image"
            run.font.size = Pt(13)
            run.font.bold = True
            run.font.color.rgb = WHITE
            run.font.name = FONT_BODY

        text_top = img_top + img_size + Inches(0.2)
        pad = Inches(0.2)
        _textbox(slide, x + pad, text_top, card_w - 2 * pad, Inches(0.35),
                 str(row["supplier article name"]), 13, NAVY, bold=True, align=PP_ALIGN.CENTER)
        _textbox(slide, x + pad, text_top + Inches(0.38), card_w - 2 * pad, Inches(0.3),
                 f"{row['cg2']} / {row['cg3']}", 11, GRAY, align=PP_ALIGN.CENTER)
        _textbox(slide, x + pad, text_top + Inches(0.72), card_w - 2 * pad, Inches(0.5),
                 f"{int(row['sold items after return'])} units", 17, NAVY, bold=True,
                 align=PP_ALIGN.CENTER, font=FONT_HEADER)

        x = x + card_w + gap

    return slide


def export_to_powerpoint(kpis, df):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    build_kpi_slide(prs, kpis)
    build_top5_slide(prs, df)

    os.makedirs("output", exist_ok=True)
    ppt_path = "output/Weekly_Sales_Report.pptx"
    prs.save(ppt_path)
    return ppt_path

REQUIRED_COLUMNS = {
    "sold items after return", "return rate", "cg2", "cg3",
    "supplier article name", "image link"
}

if uploaded_files and len(uploaded_files) == 3:
    labels = ["Current Week", "Last Week", "Same Week Last Year"]

    # Parse each file exactly once and cache the result in session_state,
    # keyed by filename+size. This avoids re-reading the upload stream on
    # every script rerun (Streamlit reruns the whole script on each button
    # click), which is what was causing the KeyError on "cg2" -- a re-read
    # of an already-consumed upload stream can come back empty/malformed.
    dataframes = []
    parse_error = False
    for file in uploaded_files:
        cache_key = f"parsed::{file.name}::{file.size}"
        if cache_key not in st.session_state:
            file.seek(0)
            df = standardize_columns(pd.read_excel(file))
            missing = REQUIRED_COLUMNS - set(df.columns)
            if missing:
                st.error(
                    f"**{file.name}** is missing required column(s): "
                    f"{', '.join(sorted(missing))}.\n\n"
                    f"Columns found: {', '.join(df.columns)}"
                )
                parse_error = True
                continue
            st.session_state[cache_key] = df
        dataframes.append(st.session_state[cache_key])

    if parse_error:
        st.stop()

    for label, df in zip(labels, dataframes):
        st.subheader(f"{label} Summary")
        st.dataframe(df.head())

    if st.button("Generate KPI Report and PowerPoint"):
        kpis = compute_kpis(*dataframes)
        ppt_path = export_to_powerpoint(kpis, dataframes[0])
        st.success("PowerPoint generated successfully!")
        with open(ppt_path, "rb") as f:
            st.download_button("📥 Download PowerPoint", f, file_name="Weekly_Sales_Report.pptx")

else:
    st.info("Please upload 3 Excel files.")
