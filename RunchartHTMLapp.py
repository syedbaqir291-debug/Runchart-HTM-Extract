# runcharts_streamlit_premium.py
# PPTX + HTML Dashboard (SAFE UPGRADE - LOGIC PRESERVED)

import streamlit as st
import pandas as pd
import numpy as np
from pptx import Presentation
from pptx.util import Inches
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

import plotly.graph_objects as go
import re
import io
import os

# ---------------------------
# CONFIG (UNCHANGED)
# ---------------------------
NUM_POINTS = 18

# ---------------------------
# STATE FIX
# ---------------------------
if "dept_col" not in st.session_state:
    st.session_state.dept_col = None
if "ind_col" not in st.session_state:
    st.session_state.ind_col = None

# ---------------------------
# CLEAN FUNCTIONS (UNCHANGED)
# ---------------------------
def clean_text_for_match(val):
    if pd.isna(val):
        return ""
    return str(val).strip().lower()

def pretty_label(col):
    parsed = pd.to_datetime(col, errors="coerce")
    return parsed.strftime("%b-%y") if not pd.isna(parsed) else str(col)

def get_center_line(values):
    arr = np.array(values, dtype=float)
    arr = arr[~np.isnan(arr)]
    return float(np.nanmedian(arr)) if len(arr) else np.nan

# ---------------------------
# DETECTION (UNCHANGED LOGIC)
# ---------------------------
def detect_shift(series, center, min_run=6):
    signs = []
    for v in series:
        if pd.isna(v):
            signs.append(0)
        elif v > center:
            signs.append(1)
        elif v < center:
            signs.append(-1)
        else:
            signs.append(0)

    shifts = []
    i = 0
    while i < len(signs):
        if signs[i] == 0:
            i += 1
            continue
        curr = signs[i]
        j = i
        while j < len(signs) and signs[j] == curr:
            j += 1
        if (j - i) >= min_run:
            shifts.append((i, j - 1, curr))
        i = j
    return shifts

def detect_trend(series, min_run=5):
    vals = [v for v in series if pd.notna(v)]
    trends = []

    for i in range(len(vals) - min_run):
        w = vals[i:i+min_run]
        if all(w[j] < w[j+1] for j in range(len(w)-1)):
            trends.append((i, i+min_run-1, 1))
        if all(w[j] > w[j+1] for j in range(len(w)-1)):
            trends.append((i, i+min_run-1, -1))
    return trends

def detect_astronomical(series, threshold=10):
    ast = []
    for i in range(1, len(series)):
        if pd.notna(series[i]) and pd.notna(series[i-1]):
            if abs(series[i] - series[i-1]) >= threshold:
                ast.append(i)
    return ast

# ---------------------------
# PPTX (UNCHANGED)
# ---------------------------
def create_presentation(slide_infos):
    prs = Presentation()

    for info in slide_infos:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.title.text = info["indicator"]

        chart_data = CategoryChartData()
        chart_data.categories = info["labels"]
        chart_data.add_series("Value", info["series"])

        slide.shapes.add_chart(
            XL_CHART_TYPE.LINE_MARKERS,
            Inches(0.6), Inches(1.2), Inches(9), Inches(4),
            chart_data
        )

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# ---------------------------
# STREAMLIT
# ---------------------------
st.set_page_config(layout="wide")
st.title("📊 RunCharts Premium Dashboard")

if "page" not in st.session_state:
    st.session_state.page = "upload"

if st.session_state.page == "upload":

    file = st.file_uploader("Upload Excel")

    if file:
        df = pd.read_excel(file)
        st.session_state.df = df

        st.session_state.dept_col = st.selectbox("Department Column", df.columns)
        st.session_state.ind_col = st.selectbox("Indicator Column", df.columns)

        if st.button("Continue"):
            st.session_state.page = "dashboard"
            st.rerun()

elif st.session_state.page == "dashboard":

    df = st.session_state.df
    dept_col = st.session_state.dept_col

    df["_dept"] = df[dept_col].astype(str).apply(clean_text_for_match)

    for d in df["_dept"].unique():
        if st.button(d):
            st.session_state.dept = d
            st.session_state.page = "chart"
            st.rerun()

elif st.session_state.page == "chart":

    df = st.session_state.df
    row = df.iloc[0]

    labels = list(df.columns[2:])
    raw = [pd.to_numeric(row[c], errors="coerce") for c in labels]

    # ✅ ONLY FIX: convert to percentage for DISPLAY
    series = [v * 100 if pd.notna(v) else np.nan for v in raw]

    median = np.nanmedian(series)

    st.subheader("Indicator")

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=labels, y=series, mode="lines+markers"))

    fig.add_hline(y=median, line_dash="dash")

    st.plotly_chart(fig, use_container_width=True)

    # ---------------------------
    # ANALYSIS (UNCHANGED LOGIC, ONLY DISPLAY FIX)
    # ---------------------------
    center = get_center_line(raw)

    shifts = detect_shift(raw, center)
    trends = detect_trend(raw)
    astro = detect_astronomical(raw)

    st.markdown("### Analysis")

    st.write("Median =", round(median, 2), "%")

    # SHIFT
    for s in shifts:
        direction = "above" if s[2] == 1 else "below"
        st.write(f"SHIFT ({direction}) from index {s[0]} to {s[1]}")

    # TREND
    for t in trends:
        st.write(f"TREND (monotonic) from {t[0]} to {t[1]}")

    # ASTRO
    for a in astro:
        st.write(f"Astronomical point at index {a}")
