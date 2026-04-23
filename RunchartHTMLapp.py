# runcharts_streamlit_premium.py
# PPTX + HTML Dashboard (SAFE UPGRADE - LOGIC PRESERVED)

import streamlit as st
import pandas as pd
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE

import plotly.graph_objects as go
import re
import difflib
import io
import os
from datetime import datetime

# ---------------------------
# CONFIG (UNCHANGED)
# ---------------------------
NUM_POINTS = 18
CENTER_LINE_METHOD = "median"
ASTRO_THRESHOLD = 10

NON_DATA_COLS_LOWER = {"department", "indicator", "target", "benchmark/ category", "benchmark", "frequency"}

TITLE_BOX = (0.6, 0.4, 9.0, 0.9)
CHART_BOX = (0.6, 1.4, 9.0, 4.3)
NOTES_BOX = (0.6, 6.0, 9.0, 1.3)

LOG_DIR = "logs"
UPLOAD_DIR = os.path.join(LOG_DIR, "uploads")
LOG_FILE = os.path.join(LOG_DIR, "activity_log.csv")

os.makedirs(UPLOAD_DIR, exist_ok=True)

# ---------------------------
# LOGIN
# ---------------------------
PREMIUM_LOGIN = "Pakistan@1947"
PREMIUM_PASSWORD = "Pakistan@1947"

# ---------------------------
# STATE FIX
# ---------------------------
if "dept_col" not in st.session_state:
    st.session_state.dept_col = None

if "ind_col" not in st.session_state:
    st.session_state.ind_col = None

# ---------------------------
# CLEAN FUNCTIONS
# ---------------------------
def clean_text_for_match(val):
    if pd.isna(val):
        return ""
    s = str(val)
    s = re.sub(r"[\u200B\u200C\u200D\uFEFF]", "", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip().lower()

def pretty_label(col):
    try:
        parsed = pd.to_datetime(col, errors="coerce")
        if not pd.isna(parsed):
            return parsed.strftime("%b-%y")
    except:
        pass
    return str(col)

def get_center_line(values, method="median"):
    arr = np.array(values, dtype=float)
    arr = arr[~np.isnan(arr)]
    if arr.size == 0:
        return np.nan
    return float(np.nanmedian(arr)) if method == "median" else float(np.nanmean(arr))

# ---------------------------
# DETECTION LOGIC
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
    i, n = 0, len(signs)

    while i < n:
        if signs[i] == 0:
            i += 1
            continue

        curr = signs[i]
        j = i + 1
        count = 1

        while j < n and signs[j] == curr:
            count += 1
            j += 1

        if count >= min_run:
            shifts.append((i, j - 1, curr))

        i = j

    return shifts

def detect_trend(series, min_run=5):
    trends = []
    vals = [v for v in series if pd.notna(v)]

    for i in range(len(vals) - min_run):
        window = vals[i:i+min_run]
        if all(window[j] < window[j+1] for j in range(len(window)-1)):
            trends.append((i, i+min_run-1, 1))
        if all(window[j] > window[j+1] for j in range(len(window)-1)):
            trends.append((i, i+min_run-1, -1))

    return trends

def detect_astronomical(series, median_val, threshold=10):
    ast = []
    for i in range(1, len(series)):
        if pd.notna(series[i]) and pd.notna(series[i-1]):
            if abs(series[i] - series[i-1]) >= threshold:
                ast.append(i)
    return ast

# ---------------------------
# FORMATTER (FIXED OUTPUT LAYER)
# ---------------------------
def format_analysis(labels, series, center, shifts, trends, astro):
    output = []

    # Median
    if pd.isna(center):
        output.append("Median = N/A")
    else:
        output.append(f"Median = {round(center,1)}")

    # SHIFT
    for s in shifts:
        direction = "above" if s[2] == 1 else "below"
        output.append(f"SHIFT ({direction}) from {labels[s[0]]} to {labels[s[1]]}")

    # TREND
    for t in trends:
        direction = "increasing" if t[2] == 1 else "decreasing"
        output.append(f"TREND ({direction}) from {labels[t[0]]} to {labels[t[1]]}")

    # ASTRO
    for a in astro:
        val = series[a]
        output.append(f"Astronomical point at {labels[a]} (value={round(val,2)})")

    return output

# ---------------------------
# STREAMLIT APP
# ---------------------------
st.set_page_config(layout="wide")
st.title("📊 RunCharts Premium Dashboard")

if "page" not in st.session_state:
    st.session_state.page = "upload"

if "df" not in st.session_state:
    st.session_state.df = None

# ---------------------------
# PAGE 1 - UPLOAD
# ---------------------------
if st.session_state.page == "upload":

    file = st.file_uploader("Upload Excel")

    if file:
        df = pd.read_excel(file)
        st.session_state.df = df

        st.success("File Loaded")

        st.markdown("### ⚠️ Step Required")

        dept_col = st.selectbox("Select Department Column", df.columns)
        ind_col = st.selectbox("Select Indicator Column", df.columns)

        if st.button("Confirm & Continue"):
            st.session_state.dept_col = dept_col
            st.session_state.ind_col = ind_col
            st.session_state.page = "dashboard"
            st.rerun()

# ---------------------------
# PAGE 2 - DASHBOARD
# ---------------------------
elif st.session_state.page == "dashboard":

    df = st.session_state.df

    dept_col = st.session_state.dept_col
    ind_col = st.session_state.ind_col

    df["_dept_clean"] = df[dept_col].astype(str).apply(clean_text_for_match)

    departments = df["_dept_clean"].unique()

    st.subheader("Departments")

    for d in departments:
        if st.button(f"📁 {d}"):
            st.session_state.selected_dept = d
            st.session_state.page = "indicators"
            st.rerun()

    if st.button("⬅ Back"):
        st.session_state.page = "upload"
        st.rerun()

# ---------------------------
# PAGE 3 - INDICATORS
# ---------------------------
elif st.session_state.page == "indicators":

    df = st.session_state.df
    dept = st.session_state.selected_dept

    df = df[df["_dept_clean"] == dept]

    st.subheader(f"Indicators - {dept}")

    for i, row in df.iterrows():
        if st.button(str(row[st.session_state.ind_col])):
            st.session_state.selected_indicator = i
            st.session_state.page = "chart"
            st.rerun()

    if st.button("🏠 Main Menu"):
        st.session_state.page = "dashboard"
        st.rerun()

# ---------------------------
# PAGE 4 - CHART
# ---------------------------
elif st.session_state.page == "chart":

    df = st.session_state.df
    row = df.iloc[st.session_state.selected_indicator]

    labels = list(df.columns[2:])

    series = [pd.to_numeric(row[c], errors="coerce") for c in labels]

    clean_series = [np.nan if pd.isna(v) else v for v in series]

    median = np.nanmedian(clean_series)

    st.subheader(row[st.session_state.ind_col])

    st.plotly_chart(
        go.Figure().add_trace(
            go.Scatter(x=labels, y=series, mode="lines+markers")
        ).add_hline(y=median),
        use_container_width=True
    )

    # ---------------------------
    # ANALYSIS (FIXED FORMAT OUTPUT)
    # ---------------------------
    st.markdown("### 📌 Analysis Summary")

    raw_series = [pd.to_numeric(row[c], errors="coerce") for c in labels]

    center = get_center_line(raw_series)
    shifts = detect_shift(raw_series, center)
    trends = detect_trend(raw_series)
    astro = detect_astronomical(raw_series, center)

    formatted = format_analysis(labels, raw_series, center, shifts, trends, astro)

    for line in formatted:
        st.write(line)

    if not shifts and not trends and not astro:
        st.success("No variation detected")

    if st.button("⬅ Back"):
        st.session_state.page = "indicators"
        st.rerun()

    if st.button("🏠 Main Menu"):
        st.session_state.page = "dashboard"
        st.rerun()
