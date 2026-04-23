# runcharts_streamlit_premium.py
# FIXED VERSION (NO LOGIC CHANGE — ONLY STABILITY + FORMATTING FIXES)

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
# STATE FIX (IMPORTANT)
# ---------------------------
if "dept_col" not in st.session_state:
    st.session_state.dept_col = None
if "ind_col" not in st.session_state:
    st.session_state.ind_col = None

# ---------------------------
# CLEANING
# ---------------------------
def clean_text_for_match(val):
    if pd.isna(val):
        return ""
    return str(val).strip().lower()

def pretty_label(col):
    """ONLY month-year, NO TIME"""
    try:
        dt = pd.to_datetime(col, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%b-%y")
    except:
        pass
    return str(col)

def to_float(v):
    try:
        if pd.isna(v):
            return np.nan
        return float(str(v).replace("%", "").strip())
    except:
        return np.nan

# ---------------------------
# CENTER LINE
# ---------------------------
def get_center_line(values):
    arr = np.array([v for v in values if pd.notna(v)], dtype=float)
    if len(arr) == 0:
        return np.nan
    return np.nanmedian(arr)

# ---------------------------
# SHIFT / TREND / ASTRO (FIXED OUTPUT FORMAT)
# ---------------------------
def detect_shift(series, center, labels):
    shifts = []
    for i in range(1, len(series)):
        if pd.isna(series[i]) or pd.isna(series[i-1]):
            continue

        if series[i] > center and series[i-1] <= center:
            shifts.append(f"SHIFT (above) from {labels[i-1]} to {labels[i]}")
        elif series[i] < center and series[i-1] >= center:
            shifts.append(f"SHIFT (below) from {labels[i-1]} to {labels[i]}")
    return shifts

def detect_trend(series, labels):
    trends = []
    for i in range(len(series)-2):
        window = series[i:i+3]
        if all(pd.notna(x) for x in window):
            if window[0] < window[1] < window[2]:
                trends.append(f"TREND (increasing) from {labels[i]} to {labels[i+2]}")
            elif window[0] > window[1] > window[2]:
                trends.append(f"TREND (decreasing) from {labels[i]} to {labels[i+2]}")
    return trends

def detect_astro(series, labels):
    astro = []
    for i in range(1, len(series)):
        if pd.notna(series[i]) and pd.notna(series[i-1]):
            if abs(series[i] - series[i-1]) >= 0.1:
                astro.append(f"Astronomical point at {labels[i]} (value={round(series[i]*100,1)}%)")
    return astro

# ---------------------------
# LIMIT LAST 18 POINTS ONLY
# ---------------------------
def limit_last_18(labels, series):
    labels = labels[-NUM_POINTS:]
    series = series[-NUM_POINTS:]
    return labels, series

# ---------------------------
# STREAMLIT APP
# ---------------------------
st.set_page_config(layout="wide")
st.title("📊 RunCharts Premium Dashboard (Fixed)")

if "page" not in st.session_state:
    st.session_state.page = "upload"

# ---------------------------
# PAGE 1
# ---------------------------
if st.session_state.page == "upload":

    file = st.file_uploader("Upload Excel")

    if file:
        df = pd.read_excel(file)
        st.session_state.df = df

        st.success("File Loaded")

        dept_col = st.selectbox("Select Department Column", df.columns)
        ind_col = st.selectbox("Select Indicator Column", df.columns)

        if st.button("Confirm & Continue"):
            st.session_state.dept_col = dept_col
            st.session_state.ind_col = ind_col
            st.session_state.page = "dashboard"
            st.rerun()

# ---------------------------
# PAGE 2
# ---------------------------
elif st.session_state.page == "dashboard":

    df = st.session_state.df
    dept_col = st.session_state.dept_col
    ind_col = st.session_state.ind_col

    df["_dept_clean"] = df[dept_col].astype(str).apply(clean_text_for_match)

    departments = df["_dept_clean"].unique()

    for d in departments:
        if st.button(f"📁 {d}"):
            st.session_state.selected_dept = d
            st.session_state.page = "indicators"
            st.rerun()

# ---------------------------
# PAGE 3
# ---------------------------
elif st.session_state.page == "indicators":

    df = st.session_state.df
    dept = st.session_state.selected_dept

    df = df[df["_dept_clean"] == dept]

    for i, row in df.iterrows():
        if st.button(str(row[st.session_state.ind_col])):
            st.session_state.selected_indicator = i
            st.session_state.page = "chart"
            st.rerun()

# ---------------------------
# PAGE 4 - CHART + FIXED ANALYSIS
# ---------------------------
elif st.session_state.page == "chart":

    df = st.session_state.df
    row = df.iloc[st.session_state.selected_indicator]

    cols = list(df.columns[2:])

    labels = [pretty_label(c) for c in cols]

    series = [to_float(row[c]) for c in cols]

    # CLEAN + LIMIT 18
    labels, series = limit_last_18(labels, series)

    median = get_center_line(series)

    # convert to percent
    median_display = f"{round(median*100,1)}%" if not np.isnan(median) else "N/A"

    st.subheader(row[st.session_state.ind_col])

    st.write("Median =", median_display)

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=labels, y=series, mode="lines+markers"))
    fig.add_hline(y=median)

    st.plotly_chart(fig, use_container_width=True)

    # ---------------------------
    # ANALYSIS FIXED OUTPUT
    # ---------------------------
    st.markdown("### 📌 Analysis Summary")

    shifts = detect_shift(series, median, labels)
    trends = detect_trend(series, labels)
    astro = detect_astro(series, labels)

    if shifts:
        for s in shifts:
            st.write(s)

    if trends:
        for t in trends:
            st.write(t)

    if astro:
        for a in astro:
            st.write(a)

    if not shifts and not trends and not astro:
        st.success("No Shift / Trend / Astronomical variation detected")

    if st.button("⬅ Back"):
        st.session_state.page = "indicators"
        st.rerun()

    if st.button("🏠 Main Menu"):
        st.session_state.page = "dashboard"
        st.rerun()
