# runcharts_streamlit_premium.py
# FULL FIXED VERSION (SHIFT + TREND + ASTRO RESTORED CORRECTLY)

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import re

# ---------------------------
# CONFIG (UNCHANGED)
# ---------------------------
NUM_POINTS = 18
SHIFT_MIN_RUN = 6
TREND_MIN_RUN = 5
ASTRO_THRESHOLD = 0.1  # for normalized values

# ---------------------------
# CLEANING
# ---------------------------
def clean_text_for_match(val):
    if pd.isna(val):
        return ""
    return str(val).strip().lower()

def pretty_label(col):
    dt = pd.to_datetime(col, errors="coerce")
    if not pd.isna(dt):
        return dt.strftime("%b-%y")
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
# SHIFT (CORRECT ORIGINAL RULE)
# ---------------------------
def detect_shift(series, center, labels):
    shifts = []
    i = 0

    while i < len(series):
        if pd.isna(series[i]):
            i += 1
            continue

        direction = 1 if series[i] > center else -1
        start = i
        count = 1
        j = i + 1

        while j < len(series):
            if pd.isna(series[j]):
                break

            if (series[j] > center and direction == 1) or (series[j] < center and direction == -1):
                count += 1
                j += 1
            else:
                break

        if count >= SHIFT_MIN_RUN:
            shifts.append(
                f"SHIFT ({'above' if direction==1 else 'below'}) from {labels[start]} to {labels[j-1]}"
            )

        i = j

    return shifts

# ---------------------------
# TREND (CORRECT ORIGINAL RULE)
# ---------------------------
def detect_trend(series, labels):
    trends = []
    i = 0

    while i < len(series) - 1:
        if pd.isna(series[i]) or pd.isna(series[i+1]):
            i += 1
            continue

        direction = 1 if series[i+1] > series[i] else -1 if series[i+1] < series[i] else 0
        if direction == 0:
            i += 1
            continue

        start = i
        count = 2
        j = i + 2

        while j < len(series):
            if pd.isna(series[j]):
                break

            if direction == 1 and series[j] > series[j-1]:
                count += 1
            elif direction == -1 and series[j] < series[j-1]:
                count += 1
            else:
                break

            j += 1

        if count >= TREND_MIN_RUN:
            trends.append(
                f"TREND ({'increasing' if direction==1 else 'decreasing'}) from {labels[start]} to {labels[j-1]}"
            )

        i = j

    return trends

# ---------------------------
# ASTRO (RESTORED PROPERLY)
# ---------------------------
def detect_astro(series, labels):
    astro = []

    for i in range(1, len(series)):
        if pd.notna(series[i]) and pd.notna(series[i-1]):
            if abs(series[i] - series[i-1]) >= ASTRO_THRESHOLD:
                astro.append(
                    f"Astronomical point at {labels[i]} (value={round(series[i]*100,1)}%)"
                )

    return astro

# ---------------------------
# STREAMLIT APP
# ---------------------------
st.set_page_config(layout="wide")
st.title("📊 RunCharts Dashboard QPSD SKMCH&RC")

file = st.file_uploader("Upload Excel")

if file:

    df = pd.read_excel(file)

    dept_col = st.selectbox("Department Column", df.columns)
    ind_col = st.selectbox("Indicator Column", df.columns)

    df["_dept"] = df[dept_col].astype(str).apply(clean_text_for_match)

    dept = st.selectbox("Select Department", df["_dept"].unique())

    df = df[df["_dept"] == dept]

    indicator = st.selectbox("Select Indicator", df[ind_col].astype(str))

    row = df[df[ind_col].astype(str) == indicator].iloc[0]

    cols = list(df.columns[2:])

    labels = [pretty_label(c) for c in cols]
    series = [to_float(row[c]) for c in cols]

    # LIMIT 18 ONLY
    labels = labels[-NUM_POINTS:]
    series = series[-NUM_POINTS:]

    median = get_center_line(series)

    # DISPLAY
    st.subheader(indicator)

    st.write("Median =", f"{round(median*100,1)}%" if not np.isnan(median) else "N/A")

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=labels, y=series, mode="lines+markers"))
    fig.add_hline(y=median)

    st.plotly_chart(fig, use_container_width=True)

    # ---------------------------
    # ANALYSIS OUTPUT (FIXED)
    # ---------------------------
    st.markdown("### 📌 Analysis Summary")

    shifts = detect_shift(series, median, labels)
    trends = detect_trend(series, labels)
    astro = detect_astro(series, labels)

    if shifts:
        st.markdown("**SHIFT:**")
        for s in shifts:
            st.write(s)

    if trends:
        st.markdown("**TREND:**")
        for t in trends:
            st.write(t)

    if astro:
        st.markdown("**ASTRONOMICAL POINTS:**")
        for a in astro:
            st.write(a)

    if not shifts and not trends and not astro:
        st.success("No Shift, Trend or Astronomical variation detected")
