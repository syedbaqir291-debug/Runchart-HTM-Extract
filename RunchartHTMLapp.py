# runcharts_streamlit_premium.py
# FIXED VERSION (NO LOGIC CHANGE + DATETIME SAFE)

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import re

# ---------------------------
# CONFIG
# ---------------------------
NUM_POINTS = 18
SHIFT_MIN_RUN = 6
TREND_MIN_RUN = 5
ASTRO_THRESHOLD = 0.1

# ---------------------------
# CLEANING (FIXED SAFE VERSION)
# ---------------------------
def clean_text_for_match(val):
    if pd.isna(val):
        return ""

    # FIX: force safe conversion for ALL types (datetime, int, float, etc.)
    s = str(val)

    s = re.sub(r"[\u200B\u200C\u200D\uFEFF]", "", s)
    s = re.sub(r"\s+", " ", s)

    return s.strip().lower()


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
# TARGET PARSER
# ---------------------------
def parse_target(val):
    if pd.isna(val):
        return np.nan

    val = str(val)
    num = re.findall(r"\d+\.?\d*", val)

    if len(num) == 0:
        return np.nan

    return float(num[0])

# ---------------------------
# CENTER LINE
# ---------------------------
def get_center_line(values):
    arr = np.array([v for v in values if pd.notna(v)], dtype=float)

    if len(arr) == 0:
        return np.nan

    return np.nanmedian(arr)

# ---------------------------
# SHIFT (UNCHANGED LOGIC)
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
# TREND (UNCHANGED LOGIC)
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
# ASTRO (UNCHANGED LOGIC)
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
# APP
# ---------------------------
st.title("RunChart Dashboard (Stable Fixed Version)")

file = st.file_uploader("Upload Excel")

if file:

    df = pd.read_excel(file)

    dept_col = st.selectbox("Department Column", df.columns)
    ind_col = st.selectbox("Indicator Column", df.columns)
    target_col = st.selectbox("Target Column", df.columns)

    # FIX: safe apply (prevents datetime crash)
    df["_dept"] = df[dept_col].apply(clean_text_for_match)

    dept = st.selectbox("Select Department", df["_dept"].unique())
    df = df[df["_dept"] == dept]

    indicator = st.selectbox("Select Indicator", df[ind_col].astype(str))

    row = df[df[ind_col].astype(str) == indicator].iloc[0]

    cols = list(df.columns[2:])

    labels = [pretty_label(c) for c in cols]
    series = [to_float(row[c]) for c in cols]

    # LIMIT 18 POINTS
    labels = labels[-NUM_POINTS:]
    series = series[-NUM_POINTS:]

    median = get_center_line(series)
    target = parse_target(row[target_col])

    # ---------------------------
    # PLOT (NO LOGIC CHANGE)
    # ---------------------------
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=labels,
        y=series,
        mode="lines+markers",
        name="Value"
    ))

    fig.add_hline(y=median, line_dash="dash", line_color="gray")

    if not np.isnan(target):
        fig.add_hline(y=target/100, line_color="green", line_width=2)

    st.plotly_chart(fig, use_container_width=True)

    # ---------------------------
    # ANALYSIS (UNCHANGED LOGIC)
    # ---------------------------
    st.markdown("### Analysis")

    shifts = detect_shift(series, median, labels)
    trends = detect_trend(series, labels)
    astro = detect_astro(series, labels)

    for s in shifts:
        st.write(s)

    for t in trends:
        st.write(t)

    for a in astro:
        st.write(a)

    st.success(f"Median = {round(median*100,1)}%")

    if not np.isnan(target):
        st.success(f"Target = {target}%")
