# runcharts_streamlit_premium.py
# FIX ONLY SHIFT & TREND ACCORDING TO ORIGINAL LOGIC

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import re

# ---------------------------
# CONFIG (UNCHANGED)
# ---------------------------
NUM_POINTS = 18
SHIFT_MIN_RUN = 6   # ✅ ORIGINAL RULE RESTORED
TREND_MIN_RUN = 5   # ✅ ORIGINAL RULE RESTORED

# ---------------------------
# CLEAN
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
        return float(str(v).replace("%",""))
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
# SHIFT (FIXED ORIGINAL LOGIC)
# ---------------------------
def detect_shift(series, center, labels):
    shifts = []
    i = 0
    n = len(series)

    while i < n:
        if pd.isna(series[i]):
            i += 1
            continue

        direction = 1 if series[i] > center else -1
        start = i
        count = 1
        j = i + 1

        while j < n:
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
# TREND (FIXED ORIGINAL LOGIC)
# ---------------------------
def detect_trend(series, labels):
    trends = []
    i = 0
    n = len(series)

    while i < n - 1:
        if pd.isna(series[i]):
            i += 1
            continue

        direction = None
        if series[i+1] > series[i]:
            direction = 1
        elif series[i+1] < series[i]:
            direction = -1
        else:
            i += 1
            continue

        start = i
        count = 2
        j = i + 2

        while j < n:
            if pd.isna(series[j]):
                break

            if direction == 1 and series[j] > series[j-1]:
                count += 1
                j += 1
            elif direction == -1 and series[j] < series[j-1]:
                count += 1
                j += 1
            else:
                break

        if count >= TREND_MIN_RUN:
            trends.append(
                f"TREND ({'increasing' if direction==1 else 'decreasing'}) from {labels[start]} to {labels[j-1]}"
            )

        i = j

    return trends

# ---------------------------
# ASTRO (UNCHANGED)
# ---------------------------
def detect_astro(series, labels):
    astro = []
    for i in range(1, len(series)):
        if pd.notna(series[i]) and pd.notna(series[i-1]):
            if abs(series[i] - series[i-1]) >= 0.1:
                astro.append(
                    f"Astronomical point at {labels[i]} (value={round(series[i]*100,1)}%)"
                )
    return astro

# ---------------------------
# STREAMLIT
# ---------------------------
st.title("RunChart Dashboard (Fixed Logic Only)")

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

    median = get_center_line(series)

    # LIMIT 18 ONLY
    labels = labels[-NUM_POINTS:]
    series = series[-NUM_POINTS:]

    st.subheader(indicator)

    st.write("Median =", f"{round(median*100,1)}%")

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=labels, y=series, mode="lines+markers"))
    fig.add_hline(y=median)

    st.plotly_chart(fig, use_container_width=True)

    # ---------------------------
    # ANALYSIS (FIXED)
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
