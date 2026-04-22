# runcharts_plotly_dashboard_final.py
# FULL REVAMP: PPTX → Interactive Plotly + SPC Logic + HTML Export

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import os
import re
from datetime import datetime

# ---------------------------
# CONFIG
# ---------------------------
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

NUM_POINTS = 18
SHIFT_MIN_RUN = 6
TREND_MIN_RUN = 5
ASTRO_THRESHOLD = 10

NON_DATA_COLS_LOWER = {
    "department", "indicator", "target",
    "benchmark/ category", "benchmark", "frequency"
}

# ---------------------------
# HELPERS
# ---------------------------
def sanitize_filename(name):
    return re.sub(r"[^A-Za-z0-9._-]", "_", str(name))


def clean(val):
    return "" if pd.isna(val) else str(val).strip().lower()


def detect_data_columns(df):
    return [
        c for c in df.columns
        if clean(c) not in NON_DATA_COLS_LOWER
    ]


def median_center(values):
    arr = np.array(values, dtype=float)
    arr = arr[~np.isnan(arr)]
    return None if len(arr) == 0 else np.median(arr)


# ---------------------------
# SHIFT (EXACT YOUR LOGIC)
# ---------------------------
def detect_shift(series, center, min_run=SHIFT_MIN_RUN):

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

        while j < n:
            if signs[j] == 0 or signs[j] != curr:
                break
            count += 1
            j += 1

        if count >= min_run:
            shifts.append((i, j - 1, curr))

        i = j

    return shifts


# ---------------------------
# TREND (EXACT YOUR LOGIC)
# ---------------------------
def detect_trend(series, min_run=TREND_MIN_RUN):

    vals = list(series)

    comp_vals = []
    comp_idx = []
    prev = None

    for idx, v in enumerate(vals):
        if pd.isna(v):
            prev = None
            continue
        if prev is None or v != prev:
            comp_vals.append(v)
            comp_idx.append(idx)
            prev = v

    trends = []
    i, m = 0, len(comp_vals)

    while i < m - 1:

        if comp_vals[i + 1] > comp_vals[i]:
            direction = 1
        elif comp_vals[i + 1] < comp_vals[i]:
            direction = -1
        else:
            i += 1
            continue

        j = i + 1

        while j < m and (
            (direction == 1 and comp_vals[j] > comp_vals[j - 1]) or
            (direction == -1 and comp_vals[j] < comp_vals[j - 1])
        ):
            j += 1

        if j - i >= min_run:
            trends.append((comp_idx[i], comp_idx[j - 1], direction))

        i = j

    return trends


# ---------------------------
# ASTONOMICAL POINTS (EXACT YOUR LOGIC)
# ---------------------------
def detect_astronomical(series, median_val, threshold=ASTRO_THRESHOLD):

    ast = []
    numeric_vals = [v for v in series if pd.notna(v)]
    if not numeric_vals:
        return ast

    max_val = max(numeric_vals)
    scale = 100.0 if max_val <= 1 else 1.0

    scaled_series = [v * scale for v in series]
    scaled_median = median_val * scale if max_val <= 1 else median_val

    for i in range(1, len(scaled_series)):
        cur = scaled_series[i]
        prev = scaled_series[i - 1]

        if pd.isna(cur) or pd.isna(prev):
            continue

        diff_prev = abs(cur - prev)
        diff_median = abs(cur - scaled_median)

        if diff_prev >= threshold and diff_median >= threshold:
            ast.append(i)

    return ast


# ---------------------------
# RUN CHART
# ---------------------------
def make_chart(df, dept, indicator, date_cols):

    dff = df[(df["Department"] == dept) & (df["Indicator"] == indicator)]
    if dff.empty:
        return None, None

    row = dff.iloc[0]

    # 👉 LATEST 18 POINTS ONLY
    date_cols = date_cols[-NUM_POINTS:]

    series = pd.to_numeric(row[date_cols], errors="coerce").values
    labels = [str(c) for c in date_cols]

    center = median_center(series)

    shifts = detect_shift(series, center)
    trends = detect_trend(series)
    astro = detect_astronomical(series, center)

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=labels,
        y=series,
        mode="lines+markers",
        name="Run Chart"
    ))

    if center is not None:
        fig.add_trace(go.Scatter(
            x=labels,
            y=[center] * len(series),
            mode="lines",
            name="Median",
            line=dict(dash="dash")
        ))

    # Highlight ASTRO points
    if astro:
        fig.add_trace(go.Scatter(
            x=[labels[i] for i in astro],
            y=[series[i] for i in astro],
            mode="markers",
            name="Astronomical",
            marker=dict(size=10)
        ))

    title = f"{dept} → {indicator}"

    if shifts:
        title += f" | SHIFT:{len(shifts)}"
    if trends:
        title += f" | TREND:{len(trends)}"
    if astro:
        title += f" | ASTRO:{len(astro)}"

    fig.update_layout(
        title=title,
        template="plotly_white",
        height=450
    )

    return fig, (shifts, trends, astro)


# ---------------------------
# STREAMLIT UI
# ---------------------------
st.set_page_config(page_title="RunChart Dashboard", layout="wide")
st.title("🏥📊 Run Chart Dashboard (Final SPC Version)")

file = st.file_uploader("Upload Excel", type=["xlsx", "xls"])

if file:

    path = os.path.join(UPLOAD_DIR, sanitize_filename(file.name))
    with open(path, "wb") as f:
        f.write(file.getbuffer())

    xl = pd.ExcelFile(path)
    sheet = st.selectbox("Sheet", xl.sheet_names)

    header = st.number_input("Header row", 1, 10, 1)

    df = pd.read_excel(path, sheet_name=sheet, header=header - 1)
    df = df.replace([np.inf, -np.inf], np.nan)

    df.columns = [str(c).strip() for c in df.columns]

    if "Department" not in df.columns or "Indicator" not in df.columns:
        st.error("Missing Department or Indicator column")
        st.stop()

    date_cols = detect_data_columns(df)

    # SIDEBAR
    dept = st.sidebar.selectbox("Department", sorted(df["Department"].unique()))
    ind = st.sidebar.selectbox(
        "Indicator",
        sorted(df[df["Department"] == dept]["Indicator"].unique())
    )

    fig, rules = make_chart(df, dept, ind, date_cols)

    if fig:
        st.plotly_chart(fig, use_container_width=True)

        shifts, trends, astro = rules

        st.markdown("### SPC Analysis")

        st.write(f"Shift events: {len(shifts)}")
        st.write(f"Trend events: {len(trends)}")
        st.write(f"Astronomical points: {len(astro)}")

    # ---------------------------
    # HTML EXPORT
    # ---------------------------
    st.markdown("---")

    if st.button("Export HTML Dashboard"):

        html_parts = []

        for d in df["Department"].unique():

            for i in df[df["Department"] == d]["Indicator"].unique():

                fig, _ = make_chart(df, d, i, date_cols)

                if fig:
                    html_parts.append(
                        f"<h2>{d} - {i}</h2>" +
                        fig.to_html(full_html=False, include_plotlyjs="cdn")
                    )

        html = f"""
        <html>
        <head><title>Run Chart Dashboard</title></head>
        <body>
        <h1>Leadership Run Chart Dashboard</h1>
        {''.join(html_parts)}
        </body>
        </html>
        """

        out = f"dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"

        with open(out, "w", encoding="utf-8") as f:
            f.write(html)

        st.success("HTML exported")
        st.download_button("Download HTML", open(out, "rb"), file_name=out)

else:
    st.info("Upload Excel to begin")
