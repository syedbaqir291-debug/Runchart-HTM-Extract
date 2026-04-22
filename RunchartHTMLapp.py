# runcharts_html_dashboard_safe.py

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import os
import re

# ---------------------------
# CONFIG
# ---------------------------
NUM_POINTS = 18
ASTRO_THRESHOLD = 10
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ---------------------------
# SAFE NUMERIC ENGINE (FIX FOR YOUR ERROR)
# ---------------------------
def safe_numeric(series):
    cleaned = []
    for v in series:
        try:
            cleaned.append(float(v))
        except:
            cleaned.append(np.nan)
    return cleaned

# ---------------------------
# HELPERS
# ---------------------------
def center_line(series):
    arr = np.array(series, dtype=float)
    arr = arr[~np.isnan(arr)]
    return np.nanmedian(arr) if len(arr) else np.nan


def detect_shift(series, center):
    flags = []
    for v in series:
        if pd.isna(v):
            flags.append(0)
        elif v > center:
            flags.append(1)
        elif v < center:
            flags.append(-1)
        else:
            flags.append(0)

    shifts = []
    i = 0
    while i < len(flags):
        if flags[i] == 0:
            i += 1
            continue

        j = i
        while j < len(flags) and flags[j] == flags[i]:
            j += 1

        if j - i >= 5:
            shifts.append((i, j))

        i = j

    return shifts


def detect_trend(series):
    trends = []
    i = 0

    while i < len(series) - 1:
        if pd.isna(series[i]):
            i += 1
            continue

        direction = 1 if series[i+1] > series[i] else -1
        j = i

        while j < len(series) - 1:
            if (direction == 1 and series[j+1] > series[j]) or \
               (direction == -1 and series[j+1] < series[j]):
                j += 1
            else:
                break

        if j - i >= 5:
            trends.append((i, j))

        i = j

    return trends


def detect_astro(series, center):
    return [
        i for i, v in enumerate(series)
        if pd.notna(v) and abs(v - center) >= ASTRO_THRESHOLD
    ]


# ---------------------------
# APP UI
# ---------------------------
st.title("📊 RunChart HTML Dashboard (Stable Version)")

file = st.file_uploader("Upload Excel", type=["xlsx"])

if file:

    path = os.path.join(UPLOAD_DIR, file.name)
    with open(path, "wb") as f:
        f.write(file.getbuffer())

    xl = pd.ExcelFile(path)

    sheet = st.selectbox("Select Sheet", xl.sheet_names)
    header = st.number_input("Header Row", 1, 10, 1)

    df = pd.read_excel(path, sheet_name=sheet, header=header-1)

    # FIX CRASH HERE
    df = df.replace([np.inf, -np.inf], np.nan)

    dept_col = st.selectbox("Department Column", df.columns)
    ind_col = st.selectbox("Indicator Column", df.columns)

    # ---------------------------
    # DEPARTMENT BUTTONS
    # ---------------------------
    st.subheader("Departments")

    depts = df[dept_col].dropna().unique()

    if "selected_dept" not in st.session_state:
        st.session_state.selected_dept = None

    cols = st.columns(4)

    for i, d in enumerate(depts):
        if cols[i % 4].button(str(d)):
            st.session_state.selected_dept = d

    # ---------------------------
    # FILTER DATA
    # ---------------------------
    if st.session_state.selected_dept:

        df_d = df[df[dept_col] == st.session_state.selected_dept]
        st.success(f"Selected Department: {st.session_state.selected_dept}")

        html_blocks = []

        for _, row in df_d.iterrows():

            labels = list(df.columns[-NUM_POINTS:])
            raw_series = [row[c] for c in labels]

            # 🔥 FIX CRASH: convert safely
            series = safe_numeric(raw_series)

            center = center_line(series)

            shift = detect_shift(series, center)
            trend = detect_trend(series)
            astro = detect_astro(series, center)

            fig = go.Figure()

            fig.add_trace(go.Scatter(
                x=labels,
                y=series,
                mode="lines+markers",
                name="Value"
            ))

            if not np.isnan(center):
                fig.add_hline(y=center, line_dash="dash", line_color="gray")

            # SHIFT (red zone)
            for s, e in shift:
                fig.add_vrect(
                    x0=labels[s],
                    x1=labels[e],
                    fillcolor="red",
                    opacity=0.15
                )

            # TREND (green zone)
            for s, e in trend:
                fig.add_vrect(
                    x0=labels[s],
                    x1=labels[e],
                    fillcolor="green",
                    opacity=0.15
                )

            # ASTRO POINTS
            fig.add_trace(go.Scatter(
                x=[labels[i] for i in astro],
                y=[series[i] for i in astro],
                mode="markers",
                marker=dict(color="red", size=10),
                name="Astronomical"
            ))

            fig.update_layout(
                title=str(row[ind_col]),
                height=400
            )

            html_blocks.append(
                fig.to_html(full_html=False, include_plotlyjs=False)
            )

        # ---------------------------
        # SINGLE HTML OUTPUT
        # ---------------------------
        full_html = f"""
        <html>
        <head>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        </head>
        <body>
        <h2>Department: {st.session_state.selected_dept}</h2>
        {"".join(html_blocks)}
        </body>
        </html>
        """

        st.subheader("📄 HTML Dashboard Preview")

        st.components.v1.html(full_html, height=800, scrolling=True)

        st.download_button(
            "Download FULL HTML Report",
            full_html,
            file_name=f"RunChart_{st.session_state.selected_dept}.html",
            mime="text/html"
        )

else:
    st.info("Upload Excel file to start")
