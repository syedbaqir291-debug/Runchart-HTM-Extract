# runcharts_html_dashboard.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import re
import os
from datetime import datetime
import streamlit.components.v1 as components

# ---------------------------
# CONFIG
# ---------------------------
NUM_POINTS = 18
ASTRO_THRESHOLD = 10
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ---------------------------
# HELPERS
# ---------------------------
def clean(val):
    return str(val).strip().lower() if pd.notna(val) else ""

def center_line(series):
    arr = np.array(series, dtype=float)
    arr = arr[~np.isnan(arr)]
    return np.nanmedian(arr) if len(arr) else np.nan

def detect_shift(series, center):
    flags = []
    for v in series:
        if pd.isna(v): flags.append(0)
        elif v > center: flags.append(1)
        elif v < center: flags.append(-1)
        else: flags.append(0)

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

        while j < len(series)-1:
            if (direction == 1 and series[j+1] > series[j]) or (direction == -1 and series[j+1] < series[j]):
                j += 1
            else:
                break

        if j - i >= 5:
            trends.append((i, j))
        i = j
    return trends

def detect_astro(series, center):
    return [i for i,v in enumerate(series)
            if pd.notna(v) and abs(v-center) >= ASTRO_THRESHOLD]

# ---------------------------
# APP
# ---------------------------
st.title("📊 RunChart HTML Dashboard (Single File View)")

file = st.file_uploader("Upload Excel", type=["xlsx"])

if file:

    path = os.path.join(UPLOAD_DIR, file.name)
    with open(path, "wb") as f:
        f.write(file.getbuffer())

    xl = pd.ExcelFile(path)

    sheet = st.selectbox("Sheet", xl.sheet_names)
    header = st.number_input("Header row", 1, 10, 1)

    df = pd.read_excel(path, sheet_name=sheet, header=header-1)
    df = df.replace([np.inf, -np.inf], np.nan)

    dept_col = st.selectbox("Department column", df.columns)
    ind_col = st.selectbox("Indicator column", df.columns)

    # ---------------------------
    # DEPARTMENT BUTTONS
    # ---------------------------
    st.subheader("Departments")

    depts = df[dept_col].dropna().unique()

    if "dept" not in st.session_state:
        st.session_state.dept = None

    cols = st.columns(4)
    for i,d in enumerate(depts):
        if cols[i % 4].button(str(d)):
            st.session_state.dept = d

    # ---------------------------
    # FILTER DATA
    # ---------------------------
    if st.session_state.dept:
        df_d = df[df[dept_col] == st.session_state.dept]
        st.success(f"Selected: {st.session_state.dept}")

        html_blocks = []

        # ---------------------------
        # EACH INDICATOR = SEPARATE CHART
        # ---------------------------
        for _, row in df_d.iterrows():

            labels = df.columns[-NUM_POINTS:]
            series = [row[c] for c in labels]

            c = center_line(series)

            shift = detect_shift(series, c)
            trend = detect_trend(series)
            astro = detect_astro(series, c)

            fig = go.Figure()

            fig.add_trace(go.Scatter(
                x=labels,
                y=series,
                mode="lines+markers",
                name="Value"
            ))

            fig.add_hline(y=c, line_dash="dash", line_color="gray")

            # MARKERS
            for s,e in shift:
                fig.add_vrect(x0=labels[s], x1=labels[e],
                              fillcolor="red", opacity=0.2)

            for s,e in trend:
                fig.add_vrect(x0=labels[s], x1=labels[e],
                              fillcolor="green", opacity=0.2)

            for i in astro:
                fig.add_scatter(
                    x=[labels[i]],
                    y=[series[i]],
                    mode="markers",
                    marker=dict(color="red", size=10)
                )

            fig.update_layout(title=str(row[ind_col]))

            html_blocks.append(fig.to_html(full_html=False, include_plotlyjs=False))

        # ---------------------------
        # SINGLE HTML PAGE
        # ---------------------------
        full_html = f"""
        <html>
        <head>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        </head>
        <body>
        <h2>Department: {st.session_state.dept}</h2>
        {"".join(html_blocks)}
        </body>
        </html>
        """

        st.subheader("📄 Dashboard Preview")
        components.html(full_html, height=800, scrolling=True)

        st.download_button(
            "Download FULL HTML Report",
            full_html,
            file_name=f"RunChart_{st.session_state.dept}.html",
            mime="text/html"
        )

else:
    st.info("Upload file first")
