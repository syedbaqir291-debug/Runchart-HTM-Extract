import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import re

# -----------------------------
# CONFIG
# -----------------------------
NUM_POINTS = 18
ASTRO_THRESHOLD = 10

st.set_page_config(
    page_title="RunChart HTML Generator",
    layout="wide"
)

st.title("📊 RunChart HTML Report Generator (Premium)")

# -----------------------------
# HELPERS
# -----------------------------
def clean(val):
    if pd.isna(val):
        return ""
    return str(val).strip()

def median_line(values):
    arr = np.array(values, dtype=float)
    arr = arr[~np.isnan(arr)]
    if len(arr) == 0:
        return np.nan
    return np.median(arr)

def detect_shift(values, center):
    if np.isnan(center):
        return []
    flags = []

    for v in values:
        if pd.isna(v):
            flags.append(0)
        elif v > center:
            flags.append(1)
        else:
            flags.append(-1)

    shifts = []
    run = 1

    for i in range(1, len(flags)):
        if flags[i] == flags[i - 1]:
            run += 1
        else:
            if run >= 6:
                shifts.append(i - 1)
            run = 1

    return shifts

def detect_trend(values):
    trends = []
    run = 1

    for i in range(1, len(values)):
        if pd.isna(values[i]) or pd.isna(values[i - 1]):
            continue

        if values[i] > values[i - 1]:
            run += 1
        else:
            if run >= 5:
                trends.append(i - 1)
            run = 1

    return trends

def detect_astro(values, median):
    astro = []
    if np.isnan(median):
        return astro

    for i in range(1, len(values)):
        if pd.isna(values[i]) or pd.isna(values[i - 1]):
            continue

        if abs(values[i] - median) >= ASTRO_THRESHOLD and abs(values[i] - values[i - 1]) >= ASTRO_THRESHOLD:
            astro.append(i)

    return astro

# -----------------------------
# HTML BUILDER
# -----------------------------
def build_chart_html(labels, values, median, title, shift, trend, astro):

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=labels,
        y=values,
        mode="lines+markers",
        name="Value"
    ))

    if not np.isnan(median):
        fig.add_trace(go.Scatter(
            x=labels,
            y=[median] * len(labels),
            mode="lines",
            name="Median",
            line=dict(dash="dash")
        ))

    chart_html = fig.to_html(full_html=False, include_plotlyjs="cdn")

    summary = f"""
    <div style="font-family:Arial; padding:10px; border:1px solid #ddd; margin-bottom:10px;">
        <h3>{title}</h3>
        <p><b>Median:</b> {round(median,2) if not np.isnan(median) else 'N/A'}</p>
        <p><b>Shift Points:</b> {len(shift)}</p>
        <p><b>Trend Points:</b> {len(trend)}</p>
        <p><b>Astronomical Points:</b> {len(astro)}</p>
    </div>
    """

    return summary + chart_html

# -----------------------------
# FILE UPLOAD
# -----------------------------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    st.success("File Loaded Successfully")

    dept_col = st.selectbox("Select Department Column", df.columns)
    ind_col = st.selectbox("Select Indicator Column", df.columns)

    data_cols = [c for c in df.columns if c not in [dept_col, ind_col]]

    departments = sorted(df[dept_col].dropna().astype(str).unique())

    selected_dept = st.selectbox("Select Department", ["ALL"] + departments)

    if st.button("Generate HTML Report"):

        filtered = df.copy()

        if selected_dept != "ALL":
            filtered = filtered[filtered[dept_col].astype(str) == selected_dept]

        html = """
        <html>
        <head>
        <title>RunChart Report</title>
        <style>
            body {font-family: Arial; margin: 20px; background:#f8fafc;}
            .dept {background:#e2e8f0; padding:10px; margin-top:20px;}
            .indicator {margin-bottom:40px;}
        </style>
        </head>
        <body>
        <h1>📊 RunChart Quality Report</h1>
        """

        for dept in sorted(filtered[dept_col].astype(str).unique()):

            html += f"<div class='dept'><h2>Department: {dept}</h2></div>"

            dept_df = filtered[filtered[dept_col].astype(str) == dept]

            for _, row in dept_df.iterrows():

                title = clean(row[ind_col]) or "Indicator"

                values = []
                labels = []

                for c in data_cols[:NUM_POINTS]:
                    labels.append(str(c))
                    values.append(pd.to_numeric(row[c], errors="coerce"))

                median = median_line(values)

                shift = detect_shift(values, median)
                trend = detect_trend(values)
                astro = detect_astro(values, median)

                html += "<div class='indicator'>"
                html += build_chart_html(
                    labels,
                    values,
                    median,
                    title,
                    shift,
                    trend,
                    astro
                )
                html += "</div>"

        html += """
        </body>
        </html>
        """

        st.download_button(
            label="⬇ Download HTML Report",
            data=html,
            file_name="RunChart_Report.html",
            mime="text/html"
        )

else:
    st.info("Upload Excel file to start")
