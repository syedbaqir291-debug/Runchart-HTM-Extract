# runcharts_hospital_bi_portal.py
# HOSPITAL BI PORTAL VERSION (Streamlit + Plotly + HTML export)
# Features:
# - Department navigation (sidebar)
# - Indicator drill-down
# - KPI cards (RAG status)
# - Shift / Trend / Astronomical detection
# - Single HTML export portal
# - Crash-proof numeric handling

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import os
from datetime import datetime
import re

# ---------------------------
# CONFIG
# ---------------------------
NUM_POINTS = 18
ASTRO_THRESHOLD = 10
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ---------------------------
# SAFE NUMERIC ENGINE
# ---------------------------
def safe_numeric(series):
    out = []
    for v in series:
        try:
            out.append(float(v))
        except:
            out.append(np.nan)
    return out

# ---------------------------
# ANALYTICS
# ---------------------------
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
    while i < len(series)-1:
        if pd.isna(series[i]):
            i += 1
            continue
        direction = 1 if series[i+1] > series[i] else -1
        j = i
        while j < len(series)-1:
            if (direction==1 and series[j+1]>series[j]) or (direction==-1 and series[j+1]<series[j]):
                j += 1
            else:
                break
        if j-i>=5:
            trends.append((i,j))
        i=j
    return trends


def detect_astro(series, center):
    return [i for i,v in enumerate(series)
            if pd.notna(v) and abs(v-center)>=ASTRO_THRESHOLD]

# ---------------------------
# KPI STATUS
# ---------------------------
def kpi_status(value, center):
    if pd.isna(value) or pd.isna(center):
        return "GRAY"
    diff = abs(value-center)
    if diff < 5:
        return "GREEN"
    elif diff < 10:
        return "AMBER"
    return "RED"

# ---------------------------
# UI
# ---------------------------
st.set_page_config(layout="wide")
st.title("🏥 Hospital BI Portal - RunCharts Intelligence System")

file = st.file_uploader("Upload Excel", type=["xlsx"])

if file:

    path = os.path.join(UPLOAD_DIR, file.name)
    with open(path,"wb") as f:
        f.write(file.getbuffer())

    xl = pd.ExcelFile(path)

    sheet = st.selectbox("Select Sheet", xl.sheet_names)
    header = st.number_input("Header Row",1,10,1)

    df = pd.read_excel(path, sheet_name=sheet, header=header-1)
    df = df.replace([np.inf,-np.inf],np.nan)

    dept_col = st.selectbox("Department Column", df.columns)
    ind_col = st.selectbox("Indicator Column", df.columns)

    # ---------------------------
    # SIDEBAR BI NAVIGATION
    # ---------------------------
    st.sidebar.header("🏥 Departments")
    dept_list = df[dept_col].dropna().unique()

    selected_dept = st.sidebar.radio("Select Department", dept_list)

    df_d = df[df[dept_col]==selected_dept]

    st.subheader(f"📌 Department Dashboard: {selected_dept}")

    # ---------------------------
    # KPI CARDS
    # ---------------------------
    st.subheader("📊 KPI Summary")

    cols = st.columns(3)

    total = len(df_d)
    shifts = 0
    trends = 0
    astro = 0

    for _,row in df_d.iterrows():
        labels = list(df.columns[-NUM_POINTS:])
        series = safe_numeric([row[c] for c in labels])
        center = center_line(series)

        shifts += len(detect_shift(series, center))
        trends += len(detect_trend(series))
        astro += len(detect_astro(series, center))

    cols[0].metric("Indicators", total)
    cols[1].metric("Shift Events", shifts)
    cols[2].metric("Trend Events", trends)

    st.metric("Astronomical Points", astro)

    # ---------------------------
    # INDICATOR VIEW
    # ---------------------------
    st.subheader("📈 RunCharts")

    html_blocks = []

    for _,row in df_d.iterrows():

        labels = list(df.columns[-NUM_POINTS:])
        raw = [row[c] for c in labels]
        series = safe_numeric(raw)

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

        for s,e in shift:
            fig.add_vrect(x0=labels[s],x1=labels[e],fillcolor="red",opacity=0.15)

        for s,e in trend:
            fig.add_vrect(x0=labels[s],x1=labels[e],fillcolor="green",opacity=0.15)

        fig.add_trace(go.Scatter(
            x=[labels[i] for i in astro],
            y=[series[i] for i in astro],
            mode="markers",
            marker=dict(color="red",size=10),
            name="AP"
        ))

        fig.update_layout(title=str(row[ind_col]),height=350)

        st.plotly_chart(fig,use_container_width=True)

        html_blocks.append(fig.to_html(full_html=False,include_plotlyjs=False))

    # ---------------------------
    # SINGLE PORTAL EXPORT
    # ---------------------------
    full_html = f"""
    <html>
    <head>
    <script src='https://cdn.plot.ly/plotly-latest.min.js'></script>
    </head>
    <body>
    <h1>🏥 Hospital BI Portal - {selected_dept}</h1>
    {''.join(html_blocks)}
    </body>
    </html>
    """

    st.download_button("📥 Download BI Portal HTML",
                       full_html,
                       file_name=f"BI_Portal_{selected_dept}.html",
                       mime="text/html")

else:
    st.info("Upload file to start")

st.markdown("---\nOMAC Developer")
