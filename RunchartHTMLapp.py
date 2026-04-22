# runcharts_plotly_dashboard.py
# FULL REVAMP: PPTX → Plotly Interactive Dashboard

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import os
import re
from datetime import datetime

# ---------------------------
# Config
# ---------------------------
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

NON_DATA_COLS_LOWER = {
    "department", "indicator", "target", "benchmark/ category",
    "benchmark", "frequency"
}

# ---------------------------
# Helper Functions
# ---------------------------
def clean_text(val):
    if pd.isna(val):
        return ""
    return str(val).strip().lower()


def sanitize_filename(name):
    return re.sub(r"[^A-Za-z0-9._-]", "_", str(name))


def get_center_line(values):
    arr = np.array(values, dtype=float)
    arr = arr[~np.isnan(arr)]
    if len(arr) == 0:
        return None
    return np.median(arr)


def make_run_chart(df, dept, indicator, date_cols):
    dff = df[
        (df["Department"] == dept) &
        (df["Indicator"] == indicator)
    ]

    if dff.empty:
        return None

    row = dff.iloc[0]
    y_values = row[date_cols].values

    # Clean numeric
    y = pd.to_numeric(pd.Series(y_values), errors="coerce").values

    x_labels = [str(c) for c in date_cols]
    center_line = get_center_line(y)

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=x_labels,
        y=y,
        mode="lines+markers",
        name="Run Chart"
    ))

    # Center line
    if center_line is not None:
        fig.add_trace(go.Scatter(
            x=x_labels,
            y=[center_line] * len(x_labels),
            mode="lines",
            name="Center Line",
            line=dict(dash="dash")
        ))

    fig.update_layout(
        title=f"{dept} → {indicator}",
        template="plotly_white",
        height=450
    )

    return fig


def detect_data_columns(df):
    cols = []
    for c in df.columns:
        if clean_text(c) not in NON_DATA_COLS_LOWER:
            cols.append(c)
    return cols


# ---------------------------
# Streamlit UI
# ---------------------------
st.set_page_config(page_title="RunChart Dashboard", layout="wide")

st.title("🏥📊 Run Charts Interactive Dashboard (Plotly Version)")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:

    # Save file
    filename = sanitize_filename(uploaded_file.name)
    path = os.path.join(UPLOAD_DIR, filename)

    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Read Excel
    xl = pd.ExcelFile(path, engine="openpyxl")
    sheet = st.selectbox("Select Sheet", xl.sheet_names)

    header_row = st.number_input("Header Row (1-based)", 1, 10, 1)

    df = pd.read_excel(
        path,
        sheet_name=sheet,
        header=header_row - 1,
        engine="openpyxl"
    )

    df = df.replace([np.inf, -np.inf], np.nan)

    st.success("File loaded successfully")

    st.dataframe(df.head())

    # ---------------------------
    # Detect structure
    # ---------------------------
    df.columns = [str(c).strip() for c in df.columns]

    required_cols = ["Department", "Indicator"]

    if not all(col in df.columns for col in required_cols):
        st.error("Excel must contain 'Department' and 'Indicator' columns")
        st.stop()

    date_cols = detect_data_columns(df)

    # ---------------------------
    # Sidebar Navigation
    # ---------------------------
    st.sidebar.header("Navigation")

    dept_list = sorted(df["Department"].dropna().unique())
    selected_dept = st.sidebar.selectbox("Select Department", dept_list)

    ind_list = sorted(
        df[df["Department"] == selected_dept]["Indicator"].dropna().unique()
    )
    selected_indicator = st.sidebar.selectbox("Select Indicator", ind_list)

    # ---------------------------
    # Generate Chart
    # ---------------------------
    fig = make_run_chart(df, selected_dept, selected_indicator, date_cols)

    if fig:
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("No data available for this selection")

    # ---------------------------
    # Optional: Export FULL HTML Dashboard
    # ---------------------------
    st.markdown("---")
    st.subheader("📤 Export Dashboard")

    if st.button("Generate HTML Dashboard"):

        html_parts = []

        for dept in dept_list:
            for ind in df[df["Department"] == dept]["Indicator"].unique():

                fig = make_run_chart(df, dept, ind, date_cols)

                if fig:
                    html_parts.append(
                        f"<h2>{dept} - {ind}</h2>" +
                        fig.to_html(full_html=False, include_plotlyjs="cdn")
                    )

        full_html = f"""
        <html>
        <head>
            <title>Run Chart Dashboard</title>
        </head>
        <body>
            <h1>Leadership Run Chart Dashboard</h1>
            {''.join(html_parts)}
        </body>
        </html>
        """

        out_file = f"dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"

        with open(out_file, "w", encoding="utf-8") as f:
            f.write(full_html)

        st.success(f"HTML Dashboard created: {out_file}")
        st.download_button("Download HTML", open(out_file, "rb"), file_name=out_file)

else:
    st.info("Upload Excel file to start dashboard")
