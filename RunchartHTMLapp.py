# runcharts_streamlit_premium.py
# Minimal-change upgrade: Department button filtering + simple Plotly HTML preview + PPT download per department

import streamlit as st
import pandas as pd
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
import re
import difflib
import io
import os
from datetime import datetime
import plotly.graph_objects as go

# ---------------------------
# Config
# ---------------------------
NUM_POINTS = 18
CENTER_LINE_METHOD = "median"
ASTRO_THRESHOLD = 10
NON_DATA_COLS_LOWER = {
    "department", "indicator", "target", "benchmark/ category",
    "benchmark", "frequency"
}

UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ---------------------------
# Credentials
# ---------------------------
PREMIUM_LOGIN = "Pakistan@1947"
PREMIUM_PASSWORD = "Pakistan@1947"

# ---------------------------
# Helpers
# ---------------------------
def clean_text(val):
    if pd.isna(val): return ""
    return str(val).strip().lower()


def get_center(values):
    arr = np.array(values, dtype=float)
    arr = arr[~np.isnan(arr)]
    return float(np.nanmedian(arr)) if len(arr) else np.nan


def sanitize(name):
    return re.sub(r"[^A-Za-z0-9._-]", "_", str(name))


# ---------------------------
# APP
# ---------------------------
st.set_page_config(layout="wide")
st.title("📊 RunCharts Dashboard (Simple + PPT Export)")

uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

if uploaded_file:

    path = os.path.join(UPLOAD_DIR, uploaded_file.name)
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    xl = pd.ExcelFile(path)

    sheet = st.selectbox("Select Sheet", xl.sheet_names)
    header = st.number_input("Header Row", 1, 10, 1)

    df = pd.read_excel(path, sheet_name=sheet, header=header-1)
    df = df.replace([np.inf, -np.inf], np.nan)

    dept_col = st.selectbox("Department Column", df.columns)
    ind_col = st.selectbox("Indicator Column", df.columns)

    # ---------------------------
    # Department BUTTON VIEW
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

        st.subheader("📈 Interactive Preview (Plotly)")

        fig = go.Figure()

        for _, row in df_d.iterrows():
            labels = df.columns[-NUM_POINTS:]
            series = [row[c] for c in labels]

            fig.add_trace(go.Scatter(
                x=list(labels),
                y=series,
                mode="lines+markers",
                name=str(row[ind_col])
            ))

        st.plotly_chart(fig, use_container_width=True)

        # ---------------------------
        # PPT EXPORT
        # ---------------------------
        if st.button("Generate PPT for Selected Department"):

            ppt = Presentation()

            for _, row in df_d.iterrows():

                slide = ppt.slides.add_slide(ppt.slide_layouts[6])

                labels = df.columns[-NUM_POINTS:]
                series = [row[c] for c in labels]

                center = get_center(series)

                title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
                title.text_frame.text = str(row[ind_col])

                chart_data = CategoryChartData()
                chart_data.categories = labels.astype(str)
                chart_data.add_series("Value", series)

                slide.shapes.add_chart(
                    XL_CHART_TYPE.LINE_MARKERS,
                    Inches(0.6), Inches(1.3), Inches(9), Inches(4)
                , chart_data)

            buf = io.BytesIO()
            ppt.save(buf)
            buf.seek(0)

            st.download_button(
                "Download PPT",
                buf,
                file_name=f"RunCharts_{st.session_state.dept}.pptx"
            )

else:
    st.info("Upload Excel file")

st.markdown("---\nOMAC Developer")
