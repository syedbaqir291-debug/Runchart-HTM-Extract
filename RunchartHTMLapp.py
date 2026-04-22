# runcharts_streamlit_premium.py
# UPGRADED VERSION: Department buttons + filtered indicators + PPTX highlights (shift/trend/astro circles)

import streamlit as st
import pandas as pd
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
import re
import io
import os
from datetime import datetime

# ---------------------------
# CONFIG
# ---------------------------
NUM_POINTS = 18
CENTER_LINE_METHOD = "median"
ASTRO_THRESHOLD = 10

TITLE_BOX = (0.6, 0.4, 9.0, 0.9)
CHART_BOX = (0.6, 1.4, 9.0, 4.3)
NOTES_BOX = (0.6, 6.0, 9.0, 1.3)

UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ---------------------------
# HELPERS
# ---------------------------
def clean_text(val):
    if pd.isna(val): return ""
    return str(val).strip().lower()


def get_center(values):
    arr = np.array(values, dtype=float)
    arr = arr[~np.isnan(arr)]
    return float(np.nanmedian(arr)) if len(arr) else np.nan


def detect_shift(series, center):
    flags = []
    for v in series:
        if pd.isna(v): flags.append(0)
        elif v > center: flags.append(1)
        elif v < center: flags.append(-1)
        else: flags.append(0)

    runs = []
    i = 0
    while i < len(flags):
        if flags[i] == 0:
            i += 1
            continue
        j = i
        while j < len(flags) and flags[j] == flags[i]:
            j += 1
        if j - i >= 5:  # 5 consecutive = shift zone
            runs.append((i, j-1, flags[i]))
        i = j
    return runs


def detect_trend(series):
    runs = []
    i = 0
    while i < len(series)-1:
        if pd.isna(series[i]):
            i += 1
            continue
        direction = None
        if series[i+1] > series[i]: direction = 1
        elif series[i+1] < series[i]: direction = -1
        else:
            i += 1
            continue

        j = i
        while j < len(series)-1:
            if (direction == 1 and series[j+1] > series[j]) or (direction == -1 and series[j+1] < series[j]):
                j += 1
            else:
                break
        if j - i >= 5:
            runs.append((i, j, direction))
        i = j
    return runs


def detect_astro(series, center):
    return [i for i,v in enumerate(series)
            if not pd.isna(v) and abs(v-center) >= ASTRO_THRESHOLD]


def x_position(i):
    x0, _, w, _ = CHART_BOX
    step = w / NUM_POINTS
    return x0 + i * step


# ---------------------------
# STREAMLIT
# ---------------------------
st.set_page_config(layout="wide")
st.title("📊 RunCharts Enhanced Dashboard")

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
    # DEPARTMENT BUTTONS
    # ---------------------------
    st.subheader("Departments")
    depts = df[dept_col].dropna().unique()

    if "selected_dept" not in st.session_state:
        st.session_state.selected_dept = None

    cols = st.columns(4)
    for i,d in enumerate(depts):
        if cols[i%4].button(str(d)):
            st.session_state.selected_dept = d

    if st.session_state.selected_dept:
        st.success(f"Selected: {st.session_state.selected_dept}")

        df_d = df[df[dept_col] == st.session_state.selected_dept]

        ppt = Presentation()

        for _, row in df_d.iterrows():

            labels = df.columns[-NUM_POINTS:]
            series = [row[c] for c in labels]

            center = get_center(series)

            shift = detect_shift(series, center)
            trend = detect_trend(series)
            astro = detect_astro(series, center)

            slide = ppt.slides.add_slide(ppt.slide_layouts[6])

            # TITLE
            tx = slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(9), Inches(0.8))
            tx.text_frame.text = str(row[ind_col])

            # CHART
            chart_data = CategoryChartData()
            chart_data.categories = labels.astype(str)
            chart_data.add_series("Value", series)

            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.LINE_MARKERS,
                Inches(CHART_BOX[0]), Inches(CHART_BOX[1]), Inches(CHART_BOX[2]), Inches(CHART_BOX[3]),
                chart_data
            ).chart

            # SHIFT CIRCLES
            for s,e,dirc in shift:
                x1 = x_position(s)
                x2 = x_position(e)
                width = x2 - x1
                slide.shapes.add_shape(
                    MSO_SHAPE.OVAL,
                    Inches(x1), Inches(1.5),
                    Inches(width), Inches(2)
                ).line.color.rgb = RGBColor(255, 0, 0)

            # TREND CIRCLES
            for s,e,dirc in trend:
                x1 = x_position(s)
                x2 = x_position(e)
                slide.shapes.add_shape(
                    MSO_SHAPE.OVAL,
                    Inches(x1), Inches(1.5),
                    Inches(x2-x1), Inches(2)
                ).line.color.rgb = RGBColor(0, 128, 0)

            # ASTRO POINTS
            for i in astro:
                slide.shapes.add_shape(
                    MSO_SHAPE.OVAL,
                    Inches(x_position(i)), Inches(2.5),
                    Inches(0.3), Inches(0.3)
                ).fill.solid()

        buf = io.BytesIO()
        ppt.save(buf)
        buf.seek(0)

        st.download_button("Download PPT", buf, file_name="runcharts.pptx")

else:
    st.info("Upload file")

st.markdown("---\nOMAC Developer")
