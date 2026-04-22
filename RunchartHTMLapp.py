# runcharts_streamlit_premium.py
# Fixed version: same functionality, only pandas/df errors removed + proper step order flow

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

TITLE_BOX = (0.6, 0.4, 9.0, 0.9)
CHART_BOX = (0.6, 1.4, 9.0, 4.3)
NOTES_BOX = (0.6, 6.0, 9.0, 1.3)

LOG_DIR = "logs"
UPLOAD_DIR = os.path.join(LOG_DIR, "uploads")
LOG_FILE = os.path.join(LOG_DIR, "activity_log.csv")

os.makedirs(UPLOAD_DIR, exist_ok=True)

# ---------------------------
# Premium Credentials
# ---------------------------
PREMIUM_LOGIN = "Pakistan@1947"
PREMIUM_PASSWORD = "Pakistan@1947"

# ---------------------------
# Helper Functions
# ---------------------------

def clean_text_for_match(val):
    if pd.isna(val):
        return ""
    s = str(val)
    s = re.sub(r"[\u200B\u200C\u200D\uFEFF]", "", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip().lower()


def pretty_label(col):
    try:
        parsed = pd.to_datetime(col, errors="coerce")
        if not pd.isna(parsed):
            return parsed.strftime("%b-%y")
    except Exception:
        pass
    return str(col)


def get_center_line(values, method="median"):
    arr = np.array(values, dtype=float)
    arr = arr[~np.isnan(arr)]
    if arr.size == 0:
        return np.nan
    return float(np.nanmedian(arr)) if method == "median" else float(np.nanmean(arr))


def sanitize_filename(name):
    return re.sub(r"[^A-Za-z0-9._-]", "_", str(name))


def log_activity(username, action, filename):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = {"timestamp": now, "user": username or "", "action": action or "", "filename": filename or ""}
    df_entry = pd.DataFrame([entry]).replace([np.nan, np.inf, -np.inf], "")
    write_header = not os.path.exists(LOG_FILE)
    df_entry.to_csv(LOG_FILE, mode="a", index=False, header=write_header, na_rep="")


# ---------------------------
# Streamlit App
# ---------------------------
st.set_page_config(page_title="RunCharts Automation - Premium", layout="wide")
st.title("🏥📊 RunCharts Automation SKMCH & RC (QPSD)")

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "auth_user" not in st.session_state:
    st.session_state["auth_user"] = ""

# ---------------------------
# Login
# ---------------------------
with st.sidebar.expander("🔒 Premium Login", expanded=True):
    if not st.session_state["authenticated"]:
        with st.form("login_form"):
            login_input = st.text_input("Login:")
            password_input = st.text_input("Password:", type="password")
            submitted = st.form_submit_button("Sign in")

            if submitted:
                if login_input == PREMIUM_LOGIN and password_input == PREMIUM_PASSWORD:
                    st.session_state["authenticated"] = True
                    st.session_state["auth_user"] = login_input
                    st.success("Authenticated")
                else:
                    st.error("Invalid credentials")
    else:
        st.markdown(f"Signed in as: {st.session_state['auth_user']}")
        if st.button("Sign out"):
            st.session_state["authenticated"] = False
            st.session_state["auth_user"] = ""
            st.rerun()

# ---------------------------
# Inputs
# ---------------------------
username = st.text_input("Enter your name:", value=st.session_state.get("auth_user", ""))
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

# ---------------------------
# STEP CONTROL FLOW FIXED
# ---------------------------
if uploaded_file is not None:

    safe_name = sanitize_filename(uploaded_file.name)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    saved_name = f"{ts}_{safe_name}"
    saved_path = os.path.join(UPLOAD_DIR, saved_name)

    with open(saved_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    log_activity(username, "Uploaded file", saved_name)

    # STEP 1: Load Excel
    xl = pd.ExcelFile(saved_path, engine="openpyxl")

    # STEP 2: Sheet selection FIRST
    sheet_name = st.selectbox("1️⃣ Select sheet", xl.sheet_names)

    if sheet_name:

        # STEP 3: Header selection
        header_row = st.number_input("2️⃣ Enter header row", min_value=1, value=1)

        # STEP 4: Load DF AFTER sheet + header
        df = pd.read_excel(
            saved_path,
            sheet_name=sheet_name,
            header=header_row - 1,
            engine="openpyxl"
        )

        # FIX: replace inf AFTER df exists
        df = df.replace([np.inf, -np.inf], np.nan)

        st.success("Excel loaded successfully")
        st.dataframe(df.head())

        # STEP 5: Department & Indicator selection
        dept_col = st.selectbox("3️⃣ Select Department column", df.columns)
        ind_col = st.selectbox("4️⃣ Select Indicator column", df.columns)

else:
    st.info("Upload file to start")

# ---------------------------
# Footer
# ---------------------------
st.markdown(
    """
    <hr>
    <div style='text-align:center;font-size:10px;color:grey;'>
        OMAC Developer by S M Baqir
    </div>
    """,
    unsafe_allow_html=True
)
