import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

# -----------------------------
# PAGE CONFIG
# -----------------------------
st.set_page_config(
    page_title="Premium Run Chart Dashboard",
    page_icon="📊",
    layout="wide"
)

# -----------------------------
# PREMIUM UI
# -----------------------------
st.markdown("""
<style>
body {
    background-color: #f8fafc;
}

.department-box {
    padding: 12px;
    border-radius: 10px;
    background: white;
    margin-bottom: 15px;
    box-shadow: 0px 2px 8px rgba(0,0,0,0.08);
}

.indicator-box {
    padding: 12px;
    border-radius: 10px;
    background: #ffffff;
    margin-bottom: 15px;
    border-left: 4px solid #2563eb;
    box-shadow: 0px 2px 6px rgba(0,0,0,0.05);
}

.footer {
    position: fixed;
    bottom: 0;
    width: 100%;
    text-align: center;
    font-size: 12px;
    padding: 8px;
    background: white;
    color: gray;
    border-top: 1px solid #e5e7eb;
}
</style>
""", unsafe_allow_html=True)

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

def pretty(col):
    return str(col)

def run_chart(x, y, title, median):
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=x,
        y=y,
        mode="lines+markers",
        name="Value"
    ))

    fig.add_trace(go.Scatter(
        x=x,
        y=[median] * len(x),
        mode="lines",
        name="Median",
        line=dict(dash="dash")
    ))

    fig.update_layout(
        title=title,
        template="plotly_white",
        height=420,
        margin=dict(l=20, r=20, t=50, b=20)
    )

    return fig

# -----------------------------
# HEADER
# -----------------------------
st.title("🏥 Premium Run Chart Dashboard")
st.caption("Interactive Quality Indicator Monitoring System")

# -----------------------------
# UPLOAD FILE
# -----------------------------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file:

    excel = pd.ExcelFile(uploaded_file)

    sheet = st.selectbox("Select Sheet", excel.sheet_names)

    header_row = st.number_input("Header Row Number", 1, value=1)

    df = pd.read_excel(uploaded_file, sheet_name=sheet, header=header_row - 1)

    st.success("File Loaded Successfully")

    col1, col2 = st.columns(2)

    with col1:
        dept_col = st.selectbox("Department Column", df.columns)

    with col2:
        ind_col = st.selectbox("Indicator Column", df.columns)

    data_cols = [c for c in df.columns if c not in [dept_col, ind_col]]

    departments = sorted(df[dept_col].dropna().astype(str).unique())

    st.divider()

    selected_dept = st.selectbox("Select Department", ["ALL"] + departments)

    if st.button("Generate Dashboard"):

        filtered = df.copy()

        if selected_dept != "ALL":
            filtered = filtered[filtered[dept_col].astype(str) == selected_dept]

        dept_list = sorted(filtered[dept_col].dropna().astype(str).unique())

        for dept in dept_list:

            dept_df = filtered[filtered[dept_col].astype(str) == dept]

            st.markdown(f"""
            <div class="department-box">
                <h2>{dept}</h2>
                <p>Total Indicators: {len(dept_df)}</p>
            </div>
            """, unsafe_allow_html=True)

            for _, row in dept_df.iterrows():

                indicator = clean(row[ind_col]) or "Indicator"

                values = []
                labels = []

                for c in data_cols:
                    labels.append(pretty(c))
                    values.append(pd.to_numeric(row[c], errors="coerce"))

                med = median_line(values)

                st.markdown(f"""
                <div class="indicator-box">
                    <h4>{indicator}</h4>
                    <p><b>Median:</b> {round(med,2) if not pd.isna(med) else "N/A"}</p>
                </div>
                """, unsafe_allow_html=True)

                fig = run_chart(labels, values, indicator, med)

                st.plotly_chart(fig, use_container_width=True)

# -----------------------------
# FOOTER
# -----------------------------
st.markdown("""
<div class="footer">
    OMAC Developers by S M Baqir
</div>
""", unsafe_allow_html=True)
