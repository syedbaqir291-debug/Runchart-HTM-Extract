```python
# app.py
# GitHub-ready Streamlit App
# Premium Run Chart HTML Dashboard Generator
# Footer: OMAC Developers by S M Baqir

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import re
import io
from datetime import datetime

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------

st.set_page_config(
    page_title="Run Charts Premium Dashboard",
    page_icon="📊",
    layout="wide"
)

# --------------------------------------------------
# PREMIUM UI CSS
# --------------------------------------------------

st.markdown("""
<style>
.main {
    background-color: #f8fafc;
}

.stButton > button {
    width: 100%;
    border-radius: 10px;
    height: 45px;
    font-weight: 600;
}

.department-box {
    padding: 12px;
    border-radius: 12px;
    background: white;
    margin-bottom: 15px;
    box-shadow: 0px 2px 8px rgba(0,0,0,0.08);
}

.indicator-box {
    padding: 15px;
    border-radius: 12px;
    background: #ffffff;
    margin-bottom: 20px;
    border-left: 5px solid #2563eb;
    box-shadow: 0px 2px 8px rgba(0,0,0,0.05);
}

.footer {
    position: fixed;
    left: 0;
    bottom: 0;
    width: 100%;
    background-color: #ffffff;
    color: gray;
    text-align: center;
    padding: 10px;
    font-size: 12px;
    border-top: 1px solid #e5e7eb;
}
</style>
""", unsafe_allow_html=True)

# --------------------------------------------------
# HELPER FUNCTIONS
# --------------------------------------------------

def pretty_label(col):
    try:
        parsed = pd.to_datetime(col, errors="coerce")
        if not pd.isna(parsed):
            return parsed.strftime("%b-%y")
    except:
        pass
    return str(col)

def get_center_line(values):
    arr = np.array(values, dtype=float)
    arr = arr[~np.isnan(arr)]
    if len(arr) == 0:
        return np.nan
    return float(np.nanmedian(arr))

def clean_text(val):
    if pd.isna(val):
        return ""
    return str(val).strip()

def create_run_chart(labels, values, median_value, title):
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=labels,
        y=values,
        mode="lines+markers",
        name="Indicator Value"
    ))

    fig.add_trace(go.Scatter(
        x=labels,
        y=[median_value] * len(labels),
        mode="lines",
        name="Median Line",
        line=dict(dash="dash")
    ))

    fig.update_layout(
        title=title,
        height=420,
        template="plotly_white",
        margin=dict(l=20, r=20, t=60, b=20)
    )

    return fig

# --------------------------------------------------
# HEADER
# --------------------------------------------------

st.title("🏥 Premium Run Charts Dashboard")
st.caption("Interactive Department-wise Run Chart Automation")

# --------------------------------------------------
# FILE UPLOAD
# --------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)

        sheet_name = st.selectbox(
            "Select Sheet",
            excel_file.sheet_names
        )

        header_row = st.number_input(
            "Header Row Number",
            min_value=1,
            value=1,
            step=1
        )

        df = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name,
            header=header_row - 1
        )

        st.success("Excel Loaded Successfully")

        col1, col2 = st.columns(2)

        with col1:
            dept_col = st.selectbox(
                "Select Department Column",
                df.columns
            )

        with col2:
            indicator_col = st.selectbox(
                "Select Indicator Column",
                df.columns,
                index=1 if len(df.columns) > 1 else 0
            )

        non_data_cols = {
            dept_col,
            indicator_col
        }

        data_cols = [
            c for c in df.columns
            if c not in non_data_cols
        ]

        departments = sorted(
            df[dept_col]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        st.divider()

        st.subheader("Departments")

        selected_department = st.selectbox(
            "Select Department",
            ["ALL"] + departments
        )

        if st.button("Generate Dashboard"):

            target_df = df.copy()

            if selected_department != "ALL":
                target_df = target_df[
                    target_df[dept_col].astype(str) == selected_department
                ]

            grouped_departments = sorted(
                target_df[dept_col]
                .dropna()
                .astype(str)
                .unique()
            )

            for dept in grouped_departments:
                dept_df = target_df[
                    target_df[dept_col].astype(str) == dept
                ]

                st.markdown(f"""
                <div class="department-box">
                    <h2>{dept}</h2>
                    <p>Total Indicators: {len(dept_df)}</p>
                </div>
                """, unsafe_allow_html=True)

                for _, row in dept_df.iterrows():
                    indicator_name = clean_text(
                        row[indicator_col]
                    ) or "Indicator"

                    values = []
                    labels = []

                    for col in data_cols:
                        val = pd.to_numeric(
                            row[col],
                            errors="coerce"
                        )

                        labels.append(
                            pretty_label(col)
                        )
                        values.append(val)

                    median_value = get_center_line(values)

                    st.markdown(f"""
                    <div class="indicator-box">
                        <h4>{indicator_name}</h4>
                        <p><strong>Median:</strong> {round(median_value, 2) if not pd.isna(median_value) else "N/A"}</p>
                    </div>
                    """, unsafe_allow_html=True)

                    fig = create_run_chart(
                        labels,
                        values,
                        median_value,
                        indicator_name
                    )

                    st.plotly_chart(
                        fig,
                        use_container_width=True
                    )

        st.divider()

    except Exception as e:
        st.error(f"Error: {str(e)}")

else:
    st.info("Please upload your Excel file to begin.")

# --------------------------------------------------
# FOOTER
# --------------------------------------------------

st.markdown("""
<div class="footer">
    OMAC Developers by S M Baqir
</div>
""", unsafe_allow_html=True)
```
