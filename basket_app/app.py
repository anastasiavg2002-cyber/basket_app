# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import os
import plotly.express as px

from pipeline import run_pipeline


st.set_page_config(
    page_title="Basket BI Dashboard",
    layout="wide"
)

st.title("📊 Basket Analysis BI Dashboard")


# =========================
# UPLOAD FILES
# =========================

projects_zip = st.file_uploader("Upload projects.zip", type=["zip"])
reference_file = st.file_uploader("Upload Список.xlsx", type=["xlsx"])
config_file = st.file_uploader("Upload sheet_config.xlsx", type=["xlsx"])
total_file = st.file_uploader("Upload total.xlsx", type=["xlsx"])


# =========================
# RUN PIPELINE
# =========================

if st.button("RUN PIPELINE"):

    if not projects_zip or not reference_file or not config_file or not total_file:
        st.error("Upload all files")

    else:

        with open("projects.zip", "wb") as f:
            f.write(projects_zip.getbuffer())

        with open("Список.xlsx", "wb") as f:
            f.write(reference_file.getbuffer())

        with open("sheet_config.xlsx", "wb") as f:
            f.write(config_file.getbuffer())

        with open("total.xlsx", "wb") as f:
            f.write(total_file.getbuffer())

        with st.spinner("Processing pipeline..."):
            output_zip = run_pipeline(
                "projects.zip",
                "Список.xlsx",
                "sheet_config.xlsx",
                "total.xlsx"
            )

        st.success("Done!")

        with open(output_zip, "rb") as f:
            st.download_button("Download Results", f, file_name="results.zip")


# =========================
# LOAD RESULTS
# =========================

st.divider()
st.header("📌 BI Dashboard (Results)")


RESULTS_DIR = "results"

if not os.path.exists(RESULTS_DIR):
    st.info("Run pipeline first")
    st.stop()

files = [f for f in os.listdir(RESULTS_DIR) if f.endswith(".xlsx")]

if not files:
    st.info("No results found")
    st.stop()


# =========================
# SELECT PROJECT / SHEET
# =========================

selected_file = st.selectbox("Select project", files)
file_path = os.path.join(RESULTS_DIR, selected_file)

xls = pd.ExcelFile(file_path)

selected_sheet = st.selectbox("Select sheet", xls.sheet_names)

df = pd.read_excel(xls, sheet_name=selected_sheet)


# =========================
# KPI DASHBOARD
# =========================

st.subheader("📊 KPI Overview")

# safe KPI calculation
col1, col2, col3 = st.columns(3)


def safe_numeric(series):
    return pd.to_numeric(series, errors="coerce")


# Trips
try:
    trips_col = df.columns[3]
    trips_mean = safe_numeric(df[trips_col]).mean()
except:
    trips_mean = None


# Affinity
try:
    affinity_col = df.columns[4]
    affinity_mean = safe_numeric(df[affinity_col]).mean()
except:
    affinity_mean = None


# Orange %
def is_orange(row):
    try:
        f_val = float(row.iloc[5])
        e_val = row.iloc[4]
        return (f_val > 30 and e_val >= 1.5)
    except:
        return False


orange_df = df[df.apply(is_orange, axis=1)]
orange_pct = len(orange_df) / len(df) * 100 if len(df) else 0


with col1:
    st.metric("Avg Trips (000)", f"{trips_mean:.2f}" if trips_mean is not None else "N/A")

with col2:
    st.metric("Avg Affinity", f"{affinity_mean:.2f}" if affinity_mean is not None else "N/A")

with col3:
    st.metric("% Orange Rows", f"{orange_pct:.1f}%")


# =========================
# PREVIEW TABLE
# =========================

st.subheader("📄 Data Preview")
st.dataframe(df, use_container_width=True)


# =========================
# SCATTER PLOT (PLOTLY)
# =========================

st.subheader("📈 Interactive Scatter Plot (Orange Rows)")


if len(orange_df) == 0:
    st.warning("No orange rows found")
    st.stop()


numeric_cols = orange_df.select_dtypes(include=["number"]).columns.tolist()

if len(numeric_cols) < 2:
    st.warning("Not enough numeric columns")
    st.stop()


col1, col2 = st.columns(2)

with col1:
    x_axis = st.selectbox("X axis", numeric_cols)

with col2:
    y_axis = st.selectbox("Y axis", numeric_cols)


fig = px.scatter(
    orange_df,
    x=x_axis,
    y=y_axis,
    text=orange_df.iloc[:, 0],  # column A labels
    title="Orange rows scatter"
)

fig.update_traces(textposition="top center")


selected_points = st.plotly_chart(fig, use_container_width=True)


# =========================
# CLICK FILTER (SIMULATION)
# =========================

st.subheader("🔍 Drill-down")

st.write("Select row manually:")

selected_index = st.selectbox(
    "Choose row index",
    orange_df.index
)

st.dataframe(orange_df.loc[[selected_index]], use_container_width=True)


# =========================
# DOWNLOAD FILTERED DATA
# =========================

csv = orange_df.to_csv(index=False).encode("utf-8")

st.download_button(
    "Download orange rows",
    csv,
    file_name="orange_rows.csv",
    mime="text/csv"
)


# =========================
# MULTI-PROJECT COMPARISON
# =========================

st.subheader("📊 Multi-project comparison")

compare_files = st.multiselect("Select projects to compare", files)

if len(compare_files) >= 2:

    combined = []

    for f in compare_files:

        xls = pd.ExcelFile(os.path.join(RESULTS_DIR, f))

        for sheet in xls.sheet_names:

            tmp = pd.read_excel(xls, sheet_name=sheet)
            tmp["project"] = f
            combined.append(tmp)

    full_df = pd.concat(combined, ignore_index=True)

    num_cols = full_df.select_dtypes(include=["number"]).columns.tolist()

    if len(num_cols) >= 2:

        x_axis = st.selectbox("Compare X", num_cols, key="cmpx")
        y_axis = st.selectbox("Compare Y", num_cols, key="cmpy")

        fig2 = px.scatter(
            full_df,
            x=x_axis,
            y=y_axis,
            color="project",
            hover_data=["project"]
        )

        st.plotly_chart(fig2, use_container_width=True)
