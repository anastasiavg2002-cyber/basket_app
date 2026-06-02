# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import os
import plotly.express as px

from pipeline import run_pipeline


# =========================
# CONFIG
# =========================

st.set_page_config(
    page_title="Basket Analysis BI",
    layout="wide"
)

st.title("📊 Basket Analysis")


# =========================
# SIDEBAR RULES (INTERACTIVE)
# =========================

st.sidebar.header("Coloring Rules")

aff_threshold = st.sidebar.slider(
    "Affinity threshold",
    0.0, 5.0, 1.5, 0.1
)

high_trips = st.sidebar.slider(
    "High Trips threshold",
    0, 100, 30, 1
)

low_trips = st.sidebar.slider(
    "Low Trips threshold",
    0, 100, 15, 1
)


# =========================
# FILE UPLOADS
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
        st.error("Upload all required files")

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
            st.download_button(
                "Download results",
                f,
                file_name="results.zip"
            )


# =========================
# LOAD RESULTS
# =========================

st.divider()
st.header("📌 Interactive Analysis")


RESULTS_DIR = "results"

if not os.path.exists(RESULTS_DIR):
    st.info("Run pipeline first")
    st.stop()

files = [f for f in os.listdir(RESULTS_DIR) if f.endswith(".xlsx")]

if not files:
    st.info("No results found")
    st.stop()


# =========================
# SELECT FILE / SHEET
# =========================

selected_file = st.selectbox("Select project file", files)

file_path = os.path.join(RESULTS_DIR, selected_file)

xls = pd.ExcelFile(file_path)

selected_sheet = st.selectbox("Select sheet", xls.sheet_names)

df = pd.read_excel(
    xls,
    sheet_name=selected_sheet,
    header=1
)


# =========================
# CLEAN COLUMN NAMES SAFE
# =========================

df.columns = [str(c) for c in df.columns]


# =========================
# KPI SECTION (SAFE)
# =========================

st.subheader("📊 KPI Overview")

col1, col2, col3 = st.columns(3)


def safe_num(series):
    return pd.to_numeric(series, errors="coerce")


# Trips
trips_mean = None
affinity_mean = None

try:
    trips_col = "Trips (000)" if "Trips (000)" in df.columns else df.columns[3]
    trips_mean = safe_num(df[trips_col]).mean()
except:
    pass

try:
    affinity_col = "Affinity index to FMCG"
    affinity_mean = safe_num(df[affinity_col]).mean()
except:
    pass


# =========================
# CLASSIFICATION LOGIC
# =========================

def classify_row(row):

    try:
        aff = float(row.get("Affinity index to FMCG", None))
    except:
        aff = None

    try:
        trips = float(row.get("target_trips_raw_CS", None))
    except:
        trips = None

    if aff is not None and trips is not None:
        if aff >= aff_threshold and trips > high_trips:
            return "orange"

    if trips is not None and trips < low_trips:
        return "light_gray"

    if aff is not None and aff >= aff_threshold:
        return "gray"

    return "none"


df["color"] = df.apply(classify_row, axis=1)

orange_df = df[df["color"] == "orange"]


# =========================
# KPI DISPLAY
# =========================

with col1:
    st.metric("Avg Trips (000)", f"{trips_mean:.2f}" if trips_mean is not None else "N/A")

with col2:
    st.metric("Avg Affinity", f"{affinity_mean:.2f}" if affinity_mean is not None else "N/A")

with col3:
    orange_pct = len(orange_df) / len(df) * 100 if len(df) else 0
    st.metric("% Orange Rows", f"{orange_pct:.1f}%")


# =========================
# PREVIEW
# =========================

st.subheader("📄 Data Preview")
st.dataframe(df, use_container_width=True)


# =========================
# SCATTER PLOT (PLOTLY)
# =========================

st.subheader("📈 Interactive Scatter (Colored by Rules)")


numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()

if len(numeric_cols) < 2:
    st.warning("Not enough numeric columns")
    st.stop()


col1, col2 = st.columns(2)

with col1:
    x_axis = st.selectbox("X axis", numeric_cols)

with col2:
    y_axis = st.selectbox("Y axis", numeric_cols)


color_map = {
    "orange": "orange",
    "light_gray": "lightgray",
    "gray": "gray",
    "none": "blue"
}


fig = px.scatter(
    df,
    x=x_axis,
    y=y_axis,
    color="color",
    color_discrete_map=color_map,
    text=df.iloc[:, 0],  # column A labels
    title="Basket Scatter Plot (Interactive Rules)"
)

fig.update_traces(textposition="top center")

st.plotly_chart(fig, use_container_width=True)


# =========================
# DRILL DOWN
# =========================

st.subheader("🔍 Drill-down (Orange rows)")

if len(orange_df) > 0:

    selected_idx = st.selectbox("Select row", orange_df.index)

    st.dataframe(orange_df.loc[[selected_idx]], use_container_width=True)

else:
    st.info("No orange rows")


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

compare_files = st.multiselect("Select projects", files)

if len(compare_files) >= 2:

    combined = []

    for f in compare_files:

        xls_tmp = pd.ExcelFile(os.path.join(RESULTS_DIR, f))

        for sheet in xls_tmp.sheet_names:

            tmp = pd.read_excel(
                    xls_tmp,
                    sheet_name=sheet,
                    header=1
                )
            tmp["project"] = f
            combined.append(tmp)

    full_df = pd.concat(combined, ignore_index=True)

    num_cols = full_df.select_dtypes(include=["number"]).columns.tolist()

    if len(num_cols) >= 2:

        col1, col2 = st.columns(2)

        with col1:
            x_axis_cmp = st.selectbox("Compare X", num_cols, key="cmpx")

        with col2:
            y_axis_cmp = st.selectbox("Compare Y", num_cols, key="cmpy")

        fig2 = px.scatter(
            full_df,
            x=x_axis_cmp,
            y=y_axis_cmp,
            color="project",
            hover_data=["project"]
        )

        st.plotly_chart(fig2, use_container_width=True)
