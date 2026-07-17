# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import os
import zipfile
import tempfile
from pathlib import Path

import plotly.express as px

from pipeline import run_pipeline


# =========================================================
# CONFIG
# =========================================================

st.set_page_config(
    page_title="Basket Analysis BI",
    layout="wide"
)

st.title("Basket Analysis")


# =========================================================
# SIDEBAR
# =========================================================

st.sidebar.header(
    "Coloring Rules"
)

aff_threshold = st.sidebar.slider(
    "Affinity threshold",
    0.0,
    5.0,
    1.5,
    0.1
)

high_trips = st.sidebar.slider(
    "High Trips threshold",
    0,
    100,
    30,
    1
)

low_trips = st.sidebar.slider(
    "Low Trips threshold",
    0,
    100,
    15,
    1
)


# =========================================================
# FILE UPLOAD
# =========================================================

st.header(
    "1. Upload files"
)

projects_zip = st.file_uploader(
    "Upload projects.zip",
    type=["zip"]
)

reference_file = st.file_uploader(
    "Upload Список.xlsx",
    type=["xlsx"]
)

total_file = st.file_uploader(
    "Upload total.xlsx",
    type=["xlsx"]
)


# =========================================================
# GENERATE SHEET CONFIG
# =========================================================

def generate_sheet_config(
    zip_path
):

    names = set()

    with tempfile.TemporaryDirectory() as tmp:

        with zipfile.ZipFile(
            zip_path,
            "r"
        ) as z:

            z.extractall(
                tmp
            )

        for root, dirs, files in os.walk(tmp):

            for file in files:

                if file.endswith(".xlsx"):

                    name = Path(
                        file
                    ).stem.split("_")[-1]

                    names.add(
                        name
                    )

    names = sorted(
        names
    )

    return pd.DataFrame(
        {
            "original": names,
            "rename": names,
            "order": range(
                1,
                len(names) + 1
            )
        }
    )


# =========================================================
# SESSION STATE
# =========================================================

if "sheet_config_df" not in st.session_state:

    st.session_state.sheet_config_df = None

if "total_mapping_df" not in st.session_state:

    st.session_state.total_mapping_df = None

if "project_mapping_df" not in st.session_state:

    st.session_state.project_mapping_df = None


# =========================================================
# LOAD FILES
# =========================================================

if projects_zip:

    projects_bytes = projects_zip.getvalue()

    if (
        st.session_state.sheet_config_df
        is None
    ):

        st.session_state.sheet_config_df = (
            generate_sheet_config(
                projects_bytes
            )
        )


if total_file:

    total_bytes = total_file.getvalue()

    if (
        st.session_state.total_mapping_df
        is None
    ):

        with tempfile.NamedTemporaryFile(
            suffix=".xlsx",
            delete=False
        ) as tmp:

            tmp.write(
                total_bytes
            )

            temp_path = tmp.name

        total_df_temp = pd.read_excel(
            temp_path
        )

        os.unlink(
            temp_path
        )

        unique_categories = sorted(
            total_df_temp.iloc[:, 1]
            .dropna()
            .astype(str)
            .unique()
        )

        st.session_state.total_mapping_df = pd.DataFrame(
            {
                "original": unique_categories,
                "rename": unique_categories
            }
        )

        unique_projects = sorted(
            total_df_temp.iloc[:, 0]
            .dropna()
            .astype(str)
            .unique()
        )

        st.session_state.total_projects = (
            unique_projects
        )


# =========================================================
# SHEET CONFIG
# =========================================================

if (
    st.session_state.sheet_config_df
    is not None
):

    st.header(
        "2. Sheet Configuration"
    )

    st.write(
        "Измени названия листов и их порядок."
    )

    st.session_state.sheet_config_df = st.data_editor(
        st.session_state.sheet_config_df,
        use_container_width=True,
        num_rows="fixed",
        key="sheet_config_editor"
    )


# =========================================================
# TOTAL MAPPING
# =========================================================

if (
    st.session_state.total_mapping_df
    is not None
):

    st.header(
        "3. Total Category Mapping"
    )

    st.write(
        "При необходимости сопоставь категории из total.xlsx."
    )

    st.session_state.total_mapping_df = st.data_editor(
        st.session_state.total_mapping_df,
        use_container_width=True,
        num_rows="fixed",
        key="total_mapping_editor"
    )


# =========================================================
# PROJECT MAPPING
# =========================================================

if (
    projects_zip
    and total_file
    and hasattr(
        st.session_state,
        "total_projects"
    )
):

    st.header(
        "4. Project Mapping"
    )

    st.write(
        "Сопоставь папку проекта с проектом из total.xlsx."
    )

    project_folders = []

    with tempfile.TemporaryDirectory() as tmp:

        with zipfile.ZipFile(
            projects_zip.getvalue(),
            "r"
        ) as z:

            z.extractall(
                tmp
            )

        for root, dirs, files in os.walk(tmp):

            for directory in dirs:

                folder = Path(
                    root
                ) / directory

                excel_files = list(
                    folder.glob("*.xlsx")
                )

                if excel_files:

                    project_folders.append(
                        directory
                    )

    project_folders = sorted(
        list(
            set(
                project_folders
            )
        )
    )

    if (
        st.session_state.project_mapping_df
        is None
    ):

        st.session_state.project_mapping_df = pd.DataFrame(
            {
                "project_folder": project_folders,
                "total_project": [
                    ""
                    for _ in project_folders
                ]
            }
        )

    st.session_state.project_mapping_df = st.data_editor(
        st.session_state.project_mapping_df,
        column_config={
            "total_project": st.column_config.SelectboxColumn(
                "Project in total.xlsx",
                options=st.session_state.total_projects
            )
        },
        use_container_width=True,
        num_rows="fixed",
        key="project_mapping_editor"
    )


# =========================================================
# RUN PIPELINE
# =========================================================

st.header(
    "5. Run analysis"
)

if st.button(
    "RUN PIPELINE",
    type="primary"
):

    if not projects_zip:

        st.error(
            "Upload projects.zip"
        )

        st.stop()

    if not reference_file:

        st.error(
            "Upload Список.xlsx"
        )

        st.stop()

    if not total_file:

        st.error(
            "Upload total.xlsx"
        )

        st.stop()

    # =====================================================
    # SAVE FILES
    # =====================================================

    with open(
        "projects.zip",
        "wb"
    ) as f:

        f.write(
            projects_zip.getbuffer()
        )

    with open(
        "Список.xlsx",
        "wb"
    ) as f:

        f.write(
            reference_file.getbuffer()
        )

    with open(
        "total.xlsx",
        "wb"
    ) as f:

        f.write(
            total_file.getbuffer()
        )

    st.session_state.sheet_config_df.to_excel(
        "sheet_config.xlsx",
        index=False
    )

    st.session_state.total_mapping_df.to_excel(
        "total_mapping.xlsx",
        index=False
    )

    st.session_state.project_mapping_df.to_excel(
        "project_mapping.xlsx",
        index=False
    )

    # =====================================================
    # RUN
    # =====================================================

    with st.spinner(
        "Processing..."
    ):

        output_zip = run_pipeline(
            "projects.zip",
            "Список.xlsx",
            "sheet_config.xlsx",
            "total.xlsx",
            "total_mapping.xlsx",
            "project_mapping.xlsx"
        )

    st.success(
        "Analysis completed!"
    )

    with open(
        output_zip,
        "rb"
    ) as f:

        st.download_button(
            "Download results.zip",
            f,
            file_name="results.zip",
            mime="application/zip"
        )


# =========================================================
# INTERACTIVE ANALYSIS
# =========================================================

st.divider()

st.header(
    "Interactive Analysis"
)


RESULTS_DIR = "results"


if not os.path.exists(
    RESULTS_DIR
):

    st.info(
        "Run pipeline first."
    )

    st.stop()


files = [
    f
    for f in os.listdir(
        RESULTS_DIR
    )
    if f.endswith(".xlsx")
]


if not files:

    st.info(
        "No results found."
    )

    st.stop()


selected_file = st.selectbox(
    "Select project file",
    files
)


file_path = os.path.join(
    RESULTS_DIR,
    selected_file
)


xls = pd.ExcelFile(
    file_path
)


selected_sheet = st.selectbox(
    "Select sheet",
    xls.sheet_names
)


df = pd.read_excel(
    xls,
    sheet_name=selected_sheet,
    header=1
)


df.columns = [
    str(c)
    for c in df.columns
]


# =========================================================
# CLASSIFICATION
# =========================================================

def safe_float(value):

    try:

        return float(
            value
        )

    except:

        return None


def classify_row(row):

    affinity = safe_float(
        row.get(
            "Affinity index to FMCG"
        )
    )

    trips = safe_float(
        row.get(
            "target_trips_raw_CS"
        )
    )

    if (
        affinity is not None
        and trips is not None
        and affinity >= aff_threshold
        and trips > high_trips
    ):

        return "orange"

    if (
        trips is not None
        and trips < low_trips
    ):

        return "light_gray"

    if (
        affinity is not None
        and affinity >= aff_threshold
    ):

        return "gray"

    return "none"


df["color"] = df.apply(
    classify_row,
    axis=1
)


# =========================================================
# KPI
# =========================================================

st.subheader(
    "KPI Overview"
)


col1, col2, col3 = st.columns(3)


trips_mean = None

if "Trips (000)" in df.columns:

    trips_mean = pd.to_numeric(
        df["Trips (000)"],
        errors="coerce"
    ).mean()


orange_df = df[
    df["color"] == "orange"
]


with col1:

    st.metric(
        "Avg Trips (000)",
        (
            f"{trips_mean:.2f}"
            if trips_mean is not None
            else "N/A"
        )
    )


with col2:

    st.metric(
        "Orange rows",
        len(orange_df)
    )


with col3:

    orange_pct = (
        len(orange_df)
        / len(df)
        * 100
        if len(df)
        else 0
    )

    st.metric(
        "% Orange Rows",
        f"{orange_pct:.1f}%"
    )


# =========================================================
# PREVIEW
# =========================================================

st.subheader(
    "Data Preview"
)

st.dataframe(
    df,
    use_container_width=True
)


# =========================================================
# ORANGE CATEGORIES
# =========================================================

st.subheader(
    "Orange Categories"
)


category_col = df.columns[0]


orange_categories = sorted(
    orange_df[category_col]
    .dropna()
    .astype(str)
    .unique()
)


selected_categories = st.multiselect(
    "Select categories for scatter plot",
    orange_categories,
    default=orange_categories
)


# =========================================================
# SCATTER
# =========================================================

st.subheader(
    "Scatter Plot"
)


numeric_cols = []

for column in df.columns:

    converted = pd.to_numeric(
        df[column],
        errors="coerce"
    )

    if converted.notna().sum() > 0:

        numeric_cols.append(
            column
        )


if len(numeric_cols) < 2:

    st.warning(
        "Not enough numeric columns."
    )

    st.stop()


col1, col2 = st.columns(2)


with col1:

    x_axis = st.selectbox(
        "X axis",
        numeric_cols,
        key="x_axis"
    )


with col2:

    y_axis = st.selectbox(
        "Y axis",
        numeric_cols,
        key="y_axis"
    )


plot_df = orange_df[
    orange_df[category_col]
    .astype(str)
    .isin(
        selected_categories
    )
].copy()


plot_df[x_axis] = pd.to_numeric(
    plot_df[x_axis],
    errors="coerce"
)

plot_df[y_axis] = pd.to_numeric(
    plot_df[y_axis],
    errors="coerce"
)


plot_df = plot_df.dropna(
    subset=[
        x_axis,
        y_axis
    ]
)


if len(plot_df) > 0:

    x_min = float(
        plot_df[x_axis].min()
    )

    x_max = float(
        plot_df[x_axis].max()
    )

    y_min = float(
        plot_df[y_axis].min()
    )

    y_max = float(
        plot_df[y_axis].max()
    )

    x_range = st.slider(
        "X axis range",
        min_value=x_min,
        max_value=x_max,
        value=(x_min, x_max),
        key=f"x_range_{selected_sheet}_{x_axis}"
    )

    y_range = st.slider(
        "Y axis range",
        min_value=y_min,
        max_value=y_max,
        value=(y_min, y_max),
        key=f"y_range_{selected_sheet}_{y_axis}"
    )

    fig = px.scatter(
        plot_df,
        x=x_axis,
        y=y_axis,
        text=category_col
    )

    fig.update_traces(
        textposition="top center",
        marker={
            "size": 10
        },
        showlegend=False
    )

    fig.update_layout(
        showlegend=False,
        xaxis_range=x_range,
        yaxis_range=y_range
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

else:

    st.info(
        "Select at least one category."
    )
