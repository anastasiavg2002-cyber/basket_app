# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import os
import zipfile
import tempfile
import plotly.express as px

from pathlib import Path

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
# SIDEBAR
# =========================

st.sidebar.header("Coloring Rules")

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


# =========================
# SESSION STATE
# =========================

if "sheet_config" not in st.session_state:
    st.session_state.sheet_config = None

if "project_mapping" not in st.session_state:
    st.session_state.project_mapping = None

if "projects_signature" not in st.session_state:
    st.session_state.projects_signature = None

if "total_signature" not in st.session_state:
    st.session_state.total_signature = None


# =========================
# FILE UPLOADS
# =========================

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


# ============================================================
# HELPERS
# ============================================================

def get_projects_from_zip(uploaded_zip):

    projects = set()

    with tempfile.TemporaryDirectory() as tmp:

        zip_bytes = uploaded_zip.getvalue()

        zip_path = os.path.join(tmp, "projects.zip")

        with open(zip_path, "wb") as f:
            f.write(zip_bytes)

        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(tmp)

        for root, dirs, files in os.walk(tmp):

            if any(file.endswith(".xlsx") for file in files):

                folder_name = os.path.basename(root)

                if folder_name != os.path.basename(tmp):

                    projects.add(folder_name)

    return sorted(projects)


def generate_sheet_config(uploaded_zip):

    names = set()

    with tempfile.TemporaryDirectory() as tmp:

        zip_bytes = uploaded_zip.getvalue()

        zip_path = os.path.join(tmp, "projects.zip")

        with open(zip_path, "wb") as f:
            f.write(zip_bytes)

        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(tmp)

        for root, dirs, files in os.walk(tmp):

            for file in files:

                if file.endswith(".xlsx"):

                    name = Path(file).stem.split("_")[-1]

                    names.add(name)

    names = sorted(names)

    return pd.DataFrame({
        "original": names,
        "rename": names,
        "order": range(1, len(names) + 1)
    })


def generate_project_mapping(uploaded_zip, total_file):

    projects = get_projects_from_zip(uploaded_zip)

    total_df = pd.read_excel(total_file)

    total_projects = sorted(
        total_df.iloc[:, 0]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
    )

    rows = []

    for project in projects:

        rows.append({
            "project_folder": project,
            "total_project": ""
        })

    mapping_df = pd.DataFrame(rows)

    # Автоматическое сопоставление,
    # если названия полностью совпадают
    for i in range(len(mapping_df)):

        project_name = str(
            mapping_df.loc[i, "project_folder"]
        ).strip().lower()

        matches = [
            x for x in total_projects
            if str(x).strip().lower() == project_name
        ]

        if len(matches) == 1:

            mapping_df.loc[i, "total_project"] = matches[0]

    return mapping_df, total_projects


# ============================================================
# SHEET CONFIGURATION
# ============================================================

if projects_zip:

    projects_signature = projects_zip.file_id

    if (
        st.session_state.sheet_config is None
        or st.session_state.projects_signature != projects_signature
    ):

        st.session_state.sheet_config = generate_sheet_config(
            projects_zip
        )

        st.session_state.projects_signature = projects_signature

    st.subheader("📄 Sheets configuration")

    st.info(
        "Можно изменить названия листов и их порядок."
    )

    st.session_state.sheet_config = st.data_editor(
        st.session_state.sheet_config,
        use_container_width=True,
        num_rows="dynamic",
        key="sheet_config_editor"
    )


# ============================================================
# PROJECT ↔ TOTAL MAPPING
# ============================================================

if projects_zip and total_file:

    total_signature = total_file.file_id

    if (
        st.session_state.project_mapping is None
        or st.session_state.total_signature != total_signature
        or st.session_state.projects_signature != projects_zip.file_id
    ):

        (
            st.session_state.project_mapping,
            total_projects
        ) = generate_project_mapping(
            projects_zip,
            total_file
        )

        st.session_state.total_signature = total_signature

    st.subheader("🔗 Project ↔ Total mapping")

    st.info(
        "Выберите, какое значение из total соответствует каждой папке проекта."
    )

    mapping_options = [""] + sorted(
        pd.read_excel(total_file)
        .iloc[:, 0]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
        .tolist()
    )

    st.session_state.project_mapping = st.data_editor(
        st.session_state.project_mapping,
        use_container_width=True,
        hide_index=True,
        column_config={
            "project_folder": st.column_config.TextColumn(
                "Project folder",
                disabled=True
            ),

            "total_project": st.column_config.SelectboxColumn(
                "Total project",
                options=mapping_options,
                required=False
            )
        },
        key="project_mapping_editor"
    )


# ============================================================
# RUN PIPELINE
# ============================================================

if st.button("🚀 RUN PIPELINE"):

    if not projects_zip:

        st.error("Upload projects.zip")
        st.stop()

    if not reference_file:

        st.error("Upload Список.xlsx")
        st.stop()

    if not total_file:

        st.error("Upload total.xlsx")
        st.stop()

    if st.session_state.sheet_config is None:

        st.error("Sheet configuration is missing")
        st.stop()

    if st.session_state.project_mapping is None:

        st.error("Project mapping is missing")
        st.stop()

    # -------------------------
    # SAVE UPLOADS
    # -------------------------

    with open("projects.zip", "wb") as f:
        f.write(projects_zip.getvalue())

    with open("Список.xlsx", "wb") as f:
        f.write(reference_file.getvalue())

    with open("total.xlsx", "wb") as f:
        f.write(total_file.getvalue())

    # -------------------------
    # SAVE CONFIG
    # -------------------------

    st.session_state.sheet_config.to_excel(
        "sheet_config.xlsx",
        index=False
    )

    st.session_state.project_mapping.to_excel(
        "project_mapping.xlsx",
        index=False
    )

    # -------------------------
    # RUN
    # -------------------------

    with st.spinner("Processing pipeline..."):

        output_zip = run_pipeline(
            "projects.zip",
            "Список.xlsx",
            "sheet_config.xlsx",
            "total.xlsx",
            "project_mapping.xlsx"
        )

    st.success("✅ Done!")

    with open(output_zip, "rb") as f:

        st.download_button(
            "⬇️ Download results",
            f,
            file_name="results.zip"
        )


# ============================================================
# RESULTS
# ============================================================

st.divider()

st.header("📊 Analysis")


RESULTS_DIR = "results"


if not os.path.exists(RESULTS_DIR):

    st.info("Run pipeline first")

    st.stop()


files = [
    f
    for f in os.listdir(RESULTS_DIR)
    if f.endswith(".xlsx")
]


if not files:

    st.info("No results found")

    st.stop()


selected_file = st.selectbox(
    "Select project file",
    files
)


file_path = os.path.join(
    RESULTS_DIR,
    selected_file
)


xls = pd.ExcelFile(file_path)


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


# ============================================================
# CLASSIFICATION
# ============================================================

def classify(row):

    try:

        aff = float(
            row.get(
                "Affinity index to FMCG",
                0
            )
        )

        trips = float(
            row.get(
                "target_trips_raw_CS",
                0
            )
        )

    except:

        return "none"


    if (
        aff >= aff_threshold
        and trips > high_trips
    ):

        return "orange"


    if trips < low_trips:

        return "light_gray"


    if aff >= aff_threshold:

        return "gray"


    return "none"


df["color"] = df.apply(
    classify,
    axis=1
)


orange_df = df[
    df["color"] == "orange"
]


# ============================================================
# DISPLAY COPY
# ============================================================

display_df = df.copy()


# ПЕРЕВОДИМ ТОЛЬКО ОТОБРАЖЕНИЕ В ПРОЦЕНТЫ
for col in [
    "Trips Share",
    "Affinity index to FMCG"
]:

    if col in display_df.columns:

        display_df[col] = pd.to_numeric(
            display_df[col],
            errors="coerce"
        ).map(
            lambda x: f"{x * 100:.2f}%"
            if pd.notna(x)
            else ""
        )


# ============================================================
# PREVIEW
# ============================================================

st.subheader("📄 Preview")

st.dataframe(
    display_df,
    use_container_width=True
)


# ============================================================
# ORANGE CATEGORIES
# ============================================================

st.subheader("🟠 Orange categories")


category_col = df.columns[0]


categories = sorted(
    orange_df[category_col]
    .dropna()
    .astype(str)
    .unique()
)


selected_categories = st.multiselect(
    "Select categories for scatter plot",
    categories,
    default=categories
)


plot_df = orange_df[
    orange_df[category_col]
    .astype(str)
    .isin(selected_categories)
].copy()


# ============================================================
# SCATTER PLOT
# ============================================================

st.subheader("📈 Scatter plot")


numeric_cols = df.select_dtypes(
    include=["number"]
).columns.tolist()


if len(numeric_cols) >= 2:

    col1, col2 = st.columns(2)


    with col1:

        x_axis = st.selectbox(
            "X axis",
            numeric_cols
        )


    with col2:

        y_axis = st.selectbox(
            "Y axis",
            numeric_cols
        )


    # -------------------------
    # RANGE
    # -------------------------

    x_min = float(
        pd.to_numeric(
            plot_df[x_axis],
            errors="coerce"
        ).min()
    )

    x_max = float(
        pd.to_numeric(
            plot_df[x_axis],
            errors="coerce"
        ).max()
    )


    y_min = float(
        pd.to_numeric(
            plot_df[y_axis],
            errors="coerce"
        ).min()
    )

    y_max = float(
        pd.to_numeric(
            plot_df[y_axis],
            errors="coerce"
        ).max()
    )


    if x_min == x_max:

        x_max = x_min + 1


    if y_min == y_max:

        y_max = y_min + 1


    range_col1, range_col2 = st.columns(2)


    with range_col1:

        x_range = st.slider(
            "X axis range",
            min_value=x_min,
            max_value=x_max,
            value=(x_min, x_max),
            key="x_axis_range"
        )


    with range_col2:

        y_range = st.slider(
            "Y axis range",
            min_value=y_min,
            max_value=y_max,
            value=(y_min, y_max),
            key="y_axis_range"
        )


    # -------------------------
    # DISPLAY COPY FOR GRAPH
    # -------------------------

    plot_display_df = plot_df.copy()


    for col in [
        "Trips Share",
        "Affinity index to FMCG"
    ]:

        if col in plot_display_df.columns:

            plot_display_df[col] = (
                pd.to_numeric(
                    plot_display_df[col],
                    errors="coerce"
                ) * 100
            )


    fig = px.scatter(
        plot_display_df,
        x=x_axis,
        y=y_axis,
        text=category_col
    )


    fig.update_traces(
        textposition="top center"
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
