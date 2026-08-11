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

aff_threshold_percent = st.sidebar.slider(
    "Affinity threshold (%)",
    0,
    500,
    150,
    10
)

aff_threshold = aff_threshold_percent / 100

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
# SESSION STATE INIT
# =========================

if "sheet_config" not in st.session_state:
    st.session_state.sheet_config = None

if "mapping" not in st.session_state:
    st.session_state.mapping = None


# =========================
# FILE UPLOADS
# =========================

st.info(
    "Папка projects.zip - архив с исходными папками проектов, где каждая папка имеет название категории, а внутри хранятся листы по каналам"
)

projects_zip = st.file_uploader(
    "Upload projects.zip",
    type=["zip"]
)

st.info(
    "Файл Список.xlsx содержит список необходимых food категорий"
)

reference_file = st.file_uploader(
    "Upload Список.xlsx",
    type=["xlsx"]
)

st.info(
    "Файл total.xlsx содержит ось каналов по категориям с Trips raw и Trips(000)"
)

total_file = st.file_uploader(
    "Upload total.xlsx",
    type=["xlsx"]
)


# =========================
# GENERATE SHEET CONFIG
# =========================

def generate_sheet_config(zip_path):

    with tempfile.TemporaryDirectory() as tmp:

        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(tmp)

        names = set()

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


# =========================
# GENERATE MAPPING
# =========================

def generate_mapping(total_file):

    total_df = pd.read_excel(total_file)

    mapping_df = pd.DataFrame({
        "original": sorted(
            total_df.iloc[:, 1]
            .dropna()
            .astype(str)
            .unique()
        )
    })

    mapping_df["rename"] = mapping_df["original"]

    return mapping_df

def generate_mapping1(total_file):

    total_df = pd.read_excel(total_file)

    mapping_df = pd.DataFrame({
        "original": sorted(
            total_df.iloc[:, 0]
            .dropna()
            .astype(str)
            .unique()
        )
    })

    mapping_df["rename"] = mapping_df["original"]

    return mapping_df


# =========================
# BUILD CONFIG UI
# =========================

if projects_zip:

    st.subheader("Sheets configuration")

    st.session_state.sheet_config = generate_sheet_config(
        projects_zip
    )

    st.session_state.sheet_config = st.data_editor(
        st.session_state.sheet_config,
        use_container_width=True
    )


if total_file:

    st.subheader("Total File")
    st.info(
    "Проверьте, что названия в правых столбцах в файле Total и Sheets configuration совпадают"
)

    st.session_state.mapping = generate_mapping(
        total_file
    )

    st.session_state.mapping = generate_mapping1(
        total_file
    )

    st.session_state.mapping = st.data_editor(
        st.session_state.mapping,
        use_container_width=True
    )


# =========================
# RUN PIPELINE
# =========================

if st.button("RUN"):

    if not projects_zip or not reference_file or not total_file:

        st.error("Upload all required files")

        st.stop()


    # save files

    with open("projects.zip", "wb") as f:

        f.write(
            projects_zip.getbuffer()
        )


    with open("Список.xlsx", "wb") as f:

        f.write(
            reference_file.getbuffer()
        )


    with open("total.xlsx", "wb") as f:

        f.write(
            total_file.getbuffer()
        )


    # save configs from UI

    if st.session_state.sheet_config is None:

        st.error("Sheet config missing")

        st.stop()


    if st.session_state.mapping is None:

        st.error("Mapping missing")

        st.stop()


    st.session_state.sheet_config.to_excel(
        "sheet_config.xlsx",
        index=False
    )


    st.session_state.mapping.to_excel(
        "total_mapping.xlsx",
        index=False
    )


    with st.spinner("Processing..."):

        output_zip = run_pipeline(
            "projects.zip",
            "Список.xlsx",
            "sheet_config.xlsx",
            "total.xlsx",
            "total_mapping.xlsx"
        )


    st.success("Done!")


    with open(output_zip, "rb") as f:

        st.download_button(
            "Download results",
            f,
            file_name="results.zip"
        )


# =========================
# RESULTS VIEW
# =========================

st.divider()

st.header("📊 Analysis")


RESULTS_DIR = "results"


if os.path.exists(RESULTS_DIR):

    files = [
        f
        for f in os.listdir(RESULTS_DIR)
        if f.endswith(".xlsx")
    ]


    if files:

        selected_file = st.selectbox(
            "Select file",
            files
        )


        xls = pd.ExcelFile(
            os.path.join(
                RESULTS_DIR,
                selected_file
            )
        )


        sheet = st.selectbox(
            "Select sheet",
            xls.sheet_names
        )


        df = pd.read_excel(
            xls,
            sheet_name=sheet,
            header=1
        )


        df.columns = [
            str(c)
            for c in df.columns
        ]


        st.subheader("Preview")


        # ============================================================
        # DISPLAY COPY
        # ============================================================

        display_df = df.copy()


        percentage_columns = [
            "Trips Share",
            "FMCG trips share",
            "Affinity index to FMCG"
        ]


        for col in percentage_columns:

            if col in display_df.columns:

                display_df[col] = pd.to_numeric(
                    display_df[col],
                    errors="coerce"
                ).map(
                    lambda x: f"{x * 100:.2f}%"
                    if pd.notna(x)
                    else ""
                )


        st.dataframe(
            display_df,
            use_container_width=True
        )


        # =========================
        # COLOR LOGIC
        # =========================

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


        # =========================
        # FILTER
        # =========================

        category_col = df.columns[0]


        st.subheader("🟠 Orange filter")


        categories = sorted(
            orange_df[
                category_col
            ]
            .dropna()
            .astype(str)
            .unique()
        )


        selected = st.multiselect(
            "Select categories",
            categories,
            default=categories
        )


        plot_df = orange_df[
            orange_df[
                category_col
            ]
            .astype(str)
            .isin(selected)
        ]


        # =========================
        # SCATTER
        # =========================

        numeric_cols = df.select_dtypes(
            include=["number"]
        ).columns.tolist()


        if len(numeric_cols) >= 2:

            col1, col2 = st.columns(2)


            with col1:

                x = st.selectbox(
                    "X axis",
                    numeric_cols
                )


            with col2:

                y = st.selectbox(
                    "Y axis",
                    numeric_cols
                )


            # ========================================================
            # DISPLAY COPY FOR PLOT
            # ========================================================

            plot_df_display = plot_df.copy()


            for col in percentage_columns:

                if col in plot_df_display.columns:

                    plot_df_display[col] = pd.to_numeric(
                        plot_df_display[col],
                        errors="coerce"
                    ) * 100


            fig = px.scatter(
                plot_df_display,
                x=x,
                y=y,
                color="color",
                text=category_col
            )


            fig.update_traces(
                textposition="top center"
            )


            fig.update_layout(
                showlegend=False
            )


            # ========================================================
            # AXIS FORMAT AS PERCENT
            # ========================================================

            if x in percentage_columns:

                fig.update_xaxes(
                    ticksuffix="%"
                )


            if y in percentage_columns:

                fig.update_yaxes(
                    ticksuffix="%"
                )


            st.plotly_chart(
                fig,
                use_container_width=True
            )
