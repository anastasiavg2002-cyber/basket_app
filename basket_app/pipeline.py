# -*- coding: utf-8 -*-

import pandas as pd
import zipfile
import shutil

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill


# =========================================================
# PROCESS SHEET
# =========================================================

def process_sheet(df, valid_values):

    df = df.iloc[:, [1, 2, 7, 8, 15, 16, 17]].copy()

    df.columns = list("ABCDEFG")

    mask = (
        (df["A"] == "        ALL VALUES")
        & (df["B"] != 0)
    )

    df.loc[mask, "A"] = df.loc[mask, "B"]

    df = df.drop(columns=["B"])

    df.columns = list("ABCDEF")

    df = df[["A", "D", "E", "F", "B", "C"]]

    df.insert(3, "Trips (000)", "")

    df["filter"] = df["A"].astype(str).apply(
        lambda x: 0 if x in valid_values else 1
    )

    df = df[
        df["A"].astype(str).str.lower() != "    rest"
    ]

    new_columns = [
        "",
        "FMCG trips share",
        "Trips Share",
        "Trips (000)",
        "Affinity index to FMCG",
        "target_trips_raw_CS",
        "target_trips_000_CS",
        "filter"
    ]

    df.columns = new_columns

    return df


# =========================================================
# NORMALIZE TEXT
# =========================================================

def normalize_text(value):

    if pd.isna(value):
        return ""

    return (
        str(value)
        .strip()
        .lower()
        .replace("_", " ")
        .replace("-", " ")
    )


# =========================================================
# FORMAT EXCEL
# =========================================================

def apply_formatting(output_file):

    wb = load_workbook(output_file)

    fill_gray = PatternFill(
        start_color="777777",
        end_color="777777",
        fill_type="solid"
    )

    fill_light_gray = PatternFill(
        start_color="CCCCCC",
        end_color="CCCCCC",
        fill_type="solid"
    )

    orange_fill = PatternFill(
        start_color="FFA500",
        end_color="FFA500",
        fill_type="solid"
    )

    for sheet_name in wb.sheetnames:

        ws = wb[sheet_name]

        # =================================================
        # HEADER
        # =================================================

        ws.insert_rows(1)

        ws.merge_cells("B1:E1")

        cell = ws["B1"]

        cell.value = sheet_name

        cell.font = Font(
            bold=True,
            size=11
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        # =================================================
        # FIND COLUMNS BY NAMES
        # =================================================

        headers = {}

        for col in range(1, ws.max_column + 1):

            value = ws.cell(
                row=2,
                column=col
            ).value

            if value is not None:

                headers[str(value)] = col

        affinity_col = headers.get(
            "Affinity index to FMCG"
        )

        target_trips_col = headers.get(
            "target_trips_raw_CS"
        )

        filter_col = headers.get(
            "filter"
        )

        # =================================================
        # COLORING
        # =================================================

        if (
            affinity_col
            and target_trips_col
            and filter_col
        ):

            for row in range(3, ws.max_row + 1):

                filter_value = ws.cell(
                    row=row,
                    column=filter_col
                ).value

                if filter_value != 1:
                    continue

                affinity_value = ws.cell(
                    row=row,
                    column=affinity_col
                ).value

                trips_value = ws.cell(
                    row=row,
                    column=target_trips_col
                ).value

                try:

                    affinity = float(
                        affinity_value
                    )

                    trips = float(
                        trips_value
                    )

                except:

                    continue

                # =================================================
                # ORANGE
                # Affinity >= 1.5
                # Trips > 30
                # =================================================

                if (
                    affinity >= 1.5
                    and trips > 30
                ):

                    for col in range(
                        2,
                        6
                    ):

                        ws.cell(
                            row=row,
                            column=col
                        ).fill = orange_fill

                # =================================================
                # DARK GRAY
                # Trips < 15
                # =================================================

                elif trips < 15:

                    for col in range(
                        2,
                        6
                    ):

                        ws.cell(
                            row=row,
                            column=col
                        ).fill = fill_gray

                    ws.cell(
                        row=row,
                        column=affinity_col
                    ).value = None

                # =================================================
                # LIGHT GRAY
                # Affinity >= 1.5
                # =================================================

                elif affinity >= 1.5:

                    for col in range(
                        2,
                        6
                    ):

                        ws.cell(
                            row=row,
                            column=col
                        ).fill = fill_light_gray

        # =================================================
        # DELETE ROWS WHERE FILTER = 0
        # =================================================

        if filter_col:

            for row in range(
                ws.max_row,
                2,
                -1
            ):

                value = ws.cell(
                    row=row,
                    column=filter_col
                ).value

                if value == 0:

                    ws.delete_rows(row)

        # =================================================
        # DELETE EXTRA COLUMNS
        # =================================================

        if ws.max_column > 5:

            ws.delete_cols(
                6,
                ws.max_column - 5
            )

        # =================================================
        # FORMAT AFFINITY AS PERCENTAGE
        # =================================================

        for row in range(
            3,
            ws.max_row + 1
        ):

            ws.cell(
                row=row,
                column=3
            ).number_format = "0.00%"

    wb.save(output_file)


# =========================================================
# MAIN PIPELINE
# =========================================================

def run_pipeline(
    projects_zip_path,
    reference_path,
    config_path,
    total_path,
    total_mapping_path,
    project_mapping_path
):

    # =========================================================
    # CLEAN
    # =========================================================

    shutil.rmtree(
        "projects",
        ignore_errors=True
    )

    shutil.rmtree(
        "results",
        ignore_errors=True
    )

    Path("projects").mkdir(
        exist_ok=True
    )

    Path("results").mkdir(
        exist_ok=True
    )

    # =========================================================
    # EXTRACT PROJECTS
    # =========================================================

    with zipfile.ZipFile(
        projects_zip_path,
        "r"
    ) as zip_ref:

        zip_ref.extractall(
            "projects"
        )

    # =========================================================
    # SHEET CONFIG
    # =========================================================

    config = pd.read_excel(
        config_path
    )

    rename_map = dict(
        zip(
            config["original"],
            config["rename"]
        )
    )

    order_map = dict(
        zip(
            config["rename"],
            config["order"]
        )
    )

    # =========================================================
    # REFERENCE FILE
    # =========================================================

    ref_df = pd.read_excel(
        reference_path,
        sheet_name="Список"
    )

    valid_values = set(
        ref_df.iloc[:, 0]
        .dropna()
        .astype(str)
    )

    # =========================================================
    # TOTAL FILE
    # =========================================================

    total_df = pd.read_excel(
        total_path
    )

    # =========================================================
    # TOTAL MAPPING
    # =========================================================

    total_mapping = pd.read_excel(
        total_mapping_path
    )

    total_mapping = total_mapping.dropna(
        subset=["original", "rename"]
    )

    total_mapping_dict = dict(
        zip(
            total_mapping["original"].astype(str),
            total_mapping["rename"].astype(str)
        )
    )

    # =========================================================
    # APPLY TOTAL MAPPING
    # =========================================================

    total_df.iloc[:, 1] = (
        total_df.iloc[:, 1]
        .astype(str)
        .map(
            lambda x: total_mapping_dict.get(
                x,
                x
            )
        )
    )

    # =========================================================
    # PROJECT MAPPING
    # =========================================================

    project_mapping = pd.read_excel(
        project_mapping_path
    )

    project_mapping = project_mapping.dropna(
        subset=[
            "project_folder",
            "total_project"
        ]
    )

    project_mapping_dict = dict(
        zip(
            project_mapping["project_folder"]
            .astype(str)
            .map(normalize_text),

            project_mapping["total_project"]
            .astype(str)
            .map(normalize_text)
        )
    )

    # =========================================================
    # PROJECT FOLDERS
    # =========================================================

    base_path = Path(
        "projects"
    )

    result_path = Path(
        "results"
    )

    project_folders = []

    for path in base_path.rglob("*"):

        if not path.is_dir():

            continue

        excel_files = list(
            path.glob("*.xlsx")
        )

        if excel_files:

            project_folders.append(
                path
            )

    # =========================================================
    # PROCESS PROJECTS
    # =========================================================

    for project_folder in project_folders:

        sheets = []

        folder_name = project_folder.name

        folder_name_normalized = normalize_text(
            folder_name
        )

        # Manual project mapping
        total_project_name = project_mapping_dict.get(
            folder_name_normalized,
            folder_name_normalized
        )

        for file in project_folder.glob("*.xlsx"):

            df = pd.read_excel(
                file
            )

            original = file.stem.split("_")[-1]

            sheet_name = rename_map.get(
                original,
                original
            )

            order = order_map.get(
                sheet_name,
                999
            )

            processed_df = process_sheet(
                df,
                valid_values
            )

            # =================================================
            # FIND TOTAL ROW
            # =================================================

            total_project_col = (
                total_df.iloc[:, 0]
                .astype(str)
                .map(normalize_text)
            )

            total_category_col = (
                total_df.iloc[:, 1]
                .astype(str)
                .map(normalize_text)
            )

            row = total_df[
                (
                    total_project_col
                    == total_project_name
                )
                &
                (
                    total_category_col
                    == normalize_text(
                        sheet_name
                    )
                )
            ]

            # =================================================
            # GET TRIPS
            # =================================================

            if not row.empty:

                trips_raw = row.iloc[0, 2]

                trips_rp = row.iloc[0, 3]

            else:

                trips_raw = None

                trips_rp = None

            processed_df[
                "trips_raw"
            ] = trips_raw

            processed_df[
                "trips_rp"
            ] = trips_rp

            # =================================================
            # CALCULATE TRIPS (000)
            # =================================================

            if (
                trips_rp is not None
                and "Trips Share" in processed_df.columns
            ):

                trips_list = []

                for index, value in enumerate(
                    processed_df["Trips Share"]
                ):

                    try:

                        value = float(
                            value
                        )

                    except:

                        value = 0

                    if index == 0:

                        trips_list.append(
                            trips_rp
                        )

                    else:

                        trips_list.append(
                            trips_rp * value
                        )

                processed_df[
                    "Trips (000)"
                ] = trips_list

            sheets.append(
                (
                    sheet_name,
                    order,
                    processed_df
                )
            )

        # =================================================
        # SORT SHEETS
        # =================================================

        sheets = sorted(
            sheets,
            key=lambda x: x[1]
        )

        # =================================================
        # SAFETY CHECK
        # =================================================

        if not sheets:

            continue

        # =================================================
        # WRITE EXCEL
        # =================================================

        output_file = (
            result_path
            / f"Basket Analysis_{folder_name}.xlsx"
        )

        with pd.ExcelWriter(
            output_file,
            engine="openpyxl"
        ) as writer:

            for sheet_name, _, df in sheets:

                df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False
                )

        # =================================================
        # FORMAT
        # =================================================

        apply_formatting(
            output_file
        )

    # =========================================================
    # ZIP RESULTS
    # =========================================================

    output_zip = "results.zip"

    with zipfile.ZipFile(
        output_zip,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zipf:

        for file in result_path.glob("*.xlsx"):

            zipf.write(
                file,
                arcname=file.name
            )

    return output_zip
