# -*- coding: utf-8 -*-

import pandas as pd
import zipfile
import shutil
import os
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill


# =========================
# PROCESS SHEET
# =========================

def process_sheet(df, valid_values):

    df = df.iloc[:, [1,2,7,8,15,16,17]].copy()
    df.columns = list("ABCDEFG")

    mask = (df["A"] == "        ALL VALUES") & (df["B"] != 0)
    df.loc[mask, "A"] = df.loc[mask, "B"]

    df = df.drop(columns=["B"])
    df.columns = list("ABCDEF")

    df = df[["A","D","E","F","B","C"]]

    df.insert(3, "Trips (000)", "")

    df["filter"] = df["A"].astype(str).apply(
        lambda x: 0 if x in valid_values else 1
    )

    df = df[df["A"].astype(str).str.lower() != "    rest"]

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

    # ❗ ВАЖНО: убираем строки сразу тут (а не в openpyxl)
    df = df[df["filter"] == 1].copy()

    return df


# =========================
# MATCH PROJECT
# =========================

def match_project(name, project_name):

    name = str(name).lower()
    parts = name.strip().split()

    if len(parts) == 0:
        return False

    if parts[0] == "total" and len(parts) > 1:
        return parts[1] == project_name

    return parts[0] == project_name


# =========================
# FORMAT EXCEL
# =========================

def apply_formatting(output_file):

    wb = load_workbook(output_file)

    gray = PatternFill("solid", fgColor="777777")
    light_gray = PatternFill("solid", fgColor="CCCCCC")
    orange = PatternFill("solid", fgColor="FFA500")

    for ws in wb.worksheets:

        # header
        ws.insert_rows(1)
        ws.merge_cells("B1:E1")

        cell = ws["B1"]
        cell.value = ws.title
        cell.font = Font(bold=True, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in range(3, ws.max_row + 1):

            try:
                f_val = ws.cell(row=row, column=6).value
                h_val = ws.cell(row=row, column=8).value

                if h_val != 1:
                    continue

                f_val = float(f_val)

                e_val = ws.cell(row=row, column=5).value
                c_val = ws.cell(row=row, column=3).value

                if f_val > 30 and e_val >= 1.5:

                    for col in [2,3,4,5]:
                        ws.cell(row=row, column=col).fill = orange

                    if c_val > 0.05:
                        ws.cell(row=row, column=1).font = Font(bold=True)

                elif f_val < 15:

                    for col in [2,3,4,5]:
                        ws.cell(row=row, column=col).fill = gray

                    ws.cell(row=row, column=5).value = None

                else:

                    for col in [2,3,4,5]:
                        ws.cell(row=row, column=col).fill = light_gray

            except:
                continue

        # удалить лишние колонки
        if ws.max_column > 5:
            ws.delete_cols(6, ws.max_column - 5)

    wb.save(output_file)


# =========================
# MAIN PIPELINE
# =========================

def run_pipeline(projects_zip_path, reference_path, config_path, total_path):

    shutil.rmtree("projects", ignore_errors=True)
    shutil.rmtree("results", ignore_errors=True)

    Path("projects").mkdir(exist_ok=True)
    Path("results").mkdir(exist_ok=True)

    # unzip
    with zipfile.ZipFile(projects_zip_path, 'r') as z:
        z.extractall("projects")

    # config
    config = pd.read_excel(config_path)
    rename_map = dict(zip(config["original"], config["rename"]))
    order_map = dict(zip(config["rename"], config["order"]))

    # reference
    ref_df = pd.read_excel(reference_path, sheet_name="Список")
    valid_values = set(ref_df.iloc[:,0].dropna().astype(str))

    # total
    total_df = pd.read_excel(total_path)

    base_path = Path("projects")
    result_path = Path("results")

    # собрать реальные project folders (устойчиво к вложенности zip)
    project_folders = []

    for path in base_path.rglob("*"):
        if path.is_dir() and list(path.glob("*.xlsx")):
            project_folders.append(path)

    for project_folder in project_folders:

        sheets = []
        project_name = project_folder.name.lower()

        for file in project_folder.glob("*.xlsx"):

            df = pd.read_excel(file)

            original = file.stem.split("_")[-1]
            sheet_name = rename_map.get(original, original)
            order = order_map.get(sheet_name, 999)

            processed_df = process_sheet(df, valid_values)

            # --- TOTAL MATCH ---
            row = total_df[
                total_df.iloc[:,0].apply(
                    lambda x: match_project(x, project_name)
                )
                &
                (
                    total_df.iloc[:,1]
                    .astype(str)
                    .str.strip()
                    .str.lower()
                    == sheet_name.lower()
                )
            ]

            if not row.empty:
                trips_rp = row.iloc[0,3]
            else:
                trips_rp = None

            # --- Trips (000) ---
            if trips_rp is not None and "Trips Share" in processed_df.columns:

                trips_list = [trips_rp]

                for val in processed_df["Trips Share"].iloc[1:]:

                    try:
                        trips_list.append(trips_rp * float(val))
                    except:
                        trips_list.append(None)

                processed_df["Trips (000)"] = trips_list

            sheets.append((sheet_name, order, processed_df))

        # если пусто — пропуск (ВАЖНО чтобы не падал ExcelWriter)
        if not sheets:
            continue

        sheets = sorted(sheets, key=lambda x: x[1])

        output_file = result_path / f"Basket Analysis_{project_folder.name}.xlsx"

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

            for sheet_name, _, df in sheets:
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        apply_formatting(output_file)

    # zip results
    output_zip = "results.zip"

    with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as z:

        for file in result_path.glob("*.xlsx"):
            z.write(file, arcname=file.name)

    return output_zip
