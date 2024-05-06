import datetime
import glob
import json
import os
import re
import time
from pathlib import WindowsPath
from typing import List

from .constants import tipos_datos, second_pass_cast

import duckdb
import openpyxl
import polars as pl
from python_calamine import CalamineWorkbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, NamedStyle, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string
from openpyxl.utils.dataframe import dataframe_to_rows
from pandas import DataFrame
from rich.console import Console

console = Console()


def load_json_config(file):
    """
    Loads a JSON file containing external config data, like cell addresses and their corresponding output labels.
    """
    with open(file, "r", encoding="utf8") as f:
        return json.loads(f.read())


def find_files_included(directory, include_pattern):
    all_files = []
    for subdir in directory.iterdir():
        if subdir.is_dir() and re.search(include_pattern, str(subdir)):
            console.print(f"Scanning offers directory {subdir.as_posix()}...")
            files = glob.iglob(subdir.as_posix() + "/**/[!~$]*.xlsx", recursive=True)
            all_files.extend(files)
    return [f for f in all_files]
    # for file in files:
    #     yield file


def auto_format_cell_width(ws):
    for letter in range(1, ws.max_column):
        maximum_value = 0
        for cell in ws[get_column_letter(letter)]:
            val_to_check = len(str(cell.value))
            if val_to_check > maximum_value:
                maximum_value = val_to_check
        ws.column_dimensions[get_column_letter(letter)].width = maximum_value + 2


def create_style(json_file):
    # Load styles from JSON file
    with open(json_file) as file:
        style_data = json.load(file)

    style_dict = {}
    for k, v in style_data.items():
        if "font" in v and "fill" in v and "alignment" in v and "number_format" in v:
            # Create Named Style
            named_style = NamedStyle(name=k)

            # Set font
            named_style.font = Font(
                name=v["font"]["name"],
                size=v["font"]["size"],
                bold=v["font"]["bold"],
                italic=v["font"]["italic"],
                color=v["font"]["color"],
            )

            # Set fill
            named_style.fill = PatternFill(
                patternType=v["fill"]["pattern"], fgColor=v["fill"]["fgColor"]
            )

            # Set alignment
            named_style.alignment = Alignment(
                horizontal=v["alignment"]["horizontal"],
                vertical=v["alignment"]["vertical"],
            )

            # Set number format
            named_style.number_format = v["number_format"]

            # Set wrap_text if it is defined
            if "wrap_text" in v:
                named_style.alignment.wrap_text = v["wrap_text"]

            style_dict[k] = named_style

    return style_dict


def apply_styles(ws, style_dict: dict, style_ranges: dict, autofit: bool = True):
    # Register named styles to the workbook
    for named_style in style_dict.values():
        if named_style.name not in ws.parent.named_styles:
            ws.parent.add_named_style(named_style)

    console.print("Applying named styles...")
    # Apply other styles
    for style_name, ranges in style_ranges.items():
        if style_name in style_dict:  # Check if the style is defined
            for cell_range in ranges:
                range_split = cell_range.split(":")
                min_cell = coordinate_from_string(range_split[0])
                min_col, min_row = column_index_from_string(min_cell[0]), min_cell[1]
                if len(range_split) > 1:  # If the range includes more than one cell
                    max_cell = coordinate_from_string(range_split[1])
                    max_col, max_row = (
                        column_index_from_string(max_cell[0]),
                        max_cell[1],
                    )
                else:  # If the range includes only one cell
                    max_col, max_row = min_col, min_row
                for row in ws.iter_rows(
                    min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col
                ):
                    for cell in row:
                        cell.style = style_dict[style_name]

    # Zoom out
    ws.sheet_view.zoomScale = 75

    # Autoadjust width for each column depending on its content
    if autofit:
        auto_format_cell_width(ws)


def create_custom_chart(
    workbook_path: WindowsPath,
    target_sheet: str,
    data: list,
    chart_type: str,
    chart_style: int,
    position: str,
    **kwargs,
):
    match chart_type:
        case "bar":
            chart = BarChart()
        case "line":
            chart = LineChart()
    workbook = openpyxl.load_workbook(workbook_path, data_only=True)
    ws = workbook[target_sheet]
    left, right, up, down = data
    chart_data = Reference(ws, min_col=left, max_col=right, min_row=up, max_row=down)
    chart.add_data(chart_data, titles_from_data=True)
    chart.style = chart_style
    ws.add_chart(chart, position)
    workbook.save(workbook_path)
    return


def create_ddb_table(df: DataFrame, db_file: str, **params):
    console.print(f"Creating table into DuckDB file {db_file}...")
    table_name = params.get("table_name")
    table_schema = params.get("table_schema")
    query_file = params.get("query_file")
    df.to_clipboard()
    with duckdb.connect(db_file) as db:
        db.register(f"{table_name}_temp", df)
        if not all([table_name, query_file]):
            console.print("table_name and query_file must be specified before running.")
            return

        console.print(
            f"Creating table {table_name} in {table_schema} from temp data..."
        )
        db.execute(
            f"create or replace table {table_schema}.{table_name} as select * from {table_name}_temp x"
        )
        # Read file and split queries
        console.print("Fixing data...")
        with open(query_file, "r", encoding="utf8") as f:
            queries = f.read().split(";")

        # Iterate over each query
        for query in queries:
            # Skip empty queries
            if not query.strip():
                continue

            # Replace placeholders with parameters
            for placeholder, value in params.items():
                query = query.replace("{" + placeholder + "}", value)

            # Execute query
            console.print("Executing query:", query)
            db.execute(query)

    return


def timing(f):
    def wrap(*args, **kwargs):
        start_time = time.perf_counter()
        ret = f(*args, **kwargs)
        end_time = time.perf_counter() - start_time
        console.print(f"Total time elapsed: {end_time:0.3f} seconds")
        return ret

    return wrap


def write_output(
    output_file: str,
    data: dict[pl.DataFrame | str],
    style_specs: str,
    style_ranges: str,
    start_row: int,
    title: str,
    **kwargs,
):
    """
    Writes the data to an output Excel workbook.
    """
    workbook = openpyxl.Workbook()
    # Remove the default sheet created and add new sheets as per data keys
    default_sheet = workbook.active
    workbook.remove(default_sheet)
    for sheet_name, dataframe in data.items():
        sheet = workbook.create_sheet(sheet_name)
        total_rows, total_columns = dataframe.shape
        last_column_as_letter = get_column_letter(total_columns)

        # Writing workbook creation date, title and description
        sheet.cell(column=1, row=1, value=title)
        sheet.cell(column=1, row=2, value="Created on:")
        sheet.cell(column=2, row=2, value=datetime.datetime.now())
        sheet.cell(column=1, row=3, value="Created by:")
        sheet.cell(column=2, row=3, value=os.environ.get("USERNAME"))
        sheet.cell(
            column=1,
            row=(start_row - 1),
            value=f"=COUNTA(A{start_row + 1}:A{start_row + total_rows})",
        )

        # Writing data from dataframe to sheet starting from start_row
        for i, row in enumerate(
            dataframe_to_rows(dataframe, index=False, header=True), 1
        ):
            for j, cell in enumerate(row, 1):
                cell = str(tuple(cell)) if isinstance(cell, list) else cell
                sheet.cell(row=i + start_row - 1, column=j, value=cell)

        autofit_check = kwargs.get("autofit", True)
        apply_styles(sheet, style_specs, style_ranges, autofit_check)
        filters = sheet.auto_filter
        filters.ref = f"A{start_row}:{last_column_as_letter}{total_rows}"
        sheet.freeze_panes = f"B{start_row + 1}"

    console.print("Saving output file")
    workbook.save(output_file)
    console.print(f"File saved in: {output_file}")


def find_shtname_from_pattern(sheet_list: List[str], sheet_pattern: str) -> str:
    p = re.compile(sheet_pattern, flags=re.IGNORECASE)
    x = list(filter(None, map(p.search, sheet_list))).pop()
    return x.string  # attribute of a re.match object


def get_valid_data(
    sheet_as_values_list: List[str], idx_column_to_count: int
) -> List[str]:
    hdrs = sheet_as_values_list[0]
    valid_headers = list(filter(lambda x: len(x) > 0, hdrs))
    raw_data = list(zip(*sheet_as_values_list[1:]))[: len(valid_headers)]
    actual_height = len(list(filter(lambda x: x != "", raw_data[1])))
    # check the lenght of non-empty items in the first or second column, then crop the data up to that number
    valid_data = list(zip(*sheet_as_values_list[1 : actual_height + 1]))[
        : len(valid_headers)
    ]
    return valid_headers, valid_data


def get_idx_of_pattern_col(pat: str, sheet_as_values_list: List[str]) -> int:
    col_pattern = re.compile(pat, re.IGNORECASE)
    [winner] = list(filter(None, map(col_pattern.search, sheet_as_values_list[0])))
    # return sheet_as_values_list.index(winner)
    return sheet_as_values_list[0].index(winner.string)


def retrieve_all_info(list_of_files: List[str], config_file: str) -> pl.DataFrame:
    data = []
    final_data = []
    # ./conf/xy_offer_labels.json
    conf = load_json_config(config_file)

    start = time.perf_counter()

    for idx, item in enumerate(list_of_files):
        file_name_short = os.path.basename(item)
        print(f"Opening file {idx}: {item}")
        wb_inmemory = CalamineWorkbook.from_path(item)
        nombre_hoja_ficha = find_shtname_from_pattern(wb_inmemory.sheet_names, "ficha")
        # Implementar regex y obtener indice
        data = wb_inmemory.get_sheet_by_name(nombre_hoja_ficha).to_python(
            skip_empty_area=False
        )
        erres = {}
        for label, coords in conf.items():
            x, y = coords
            # print(f"{label} is:", wb_prueba[x][y])
            erres[label] = data[x][y]

        # part 2, the sap data
        nombre_hoja_sap = find_shtname_from_pattern(wb_inmemory.sheet_names, "sap")
        hoja = CalamineWorkbook.from_path(item).get_sheet_by_name(nombre_hoja_sap)

        if hoja.total_height > 1:
            rows_to_select = hoja.total_height
        else:
            rows_to_select = None
        data_rows = hoja.to_python(skip_empty_area=False, nrows=rows_to_select)
        idx_of_ur_col = get_idx_of_pattern_col("registral", data_rows)
        _, valid_data = get_valid_data(data_rows, idx_of_ur_col)
        df_data = (
            pl.DataFrame(erres)
            .with_columns(
                [
                    pl.when(pl.col(pl.Utf8).str.len_bytes() == 0)
                    .then(None)
                    .otherwise(pl.col(pl.Utf8))
                    .name.keep()
                ]
            )
            # casting, relaxed first then enforced
            .cast(pl.String)
            .cast(tipos_datos, strict=False)
            .cast(second_pass_cast)
        )

        # get the sap data metrics/aggregates
        df_sap = pl.DataFrame(valid_data)
        try:
            df_sap = df_sap.select(
                pl.col(f"column_{idx_of_ur_col}")
                .cast(pl.UInt32)
                .implode()
                .list.unique()
                .alias("urs_unicos"),
                pl.col(f"column_{idx_of_ur_col}")
                .implode()
                .list.unique()
                .list.len()
                .cast(pl.UInt16)
                .alias("total_urs"),
            )
        except pl.ComputeError as e:
            print(df_sap)
            print(idx_of_ur_col)
            raise e
        united_df = pl.concat([df_data, df_sap], how="horizontal").with_columns(
            pl.col("delegate").str.to_titlecase().name.keep(),
            pl.col("client_name").str.to_titlecase().name.keep(),
            pl.col("svh_recommendation").str.to_titlecase().name.keep(),
            pl.col("client_email").str.to_lowercase().name.keep(),
            full_path=pl.lit(item),
            file_name=pl.lit(file_name_short),
        )

        final_data.append(united_df)

    print(
        f"Done. Loaded {len(list_of_files)} files in {time.perf_counter() - start} seconds"
    )
