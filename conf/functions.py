import glob
import json
import re
import time
from pathlib import WindowsPath

import duckdb
import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, NamedStyle, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string
from pandas import DataFrame
from rich.console import Console

console = Console()


def load_json_config(file):
    """
    Loads a JSON file containing external config data, like cell addresses and their corresponding output labels.
    """
    with open(file, "r", encoding="utf8") as f:
        return json.load(f)


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

        console.print(f"Creating table {table_name} in {table_schema} from temp data...")
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
