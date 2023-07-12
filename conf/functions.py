from pathlib import Path, WindowsPath
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, NamedStyle
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.utils import get_column_letter
from rich.console import Console
import duckdb
import openpyxl
import json
import re

console = Console()

def load_json_config(file):
    """
    Loads a JSON file containing external config data, like cell addresses and their corresponding output labels.
    """
    with open(file, "r", encoding="utf8") as f:
        return json.load(f)


def get_missing_values_by_id(
    df, column: str, connection: duckdb.DuckDBPyConnection, lookup_source: str, column_id_source: str
):
    # Step 1: Split the column into separate rows for each ID
    df = (
        df[column]
        .fillna(0)
        # .astype(str)
        # .str.replace("[()]", "", regex=True)
        # .str.replace("|", ",")
        .str.split(",")
        .explode()
    )
    # Step 2: Load the DataFrame into DuckDB
    query = f"""SELECT df.unique_id, 
    count(*) as count_urs, 
    sum(v.ppa) as sum_ppa,
    sum(v.lsev_dec19) as sum_lsev,
    FROM df
    LEFT JOIN {lookup_source} v ON df.{column} = v.{column_id_source}
    GROUP BY 1
    """
    with connection as con:
        con.register("df", df.reset_index())
        # Step 3: Join the DataFrame with the table containing the corresponding values for each ID
        # Step 4: Group by the original index and sum the corresponding values
        result = con.execute(query).df().set_index("unique_id")
        # Step 5: Return a Series with the sum of corresponding values for each original index
    return result


def find_files_included(directory, include_pattern):
    p = Path(directory)
    for subdir in p.iterdir():
        if subdir.is_dir() and re.search(include_pattern, str(subdir)):
            for file in subdir.glob("**/[!~$]*.xlsx"):
                if file.is_file():
                    yield file

def auto_format_cell_width(ws):
    for letter in range(1,ws.max_column):
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

    console.print(f"Applying named styles...")
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

def create_custom_chart(workbook_path: WindowsPath, target_sheet: str, data: list, chart_type: str, chart_style: int, position: str, **kwargs):
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

