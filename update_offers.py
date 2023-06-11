import re
import os
import openpyxl
import json
import duckdb
import pandas as pd
import pendulum as pdl
from datetime import datetime
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment, NamedStyle
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from rich.console import Console
from conf.settings import DIR_OUTPUT, DIR_PARQUET, DIR_INPUT_LOCAL

# Instanciar la consola bonita
console = Console()

# Instanciar las variables de ficheros, carpetas y otros
# directory = DIR_INPUT_LOCAL + "offer_files/"
directory = "//EURFL01/advisors.hal/non-hudson/Coral Homes/CoralHudson/1. AM/8. Wholesale Channel/Ofertas recibidas SVH"
name_structure = ["[!$~]*_OF_*", "[!$~][0-9]*[_ ]*"]
limit_files = None
extensions = [".xlsx"]
cell_address_file = "./conf/cell_addresses.json"
sap_mapping_file = "./conf/sap_columns_mapping.json"
date_append_output_name = datetime.strftime(datetime.now(), "%Y%m%d")
output_file = f"#{date_append_output_name}_Coral_Homes_Offers_Data.xlsx"
sheet_name = "Output Sheet"
header_start = 5


def load_json_config(file):
    """
    Loads a JSON file containing external config data, like cell addresses and their corresponding output labels.
    """
    with open(file, "r", encoding="utf8") as f:
        return json.load(f)


def get_files_in_directory(
    directory: str, name_structure: list, extensions: list, limit: int = None
):
    """
    Scans a directory for files with a specific name structure and extensions and returns a list of those file paths.
    """
    dir_path = Path(directory)
    console.print(f"Scanning folder {str(dir_path)} for offers...")
    files = []

    for pattern in name_structure:
        for ext in extensions:
            files.extend(dir_path.glob(f"**/{pattern}{ext}"))

    if limit:
        files = files[:limit]

    console.print("Filtering out non-valid files (Excel binary and others)")
    uniques = set(files)
    console.print(
        f"Found a total of {len(uniques)} files in the {str(dir_path)} folder."
    )
    return [str(file) for file in uniques]


def extract_cell_values(
    file, search_strings, table_columns: str, table_sheet: str = "SAP"
):
    """
    Opens an Excel file and converts the worksheet to a dictionary. It then looks for
    specific strings, moves to the relative cell addresses based on both x and y offsets,
    and extracts the value at that location. Also extracts unique values from specified
    columns in a table on another sheet.
    Includes the file name in the returned data.
    """
    try:
        console.print(f"[ Loading data from file ] > {file}")
        workbook = openpyxl.load_workbook(file, read_only=True, data_only=True)
    except Exception as e:
        console.print(f"Error when loading file {file}. Details: {e}")
        data = {}
        data["read_status"] = "Fail"
        data["read_details"] = str(e)
        data["full_path"] = os.path.abspath(file)
        data["file_name"] = os.path.basename(file)
        return data
    sheet = workbook.active
    data = {label: None for label in search_strings.keys()}
    data["full_path"] = os.path.abspath(file)
    data["file_name"] = os.path.basename(file)
    data["read_status"] = "Success"
    data["read_details"] = None

    # Convert the worksheet to a dictionary for faster searching
    sheet_dict = {
        (cell.row, cell.column): cell.value
        for row in sheet.iter_rows()
        for cell in row
        if cell.value is not None
    }

    for (cell_row, cell_col), cell_value in sheet_dict.items():
        for label, [regex, offset_y_range, offset_x_range] in search_strings.items():
            if re.match(
                regex, str(cell_value), re.IGNORECASE
            ):  # compare values using the regular expression
                for offset_y in offset_y_range:
                    for offset_x in offset_x_range:
                        try:
                            target_cell_address = (
                                cell_row - offset_y,
                                cell_col + offset_x,
                            )
                            target_cell_value = sheet_dict.get(target_cell_address)
                            if target_cell_value is not None:
                                data[label] = target_cell_value
                                break
                        except IndexError:
                            continue
                        else:
                            break
                    else:
                        continue
                    break

    # Load the JSON file for table columns and process table columns if specified
    if table_columns and table_sheet:
        columns_dict = load_json_config(table_columns)
        try:
            table_sheet = workbook[table_sheet]
            for col in columns_dict:
                # Assume the first row contains headers
                header_row = next(
                    table_sheet.iter_rows(min_row=1, max_row=1, values_only=True)
                )
                # Get index of the target column based on header
                try:
                    col_idx = header_row.index(col)
                except ValueError:
                    console.print(
                        f'Warning: Column header "{col}" not found in sheet. Skipping this column.'
                    )
                    continue
                # Extract unique values from the column
                column_values = [
                    cell[col_idx]
                    if not isinstance(cell[col_idx], str) or not cell[col_idx].isdigit()
                    else int(cell[col_idx].lstrip("0"))
                    for cell in table_sheet.iter_rows(min_row=2, values_only=True)
                ]
                # Remove duplicates by converting to a set, after filtering out the null values
                unique_values = set(filter(None, column_values))
                # If there's only one unique value, unpack it from the set
                if len(unique_values) == 1:
                    unique_values = unique_values.pop()
                    if isinstance(unique_values, str) and unique_values.isdigit():
                        unique_values = int(
                            unique_values.lstrip("0")
                        )  # Remove leading zeros for numeric strings
                else:
                    # Convert to list and sort to make output more predictable
                    try:
                        unique_values = sorted(unique_values)
                    except TypeError as e:
                        console.print(f"Error en fichero {file} debido a {e}")
                        data["read_status"] = "Warning"
                        data["read_details"] = str(e)
                        continue
                # Add the unique values to the output
                data[columns_dict.get(col)] = unique_values if unique_values else None
        except KeyError as e:
            console.print(f"Error encontrado: {e}\nFichero causante: {file}")
            data["read_status"] = "Warning"
            data["read_details"] = str(e)

    # Cleanup of bad data or strings
    for label, value in data.items():
        if label == "contract_deposit" and value == "-":
            data[label] = 0
        elif label == "client_description" and value == "NOMBRE":
            data[label] = None
        elif label == "offer_date":
            try:
                value = pdl.parse(value, strict=False).to_date_string()
            except Exception as e:
                console.print(f"Error al identificar la fecha de la oferta: >> {e}")
                if data["full_path"].startswith("\\\\EURFL01"):
                    posible_fecha = re.findall("\d{6,8}", data["full_path"])[0]
                    data[label] = datetime.strptime(posible_fecha, "%Y%m%d")
                else:
                    data[label] = datetime.strptime(
                        "".join(data["full_path"].split("\\")[-4:-2]), "%Y%m%d"
                    )

    return data


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


def apply_styles(ws, style_dict, cell_ranges):
    # Register named styles to the workbook
    for named_style in style_dict.values():
        if named_style not in ws.parent.named_styles:
            ws.parent.add_named_style(named_style)

    console.print("Applying styles to the output workbook...")
    # Apply default style to all cells if it exists
    if "default" in style_dict:
        for row in ws.iter_rows():
            for cell in row:
                cell.style = style_dict["default"]
                if (
                    cell.value
                    and isinstance(cell.value, str)
                    and os.path.exists(cell.value)
                ):
                    cell.hyperlink = cell.value  # set hyperlink
                    cell.style = style_dict["hyperlink"]

    # Apply other styles
    for style_name, ranges in cell_ranges.items():
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
    for columns in ws.columns:
        col = get_column_letter(columns[0].column)
        ws.column_dimensions[col].auto_size = True


def load_previous_data(previous_file: str, sheet: str = sheet_name):
    if os.path.isfile(previous_file):
        previous_df = pd.read_excel(
            previous_file, sheet_name=sheet, skiprows=header_start - 1, index_col=0
        )  # load the existing data
        previous_df.index = previous_df.index.fillna("n/a")
        previous_data = previous_df.iloc[
            :, header_start:
        ]  # select only the additional columns (assuming they start at column F)
    else:
        previous_data = (
            pd.DataFrame()
        )  # create an empty DataFrame for the case where no previous data exists
    return previous_data


def write_output(
    output_file: str,
    sheet_name: str,
    dataframe: pd.DataFrame,
    style_specs: str,
    style_ranges: str,
    start_row: int,
    use_previous_data: bool = False,
):
    """
    Writes the data to an output Excel workbook.
    """
    workbook = openpyxl.Workbook()
    # Remove the default sheet created and add new sheets as per data keys
    default_sheet = workbook.active
    workbook.remove(default_sheet)
    sheet = workbook.create_sheet(sheet_name)

    # Writing workbook creation date, title and description
    sheet.cell(column=1, row=1, value="Coral Homes Offers Data")
    sheet.cell(column=1, row=3, value="Created by ")
    sheet.cell(column=2, row=3, value=os.environ.get("USERNAME"))
    sheet.cell(column=3, row=3, value=" on the ")
    sheet.cell(column=4, row=3, value=datetime.now())

    if use_previous_data:
        previous_data = load_previous_data(output_file, sheet_name)
        previous_data.to_clipboard()

        if not previous_data.empty:
            dataframe = pd.concat(
                [dataframe, previous_data], axis=1
            )  # concatenate the new data and the previous data

    # Writing data from dataframe to sheet starting from start_row
    for i, row in enumerate(dataframe_to_rows(dataframe, index=False, header=True), 1):
        for j, cell in enumerate(row, 1):
            cell = str(tuple(cell)) if isinstance(cell, list) else cell
            sheet.cell(row=i + start_row - 1, column=j, value=cell)

    apply_styles(sheet, style_specs, style_ranges)

    console.print(f"Saving output file in: {output_file}")
    workbook.save(output_file)


def main():
    # TODO: Keep track of all the offers that have already been read in the file
    # We can do that by listing all the filenames in the Excel and compute the differente vs the found files
    # Create the output directory if not exists
    Path(DIR_OUTPUT).mkdir(exist_ok=True)
    # Extract the workbooks information one by one, then append the dictionary records to a 'data' variable
    cell_addresses = load_json_config(cell_address_file)
    files = get_files_in_directory(directory, name_structure, extensions, limit_files)
    console.print("Extracting cell values from files...")
    data = [
        extract_cell_values(file, cell_addresses, sap_mapping_file) for file in files
    ]

    # Use the 'data' variable to create a DataFrame structure which will be manipulated/modified
    console.print("Assembling offer data into a DataFrame...")
    df = pd.DataFrame(data)

    # Get the data from disk sources
    console.print("Getting data from portfolio management...")
    db = duckdb.connect()
    db.execute(
        f"CREATE VIEW master_tape AS SELECT * FROM parquet_scan('{DIR_PARQUET}/master_tape.parquet')"
    )
    db.execute(
        f"CREATE VIEW offers_data AS SELECT * FROM parquet_scan('{DIR_PARQUET}/offers.parquet')"
    )

    with open("./queries/offers_query.sql", encoding="utf8") as sql_file:
        query = sql_file.read()

    offers_data = db.execute(query).df()

    # Plug said data into the offers dataframe
    expanded_df = (
        df
        .merge(offers_data, how="left", left_on="offer_id", right_on="offerid")
        .assign(
            commercialdev=lambda df_: df_.commercialdev_x.fillna(df_.commercialdev_y),
            jointdev=lambda df_: df_.jointdev_x.fillna(df_.jointdev_y),
            unique_urs=lambda df_: df_.unique_urs_x.fillna(df_.unique_urs_y),
        )
        .drop(
            columns=[
                "offerid",
                "commercialdev_x",
                "commercialdev_y",
                "jointdev_x",
                "jointdev_y",
                "unique_urs_x",
                "unique_urs_y",
            ]
        )
        .sort_values(by=["offer_date"], ascending=False)
    )

    no_of_files = len(files)

    # Define your custom formatting schema here
    cell_ranges = {
        "header": [f"A{header_start}:BA{header_start}"],
        "data": [
            f"C{header_start + 1}:F{header_start + no_of_files + 1}",
            f"AY{header_start + 1}:BA{header_start + no_of_files + 1}",
        ],
        "dates": [f"B{header_start + 1}:B{header_start + no_of_files + 1}", "D3"],
        "input": ["B3"],
        "title": ["A1"],
        "subtitle": ["A3"],
    }

    write_output(
        DIR_OUTPUT + output_file,
        sheet_name,
        expanded_df,
        create_style("./conf/styles.json"),
        cell_ranges,
        5,
    )


if __name__ == "__main__":
    main()
