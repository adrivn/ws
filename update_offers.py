import os
import openpyxl
import json
import pandas as pd
from datetime import datetime
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment, NamedStyle
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from rich.console import Console

# Instanciar la consola bonita
console = Console()


def load_json_config(file):
    """
    Loads a JSON file containing external config data, like cell addresses and their corresponding output labels.
    """
    with open(file, "r", encoding="utf8") as f:
        return json.load(f)


def get_files_in_directory(directory: str, name_structure: list, extensions: list):
    """
    Scans a directory for files with a specific name structure and extensions and returns a list of those file paths.
    """
    dir_path = Path(directory)
    console.print(f"Scanning folder {str(dir_path)} for offers...")
    files = []
    for pattern in name_structure:
        for ext in extensions:
            files.extend(dir_path.glob(f"**/{pattern}{ext}"))
    uniques = set(files)
    console.print(f"Found a total of {len(uniques)} files in the {str(dir_path)} folder.")
    return [str(file) for file in uniques]


def extract_cell_values(file, search_strings, table_columns: str, table_sheet: str = "SAP"):
    """
    Opens an Excel file and converts the worksheet to a dictionary. It then looks for
    specific strings, moves to the relative cell addresses based on both x and y offsets,
    and extracts the value at that location. Also extracts unique values from specified
    columns in a table on another sheet.
    Includes the file name in the returned data.
    """
    workbook = openpyxl.load_workbook(file, read_only=True, data_only=True)
    sheet = workbook.active
    data = {label: None for label in search_strings.keys()}
    data["file_name"] = os.path.basename(file)

    # Convert the worksheet to a dictionary for faster searching
    sheet_dict = {(cell.row, cell.column): cell.value for row in sheet.iter_rows() for cell in row if cell.value is not None}

    for (cell_row, cell_col), cell_value in sheet_dict.items():
        for label, [string, offset_y, offset_x] in search_strings.items():
            if cell_value == string:
                target_cell_address = (cell_row - offset_y, cell_col + offset_x)
                data[label] = sheet_dict.get(target_cell_address)

    # Load the JSON file for table columns and process table columns if specified
    if table_columns and table_sheet:
        columns_dict = load_json_config(table_columns)
        table_sheet = workbook[table_sheet]
        for col in columns_dict:
            # Assume the first row contains headers
            header_row = next(table_sheet.iter_rows(min_row=1, max_row=1, values_only=True))
            # Get index of the target column based on header
            try:
                col_idx = header_row.index(col)
            except ValueError:
                print(f'Warning: Column header "{col}" not found in sheet. Skipping this column.')
                continue
            # Extract unique values from the column
            column_values = [cell[col_idx] if not isinstance(cell[col_idx], str) or not cell[col_idx].isdigit() else int(cell[col_idx].lstrip("0")) for cell in table_sheet.iter_rows(min_row=2, values_only=True)]
            unique_values = set(filter(None, column_values)) # Remove duplicates by converting to a set, after filtering out the null values
            # If there's only one unique value, unpack it from the set
            if len(unique_values) == 1:
                unique_values = unique_values.pop()
                if isinstance(unique_values, str) and unique_values.isdigit():
                    unique_values = int(unique_values.lstrip('0'))  # Remove leading zeros for numeric strings
            else:
                # Convert to list and sort to make output more predictable
                unique_values = sorted(unique_values)
            # Add the unique values to the output
            data[columns_dict.get(col)] = unique_values if unique_values else None

    # Cleanup of bad data or strings
    for label, value in data.items():
        if label == "contract_deposit" and value == "-":
            data[label] = 0
        elif label == "client_description" and value == "NOMBRE":
            data[label] = None

    return data


def create_style(json_file):
    # Load styles from JSON file
    with open(json_file) as file:
        style_data = json.load(file)

    style_dict = {}
    for k, v in style_data.items():
        if 'font' in v and 'fill' in v and 'alignment' in v and 'number_format' in v:
            # Create Named Style
            named_style = NamedStyle(name=k)

            # Set font
            named_style.font = Font(name=v['font']['name'],
                                    size=v['font']['size'],
                                    bold=v['font']['bold'],
                                    italic=v['font']['italic'],
                                    color=v['font']['color'])

            # Set fill
            named_style.fill = PatternFill(patternType=v['fill']['pattern'],
                                           fgColor=v['fill']['fgColor'])

            # Set alignment
            named_style.alignment = Alignment(horizontal=v['alignment']['horizontal'],
                                              vertical=v['alignment']['vertical'])

            # Set number format
            named_style.number_format = v['number_format']

            # Set wrap_text if it is defined
            if 'wrap_text' in v:
                named_style.alignment.wrap_text = v['wrap_text']

            style_dict[k] = named_style

    return style_dict


def apply_styles(ws, style_dict, cell_ranges):
    # Register named styles to the workbook
    for named_style in style_dict.values():
        if named_style not in ws.parent.named_styles:
            ws.parent.add_named_style(named_style)

    console.print("Applying styles to the output workbook...")
    # Apply default style to all cells if it exists
    if 'default' in style_dict:
        for row in ws.iter_rows():
            for cell in row:
                cell.style = style_dict['default']

    # Apply other styles
    for style_name, ranges in cell_ranges.items():
        if style_name in style_dict:  # Check if the style is defined
            for cell_range in ranges:
                range_split = cell_range.split(':')
                min_cell = coordinate_from_string(range_split[0])
                min_col, min_row = column_index_from_string(min_cell[0]), min_cell[1]
                if len(range_split) > 1:  # If the range includes more than one cell
                    max_cell = coordinate_from_string(range_split[1])
                    max_col, max_row = column_index_from_string(max_cell[0]), max_cell[1]
                else:  # If the range includes only one cell
                    max_col, max_row = min_col, min_row
                for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                    for cell in row:
                        cell.style = style_dict[style_name]

    # Zoom out
    ws.sheet_view.zoomScale = 75

    # Autoadjust width for each column depending on its content
    for columns in ws.columns:
        col = get_column_letter(columns[0].column)
        ws.column_dimensions[col].auto_size = True




def write_output(output_file, sheet_name, dataframe, style_specs, style_ranges, start_row):
    """
    Writes the data to an output Excel workbook.
    """
    print("Printing input data:", dataframe, sep="\n\n")
    dataframe.to_clipboard()
    workbook = openpyxl.Workbook()
    # Remove the default sheet created and add new sheets as per data keys
    default_sheet = workbook.active
    workbook.remove(default_sheet)
    sheet = workbook.create_sheet(sheet_name)

    # Writing workbook creation date, title and description
    sheet.cell(column=1, row=1, value="Coral Homes Offers Data")
    sheet.cell(column=1, row=2, value=datetime.now())
    sheet.cell(column=1, row=3, value="Created by ")
    sheet.cell(column=2, row=3, value=os.environ.get("USERNAME"))


    # Writing data from dataframe to sheet starting from start_row
    for i, row in enumerate(dataframe_to_rows(dataframe, index=False, header=True), 1):
        for j, cell in enumerate(row, 1):
            cell = str(tuple(cell)) if isinstance(cell, list) else cell
            sheet.cell(row=i + start_row - 1, column=j, value=cell)

    apply_styles(sheet, style_specs, style_ranges)

    console.print(f"Saving output file in: {output_file}")
    workbook.save(output_file)


def main():
    directory = "./_attachments/offer_files/"
    # directory = "//EURFL01/advisors.hal/non-hudson/Coral Homes/CoralHudson/1. AM/8. Wholesale Channel/Ofertas recibidas SVH"
    name_structure = ["[!$~]*_OF_*", "[!$~][0-9]*[_ ]*"]
    extensions = [".xlsx", ".xls"]
    cell_address_file = "./const/cell_addresses.json"
    sap_mapping_file = "./const/sap_columns_mapping.json"
    output_file = "output.xlsx"
    sheet_name = "Output Sheet"

    # Read the list of files ALREADY downloaded and parsed
    # files_already_read = None

    # Extract the workbooks information one by one, then append the dictionary records to a 'data' variable
    cell_addresses = load_json_config(cell_address_file)
    files = get_files_in_directory(directory, name_structure, extensions)
    console.print("Extracting cell values from files...")
    data = [extract_cell_values(file, cell_addresses, sap_mapping_file) for file in files]

    # Use the 'data' variable to create a DataFrame structure which will be manipulated/modified
    console.print("Assembling offer data into a DataFrame...")
    df = pd.DataFrame(data)
    # Save the current offers data to disk
    # console.print("Saving data to disk...")
    # df.to_pickle(f"//EURFL01/advisors.hal/non-hudson/Coral Homes/CoralHudson/6. Stock/Parquet/#{datetime.now().strftime('%Y%m%d')}_Bundled_Offers_Data.pickle")

    # Manipulate the data
    expanded_df = df

    no_of_files = len(files)
    header_start = 5

    # Define your custom formatting schema here
    cell_ranges = {
        "header": [f"A{header_start}:AE{header_start}"],
        "data": [f"C{header_start + 1}:F{header_start + no_of_files + 1}"],
        "dates": [f"B{header_start + 1}:B{header_start + no_of_files + 1}", "A2"],
        "input": ["B3"],
        "title": ["A1"],
        "subtitle": ["A3"],
    }

    write_output(output_file, sheet_name, expanded_df, create_style("./const/styles.json"), cell_ranges, 5)


if __name__ == "__main__":
    main()
