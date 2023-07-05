import re
import os
import openpyxl
import duckdb
import pandas as pd
import pendulum as pdl
import datetime
import uuid
from pathlib import Path
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from rich.console import Console
from conf.settings import BASE_DIR, offers_conf, sap_mapping_file, cell_address_file, styles_file
from conf.functions import (
    load_json_config,
    get_missing_values_by_id,
    find_files_included,
    apply_styles,
    create_style,
)

# Instanciar la consola bonita
console = Console()

# Instanciar las variables de ficheros, carpetas y otros
database_file = (BASE_DIR / "basedatos_wholesale.db").as_posix()
header_start = offers_conf.get("header_start")
output_sheet = offers_conf.get("sheet_name")
output_dir = offers_conf.get("output_dir")
output_file = "".join(["#", offers_conf.get("output_date"), offers_conf.get("output_file")])
stylesheet = create_style(styles_file)


def get_files_in_directory(
    directory: str, name_structure: list, extensions: list, limit: int = None
):
    """
    Scans a directory for files with a specific name structure and extensions and returns a list of those file paths.
    """
    dir_path = Path(directory)
    with console.status(f"Scanning folder {str(dir_path)} for offers...") as status:
        files = []

        for pattern in name_structure:
            for ext in extensions:
                for file in list(dir_path.glob(f"**/{pattern}{ext}")):
                    files.append(file)
                    status.update(f"{len(files)} ficheros leídos")

    if limit:
        files = files[:limit]

    console.print("Filtering out non-valid files (Excel binary and others)")
    uniques = set(files)
    console.print(
        f"Found a total of {len(uniques)} files in the {str(dir_path)} folder."
    )
    return [str(file) for file in uniques]


def extract_cell_values(
    file: str, search_strings: dict, columns_dict: dict, table_sheet: str = "SAP"
):
    """
    Opens an Excel file and converts the worksheet to a dictionary. It then looks for
    specific strings, moves to the relative cell addresses based on both x and y offsets,
    and extracts the value at that location. Also extracts unique values from specified
    columns in a table on another sheet.
    Includes the file name in the returned data.
    """
    try:
        workbook = openpyxl.load_workbook(file, read_only=True, data_only=True)
    except Exception as e:
        console.print(f"Error when loading file {file}. Details: {e}")
        data = {}
        data["read_status"] = "Fail"
        data["read_details"] = str(e)
        data["full_path"] = os.path.abspath(file)
        data["file_name"] = os.path.basename(file)
        return data

    data = {label: None for label in search_strings.keys()}

    try:
        hoja_ficha = list(
            filter(
                lambda z: re.match("ficha", z, flags=re.IGNORECASE), workbook.sheetnames
            )
        )
        hoja_ficha = hoja_ficha.pop() if hoja_ficha else "FICHA"
        sheet = workbook[hoja_ficha]
    except KeyError as e:
        data["full_path"] = os.path.abspath(file)
        data["file_name"] = os.path.basename(file)
        data["read_status"] = "Fail"
        data["read_details"] = e
        return data

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
        for label, [regex, offset_y, offset_x] in search_strings.items():
            if re.match(
                regex, str(cell_value), re.IGNORECASE
            ):  # compare values using the regular expression
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

    # Load the JSON file for table columns and process table columns if specified
    if columns_dict and table_sheet:
        try:
            try:
                table_sheet = workbook[table_sheet]
            except KeyError:
                table_sheet = workbook["Oferta"]
            for col in columns_dict:
                # Assume the first row contains headers
                header_row = next(
                    table_sheet.iter_rows(min_row=1, max_row=1, values_only=True)
                )
                # Get index of the target column based on header
                try:
                    col_idx = header_row.index(col)
                except ValueError:
                    # console.print(
                    #     f'Warning: Column header "{col}" not found in sheet. Skipping this column.'
                    # )
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
                if isinstance(value, str):
                    value = pdl.parse(value, strict=False).to_date_string()
                elif isinstance(value, datetime.datetime):
                    data[label] = value
            except Exception as e:
                console.print(f"Error al identificar la fecha de la oferta: >> {e}")
                if data["full_path"].startswith("\\\\EURFL01"):
                    posible_fecha = re.findall("\d{6,8}", data["full_path"])[0]
                    data[label] = datetime.datetime.strptime(posible_fecha, "%Y%m%d")
                else:
                    data[label] = datetime.datetime.strptime(
                        "".join(data["full_path"].split("\\")[-4:-1]), "%Y%m%d"
                    )

    return data



## TODO: Usar nombres de fichero por cada script, en variables como FILENAME_OUTPUT

def load_previous_data(sheet: str = output_sheet, start_row: int = header_start):

    latest_file_name_pattern = offers_conf.get("output_file")
    matching_files = []
    for f in Path(output_dir).glob("*.*"):
        if re.match(latest_file_name_pattern, f.name):
            matching_files.append(f)

    sorted_files = sorted([f for f in matching_files], key=os.path.getmtime)
    previous_file = sorted_files[-1]
    if os.path.isfile(previous_file):
        previous_df = pd.read_excel(
            previous_file, sheet_name=sheet, skiprows=start_row - 1, index_col=0
        )  # load the existing data
        previous_df.index = previous_df.unique_id
    else:
        previous_data = (
            pd.DataFrame()
        )  # create an empty DataFrame for the case where no previous data exists
    return previous_data

def enrich_offers(
    dataframe: pd.DataFrame,
    reuse_latest_file: bool = False
):

    # Ingest previous data
    if reuse_latest_file:
        previous_data = load_previous_data(output_sheet)
        joined_df = dataframe.join(previous_data)
    else:
        joined_df = dataframe

    console.print("Getting data from portfolio management...")
    with open("./queries/offers_query.sql", encoding="utf8") as sql_file:
        query = sql_file.read()

    with duckdb.connect(database_file) as db:
        offers_data = db.execute(query).df()


        expanded_df = (
            joined_df.reset_index().merge(offers_data, how="left", left_on="offer_id", right_on="offerid")
            .assign(
                commercialdev=lambda df_: df_.commercialdev_x.fillna(
                    df_.commercialdev_y
                ),
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
            .set_index("unique_id")
        )

        split_explode_columns = ["count_urs", "sum_ppa", "sum_lsev"]

        expanded_df[split_explode_columns] = expanded_df.pipe(
            get_missing_values_by_id, "unique_urs", db, "master_tape", "ur_current"
        )

    expanded_df = (expanded_df.assign(
            total_urs=lambda df_: df_.total_urs.fillna(df_.count_urs),
            ppa=lambda df_: df_.ppa.fillna(df_.sum_ppa),
            lsev_dec19=lambda df_: df_.lsev_dec19.fillna(df_.sum_lsev),
        )
        .drop(columns=split_explode_columns)
        .reset_index()
    )

    return expanded_df

def write_output(
    output_file: str,
    sheet_name: str,
    dataframe: pd.DataFrame,
    style_specs: str,
    style_ranges: str,
    start_row: int,
    title: str,
    **kwargs
):
    """
    Writes the data to an output Excel workbook.
    """
    workbook = openpyxl.Workbook()
    # Remove the default sheet created and add new sheets as per data keys
    default_sheet = workbook.active
    workbook.remove(default_sheet)
    sheet = workbook.create_sheet(sheet_name)
    total_rows, total_columns = dataframe.shape
    last_column_as_letter = get_column_letter(total_columns)

    # Writing workbook creation date, title and description
    sheet.cell(column=1, row=1, value=title)
    sheet.cell(column=1, row=2, value="Created on:")
    sheet.cell(column=2, row=2, value=datetime.datetime.now())
    sheet.cell(column=1, row=3, value="Created by:")
    sheet.cell(column=2, row=3, value=os.environ.get("USERNAME"))
    sheet.cell(column=1, row=(start_row - 1), value=f"=COUNTA(A{start_row + 1}:A{start_row + total_rows})")


    # Writing data from dataframe to sheet starting from start_row
    for i, row in enumerate(dataframe_to_rows(dataframe, index=False, header=True), 1):
        for j, cell in enumerate(row, 1):
            cell = str(tuple(cell)) if isinstance(cell, list) else cell
            sheet.cell(row=i + start_row - 1, column=j, value=cell)

    autofit_check = kwargs.get("autofit", True)
    apply_styles(sheet, style_specs, style_ranges, autofit_check)
    filters = sheet.auto_filter
    filters.ref = f"A{start_row}:{last_column_as_letter}{total_rows}"
    sheet.freeze_panes = f"B{start_row + 1}"
    console.print(f"Saving output file in: {output_file}")
    workbook.save(output_file)


def create_ddb_table(df: pd.DataFrame, db_file: str, **params):
    console.print(f"Creating table into DuckDB file {db_file}...")
    table_name = params.get("table_name")
    query_file = params.get("query_file")
    with duckdb.connect(db_file) as db:
        db.register(f"{table_name}_temp", df)
        if not all([table_name, query_file]):
            console.print("table_name and query_file must be specified before running.")
            return

        db.execute(
            f"create or replace table {table_name} as select * from {table_name}_temp"
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
            print(f"Executing query: {query}")
            db.execute(query)

    return


def main(update_offers: bool = False, current_year: bool = True):
    # TODO: Keep track of all the offers that have already been read in the file
    # We can do that by listing all the filenames in the Excel and compute the differente vs the found files

    # Create the output directory if not exists
    console.print(f"Creating path to files: {output_dir}")
    Path(output_dir).mkdir(exist_ok=True)

    all_duckdb_tables = ["ws_current_offers", "ws_hist_offers"]

    if update_offers:
        # Extract the workbooks information one by one, then append the dictionary records to a 'data' variable
        cell_addresses = load_json_config(cell_address_file)
        sap_columns_mapping = load_json_config(sap_mapping_file)
        if current_year:
            folder_pattern = r"2023"
            ddb_table_name = all_duckdb_tables[0]
        else:
            folder_pattern = r"20[12][^3]"
            ddb_table_name = all_duckdb_tables[1]

        files = find_files_included(offers_conf.get("directory"), folder_pattern)

        with console.status(f"Extracting cell values from files...") as status:
            data = []
            for file in files:
                file_shortened = file.name
                status.update(f"Loading data from file: [bold green]{file_shortened}")
                data.append(
                    extract_cell_values(file, cell_addresses, sap_columns_mapping)
                )
        # Use the 'data' variable to create a DataFrame structure which will be manipulated/modified
        console.print("Assembling offer data into a DataFrame...")
        df = (pd.DataFrame(data)
            )
        console.print("Creating UUIDs")
        df = df.assign(unique_id=lambda df_: df_.full_path.apply(lambda x: uuid.uuid5(uuid.NAMESPACE_DNS, x)))
        create_ddb_table(
            df,
            database_file,
            query_file="./queries/fix_offers.sql",
            table_name=ddb_table_name,
        )

    # Get the data from disk sources
    # TODO: 1) Traer los datos en una sola query, o varias y usar pandas para rellenar los que falten.
    # 2) Traer solamente los datos que tengan ya en su fichero de ofertas, más los que hayan escrito
    with duckdb.connect(database_file) as db:
        df = db.execute(" UNION ".join(["select * from " + t for t in all_duckdb_tables])).df()

    df["unique_id"] = df.unique_id.astype(str)
    df = df.set_index("unique_id")
    # Plug said data into the offers dataframe
    expanded_df = enrich_offers(df)
    no_of_files = expanded_df.shape[0]
    # Define your custom formatting schema here
    cell_ranges = {
        "default": [f"A{header_start}:AZ{header_start + no_of_files + 1}"],
        "header": [f"A{header_start}:AZ{header_start}"],
        "data": [
            f"D{header_start + 1}:G{header_start + no_of_files + 1}",
            f"AV{header_start + 1}:AW{header_start + no_of_files + 1}",
        ],
        "dates": [f"C{header_start + 1}:C{header_start + no_of_files + 1}"],
        "input": ["B3"],
        "title": ["A1"],
        "subtitle": ["A3"],
    }

    write_output(
        output_dir / output_file,
        output_sheet,
        expanded_df,
        stylesheet,
        cell_ranges,
        header_start,
        output_sheet,
        # reuse_latest_file=True,
        autofit=False
    )


if __name__ == "__main__":
    main()
    # main(update_offers=True, current_year=False)
    # main(update_offers=True)
