import re
import os
import openpyxl
import duckdb
import pandas as pd
import pendulum as pdl
import datetime
import argparse
from pathlib import Path
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from rich.console import Console
from conf.settings import offersconf as conf, sap_mapping_file, cell_address_file, styles_file
from conf.functions import (
    load_json_config,
    find_files_included,
    apply_styles,
    create_style,
)

# Instanciar la consola bonita
console = Console()

# Instanciar las variables de ficheros, carpetas y otros
stylesheet = create_style(styles_file)


def extract_cell_values(
    file: str, search_strings: dict, columns_dict: dict
):
    """
    Opens an Excel file and converts the worksheet to a dictionary. It then looks for
    specific strings, moves to the relative cell addresses based on both x and y offsets,
    and extracts the value at that location. Also extracts unique values from specified
    columns in a table on another sheet.
    Includes the file name in the returned data.
    """
    try:
        workbook = openpyxl.load_workbook(file, data_only=True)
    except Exception as e:
        console.print(f"Error when loading file {file}. Details: {e}")
        data = {}
        data["read_status"] = "Fail"
        data["read_details"] = str(e)
        data["full_path"] = Path(file).as_posix()
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
        data["full_path"] = Path(file).as_posix()
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
    if columns_dict:
        regex = re.compile(r"\b(?:SAP|^Oferta)\b", re.IGNORECASE)
        posibles_hojas = list(filter(regex.search, workbook.sheetnames))
        hojas_sin_imagenes = []
        for sh in posibles_hojas:
            if not workbook[sh]._images:
                hojas_sin_imagenes.append(sh)
        try:
            hoja_detalle = hojas_sin_imagenes[0]
        except IndexError:
            hoja_detalle = posibles_hojas[0]
        try:
            table_sheet = workbook[hoja_detalle]
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
            console.print(f"{e}\nFichero causante: {file}")
            data["read_status"] = "Warning"
            data["read_details"] = str(e)

    # Cleanup of bad data or strings
        if data["offer_date"] is not None and not isinstance(data["offer_date"], datetime.datetime):
            try:
                data["offer_date"] = pdl.parse(data["offer_date"], strict=False).to_date_string()
            except Exception as e:
                console.print(f"Error al identificar la fecha de la oferta: >> {e}")

    return data



def load_previous_data(sheet: str = conf.sheet_name, start_row: int = conf.header_start):

    latest_file_name_pattern = conf.output_file
    matching_files = []
    for f in Path(conf.output_dir).glob("*.*"):
        if re.search(latest_file_name_pattern, f.name):
            matching_files.append(f)

    sorted_files = sorted([f for f in matching_files], key=os.path.getmtime)
    previous_file = sorted_files[-1]
    if os.path.isfile(previous_file):
        previous_data = pd.read_excel(
            previous_file, sheet_name=sheet, skiprows=start_row - 1
        )  # load the existing data
    else:
        previous_data = (
            pd.DataFrame()
        )  # create an empty DataFrame for the case where no previous data exists
    return previous_data

def enrich_offers(
    input_query: str,
    reuse_latest_file: bool = False
):

    console.print("Getting data from portfolio management...")
    with open("./queries/unnest_unique_urs.sql", encoding="utf8") as sql_file:
        query = sql_file.read()
        improved_query = f"CREATE TEMP TABLE unnested_data AS {query}"

    with duckdb.connect(conf.db_file) as db:
        # Ingest previous data
        db.execute("""set global pandas_analyze_sample=10000""")
        db.execute(improved_query)

        if reuse_latest_file:
            console.print("Opening latest offers file and retrieving additional columns...")
            previous_data = load_previous_data(conf.sheet_name)
            db.register("tmp_enrich_df", previous_data)
            excluded_columns_from_origin = ["unique_urs", "commercialdevs", "jointdevs", "offer_id"]
        else:
            db.execute(f"""create temp table tmp_enrich_df as {input_query}""")
            excluded_columns_from_origin = ["unique_urs", "commercialdev", "jointdev", "offer_id"]

        # Create enriched table, for later usage
        db.execute(f"""
            create or replace table offers_enriched_table as (
            with all_data as (
            select  t.* exclude({",".join(excluded_columns_from_origin) if len(excluded_columns_from_origin) > 1 else excluded_columns_from_origin.pop()}),
                    u.* exclude(unique_id)
            from tmp_enrich_df t 
            left join unnested_data u 
            on t.unique_id = u.unique_id
            ),
            filtered_columns as (
            select columns(x -> x not similar to '.+:1')
            from all_data)
            select * from filtered_columns);
            """)

        expanded_df = db.execute(
            """select unique_id, offer_id, * exclude(unique_id, offer_id) from offers_enriched_table order by offer_date desc"""
        ).df()


    return expanded_df


def write_output(
    output_file: str,
    data: dict[pd.DataFrame | str],
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

    console.print(f"Saving output file")
    workbook.save(output_file)
    console.print(f"File saved in: {output_file}")


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
            db.execute(query)

    return


def main(update_offers: bool = False, current_year: bool = True, reuse: bool = False, fix_data: bool = True):
    # TODO: Keep track of all the offers that have already been read in the file
    # We can do that by listing all the filenames in the Excel and compute the differente vs the found files

    # Create the output directory if not exists
    console.print(f"Creating path to files: {conf.output_dir}")
    Path(conf.output_dir).mkdir(exist_ok=True)

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

        files = find_files_included(conf.directory, folder_pattern)
        files_count = len(files)

        with console.status(f"Extracting cell values from files...") as status:
            data = []
            for idx, file in enumerate(files, start=1):
                status.update(f"[{idx}/{files_count}] ~ Loading data from file: [bold green]{file}")
                data.append(
                    extract_cell_values(file, cell_addresses, sap_columns_mapping)
                )
        # Use the 'data' variable to create a DataFrame structure which will be manipulated/modified
        console.print("Assembling offer data into a DataFrame...")
        df = (pd.DataFrame(data)
            )
        create_ddb_table(
            df,
            conf.db_file,
            query_file="./queries/fix_offers.sql",
            table_name=ddb_table_name,
        )

    # Get the data from disk sources
    with duckdb.connect(conf.db_file) as db:
        if not update_offers and fix_data:
            # Fix the data from the offers table, if it hasn't been already
            with open("./queries/fix_offers.sql", "r", encoding="utf8") as f:
                queries = f.read().split(";")

            # Iterate over each query
            for table in all_duckdb_tables:
                console.print(f"Fixing data from table {table}...")
                for query in queries:
                # Skip empty queries
                    if not query.strip():
                        continue
                    # Replace placeholders with parameters
                    new_query = query.replace("{table_name}", table)
                    # Execute query
                    db.execute(new_query)

        data = db.execute("select table_name from information_schema.tables where regexp_matches(table_name, 'ws_.+_offers')").fetchall()
        existing_tables = [d[0] for d in data]
        if len(existing_tables) > 1:
            query_para_crear_tablas = (" UNION ".join(
                [
                    """select * 
                    replace(
                    --list_aggregate(string_to_array(regexp_replace(address, '(\[|\])', '', 'g'), ','), 'string_agg', ' | ') as address,
                    list_aggregate(string_to_array(regexp_replace(asset_type, '(\[|\])', '', 'g'), ','), 'string_agg', ' | ') as asset_type,
                    --list_aggregate(string_to_array(regexp_replace(asset_location, '(\[|\])', '', 'g'), ','), 'string_agg', ' | ') as asset_location
                    ) 
                    from 
                    """ + t
                    for t in existing_tables
                ]
            ))
        else:
            query_para_crear_tablas = (
                "select * from " + existing_tables[0]
            )

    # Plug said data into the offers dataframe
    expanded_df = enrich_offers(query_para_crear_tablas, reuse_latest_file=reuse)
    no_of_files, no_of_variables = expanded_df.shape
    columns_start_at = 1
    first_column_letter, last_column_letter = get_column_letter(columns_start_at), get_column_letter(no_of_variables)
    datos = {conf.sheet_name: expanded_df}

    # Define your custom formatting schema here
    entire_range = f"{first_column_letter}{conf.header_start}:{last_column_letter}{conf.header_start + no_of_files}"
    header_range = f"{first_column_letter}{conf.header_start}:{last_column_letter}{conf.header_start}"

    cell_ranges = {
        "default": [entire_range],
        "header": [header_range],
        "data": [
            f"D{conf.header_start + 1}:G{conf.header_start + no_of_files}",
            f"AV{conf.header_start + 1}:AW{conf.header_start + no_of_files}",
        ],
        "dates": [f"C{conf.header_start + 1}:C{conf.header_start + no_of_files}"],
        "input": ["B3"],
        "title": ["A1"],
        "subtitle": ["A3"],
    }


    write_output(
        conf.get_output_path(),
        datos,
        stylesheet,
        cell_ranges,
        conf.header_start,
        conf.sheet_name,
        # reuse_latest_file=True,
        autofit=False
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--update",
        default=False,
        action="store_true",
        help="Whether or not to scan the update directory and update the offer data. Setting this to TRUE without the current_year option will update latest offers (2023 in this case)"
    )
    parser.add_argument(
        "--current", 
        default=False,
        action="store_true",
        help="Whether to update the current offers, meaning this year's (2023)"
    )
    parser.add_argument(
        "--reuse", 
        default=False,
        action="store_true",
        help="Use the latest offers file as base and bring any custom data columns that may have been added"
    )
    parser.add_argument(
        "--fix", 
        default=False,
        action="store_true",
        help="Use the latest offers file as base and bring any custom data columns that may have been added"
    )
    args = parser.parse_args()
    main(update_offers=args.update, current_year=args.current, reuse=args.reuse, fix_data=args.fix)
