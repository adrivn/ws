import argparse
import datetime

from pathlib import Path

from rich.console import Console

from conf.functions import (
    create_ddb_table,
    create_style,
    find_files_included,
    retrieve_all_info,
)
from conf.settings import offersconf as conf
from conf.settings import styles_file

# Instanciar la consola bonita
console = Console()

# Instanciar las variables de ficheros, carpetas y otros
stylesheet = create_style(styles_file)


def main(
    update_offers: bool = False,
    current_year: bool = True,
):
    # TODO: Keep track of all the offers that have already been read in the file
    # We can do that by listing all the filenames in the Excel and compute the differente vs the found files

    # Create the output directory if not exists
    console.print(f"Creating path to files: {conf.output_dir}")
    Path(conf.output_dir).mkdir(exist_ok=True)

    all_duckdb_tables = ["current_offers", "hist_offers"]

    if update_offers:
        # Extract the workbooks information one by one, then append the dictionary records to a 'data' variable
        current_year_number = datetime.datetime.now().year
        if current_year:
            folder_pattern = rf"{current_year_number}"
            ddb_table_name = all_duckdb_tables[0]
        else:
            end_year = str(current_year_number)[-1]
            folder_pattern = rf"20[12][^{end_year}]"
            ddb_table_name = all_duckdb_tables[1]

        files = find_files_included(conf.directory, folder_pattern)

        console.print("Extracting cell values from files...")
        df = retrieve_all_info(files, "./conf/xy_offer_labels.json", 6)
        create_ddb_table(
            df,
            conf.db_file,
            query_file="./queries/fix_offers_new.sql",
            table_name=ddb_table_name,
            table_schema=conf.db_schema,
        )

    df.write_excel(conf.get_output_path(), conf.sheet_name, autofit=True)
    return "All okay!"


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--update",
        default=False,
        action="store_true",
        help="Whether or not to scan the update directory and update the offer data. Setting this to TRUE without the current_year option will update latest offers",
    )
    parser.add_argument(
        "--current",
        default=False,
        action="store_true",
        help="Whether to scan the current year offers directory and then update the file with the new information.",
    )
    args = parser.parse_args()
    main(
        update_offers=args.update,
        current_year=args.current,
    )
