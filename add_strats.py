from openpyxl import Workbook
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from pathlib import Path, WindowsPath
import re
import pandas as pd
import duckdb
import argparse
from conf.settings import offersconf as conf

file_to_extend = Path(conf.get_output_path())

def insert_dataframes_to_workbook(dataframes_dict: dict, output_file: WindowsPath) -> None:
    workbook = openpyxl.Workbook()
    # Remove the default sheet created and add new sheets as per data keys
    default_sheet = workbook.active
    workbook.remove(default_sheet)
    for sheet_name, (df, start_row, start_column) in dataframes_dict.items():
        sheet = workbook.create_sheet(sheet_name)
        rows = dataframe_to_rows(df, index=False, header=True)
        for r_idx, row in enumerate(rows, start=start_row):
            for c_idx, value in enumerate(row, start=start_column):
                sheet.cell(row=r_idx, column=c_idx, value=value)

    workbook.save(output_file)


def parse_sql_file(file_path: str) -> dict:
    queries = {}
    with open(file_path, 'r') as file:
        content = file.read()
        queries_list = re.split(r'--\s*(\w+)\s*\n', content)[1:]
        queries = {queries_list[i]: queries_list[i+1] for i in range(0, len(queries_list), 2)}
    return queries


def get_strat_from_query(query: str, connection: duckdb.DuckDBPyConnection) -> pd.DataFrame:
    return connection.execute(query).df()


if __name__ == "__main__":
    first_dict = parse_sql_file("./queries/strats.sql")
    output_dir = {}
    with duckdb.connect(conf.db_file) as db:
        for i, (k, v) in enumerate(first_dict.items(), start = 1):
            output_dir[k] = (get_strat_from_query(v, db), i * 6, 2)
    insert_dataframes_to_workbook(output_dir, WindowsPath("./_output/test_strats.xlsx"))
