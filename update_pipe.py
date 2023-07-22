from conf.settings import pipeconf as conf, styles_file
from rich.console import Console
from conf.functions import create_style, create_ddb_table
from update_offers import write_output
import re
import os
import duckdb
import pandas as pd

console = Console()

# Instanciar las variables de ficheros, carpetas y otros
strat_sheet = "Strats"
stylesheet = create_style(styles_file)

files = [p for p in conf.directory.rglob("*") if p.suffix in [".xlsx", ".xls"]]
sorted_files = sorted([f for f in files], key=os.path.getmtime)
latest_pipe_file = sorted_files[-1]

console.print("Obteniendo datos externos...")


console.print(f"Cargando fichero de pipe: {latest_pipe_file}")
pipe_data = (
            pd.read_excel(latest_pipe_file, sheet_name="PIPE", skiprows=2, usecols="A:AE")
        .rename(columns=lambda c: re.sub("[^A-Za-z0-9 ]", "", c).strip().replace(" ", "_").lower())
)

console.print("Guardando tabla de pipe en base de datos...")

with open("./queries/pipe_aggregates.sql", encoding="utf8") as sql_file:
    query = sql_file.read()

create_ddb_table(pipe_data, conf.db_file, table_name="pipeline", query_file="./queries/fix_pipe.sql")

with duckdb.connect(conf.db_file) as db:
    agg_data_offers = db.execute(query).df()

# Variables, columnas y nombres
PK_PIPE = "id_offer"
PK_AGG_DATA = "offerid"
LSEV_COLUMN = "lsev_offer"
PPA_COLUMN = "ppa_offer"

# Create a Pandas Excel writer using XlsxWriter as the engine.
cols_to_use = agg_data_offers.columns.difference(pipe_data.columns)

console.print("Agregando datos y completando pipe con datos externos...")

# Convert the DataFrame to an XlsxWriter Excel object.
merged = (
    pd.merge(
        pipe_data,
        agg_data_offers[cols_to_use],
        left_on=PK_PIPE,
        right_on=PK_AGG_DATA,
        how="left",
    )
    .drop(columns=[PK_AGG_DATA])
)

data = {conf.sheet_name: merged}

console.print("Creando strats...")

rows = merged.shape[0]
main_styles = create_style("./conf/styles.json")
custom_styles = {
    "default": [
        f"A{conf.header_start}:AQ{conf.header_start + rows}",
    ],
    "header": [
        f"A{conf.header_start}:AQ{conf.header_start}",
    ],
    "percents": [
        f"AP{conf.header_start + 1}:AQ{conf.header_start + rows}",
    ],
    "dates": [
        "B2",
        f"N{conf.header_start + 1}:N{conf.header_start + rows}",
        f"Q{conf.header_start + 1}:R{conf.header_start + rows}",
        f"V{conf.header_start + 1}:V{conf.header_start + rows}",
        f"Y{conf.header_start + 1}:Y{conf.header_start + rows}",
        f"AC{conf.header_start + 1}:AC{conf.header_start + rows}",
        f"AF{conf.header_start + 1}:AF{conf.header_start + rows}",
        f"AH{conf.header_start + 1}:AH{conf.header_start + rows}",
    ],
    "data": [
        f"H{conf.header_start + 1}:M{conf.header_start + rows}",
        f"AK{conf.header_start + 1}:AN{conf.header_start + rows}",
    ],
    "input": ["B3"],
    "title": ["A1"],
    "subtitle": ["A2:A3"],
}

# Create a Pandas Excel writer using openpyxl as the engine
write_output(
    conf.get_output_path(),
    data,
    main_styles,
    custom_styles,
    conf.header_start,
    conf.sheet_name
)

