import os
import re

import duckdb
import polars as pl
from rich.console import Console

from conf.functions import create_ddb_table, create_style
from conf.settings import pipeconf as conf
from conf.settings import styles_file
from update_offers import write_output

console = Console()

# Instanciar las variables de ficheros, carpetas y otros
strat_sheet = "Strats"
stylesheet = create_style(styles_file)

files = [
    p
    for p in conf.directory.rglob("*")
    if p.suffix in [".xlsx", ".xls"] and not p.name.startswith("~")
]
sorted_files = sorted([f for f in files], key=os.path.getmtime)
latest_pipe_file = sorted_files[-1]

console.print("Obteniendo datos externos...")


console.print(f"Cargando fichero de pipe: {latest_pipe_file}")
pipe_data = (
    pl.read_excel(
        latest_pipe_file,
        sheet_name="PIPE",
        engine="calamine",
        read_options={
            "header_row": 2,
            "skip_rows": 0,
            "use_columns": "A:AF",
        },
    )
    .rename(
        lambda c: re.sub(r"[^A-Za-z0-9\s]", "", c).strip().replace(" ", "_").lower()
    )
    .with_columns(
        pl.col("promo__ur")
        .str.split("/")
        .list.eval(pl.element().str.strip_chars().cast(pl.UInt32)),
        pl.col("id_offer").cast(pl.Int32),
    )
)

console.print("Guardando tabla de pipe en base de datos...")

with open("./queries/pipe_aggregates.sql", encoding="utf8") as sql_file:
    query = sql_file.read()

with duckdb.connect(conf.db_file) as db:
    agg_data_offers = db.execute(query).pl()

to_output = pipe_data.join(agg_data_offers, left_on="id_offer", right_on="offerid")

create_ddb_table(
    to_output.to_pandas(),
    conf.db_file,
    table_name="pipeline",
    table_schema=conf.db_schema,
)


console.print("Agregando datos y completando pipe con datos externos...")

data = {
    conf.sheet_name: to_output.with_columns(
        pl.col("promo__ur").list.eval(pl.element().cast(pl.Utf8)).list.join("-")
    ).to_pandas()
}

console.print("Creando strats...")

rows = to_output.shape[0]
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
    conf.sheet_name,
)
