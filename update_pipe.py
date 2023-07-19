from conf.settings import pipeconf as conf, styles_file
from rich.console import Console
from conf.functions import create_style, create_custom_chart
from update_offers import write_output
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

db = duckdb.connect(conf.db_file)

console.print(f"Cargando fichero de pipe: {latest_pipe_file}")
pipe_data = pd.read_excel(latest_pipe_file, sheet_name="PIPE", skiprows=2, usecols="A:AE")

console.print("Guardando tabla de pipe en base de datos...")
db.register("pipe", pipe_data)

with open("./queries/pipe_aggregates.sql", encoding="utf8") as sql_file:
    query = sql_file.read()

agg_data_offers = db.execute(query).df()

# Variables, columnas y nombres
PK_PIPE = "ID Offer"
PK_AGG_DATA = "offerid"
OFFER_PRICE_COLUMN = "Importe Oferta Aprobada / Pdte aprobar"
LSEV_COLUMN = "lsev_offer"
PPA_COLUMN = "ppa_offer"
ASSET_TYPE_COLUMN = "Tipo Inmueble Agrupado Coral"
OFFER_PROB_COLUMN = "Offer probability"
LSEV_DELTA_COLUMN = "Delta % LSEV"
PPA_DELTA_COLUMN = "Delta % PPA"
YEAR_DEED_COLUMN = "Year Planned EP"
QUARTER_DEED_COLUMN = "Q Planned EP"
MONTH_SALE_COLUMN = "Month Sale EP"

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
    .assign(
        LSEV_DELTA_COLUMN=lambda df_: df_[OFFER_PRICE_COLUMN] / df_[LSEV_COLUMN] - 1,
        PPA_DELTA_COLUMN=lambda df_: df_[OFFER_PRICE_COLUMN] / df_[PPA_COLUMN] - 1
    )
    .rename(columns=lambda c: c.replace("\n", " "))
)

data = {conf.sheet_name: merged}

console.print("Creando strats...")

# Create a pivot table from the DataFrame
pivot_table1 = pd.pivot_table(
    merged,
    index=[OFFER_PROB_COLUMN, ASSET_TYPE_COLUMN],
    columns=[YEAR_DEED_COLUMN, QUARTER_DEED_COLUMN],
    values=[LSEV_COLUMN, PPA_COLUMN],
    aggfunc="sum",
)
pivot_table2 = pd.pivot_table(
    merged,
    index=[MONTH_SALE_COLUMN],
    values=[OFFER_PRICE_COLUMN, LSEV_COLUMN, PPA_COLUMN],
    aggfunc={OFFER_PRICE_COLUMN: "sum", LSEV_COLUMN: "sum", PPA_COLUMN: "sum"},
).assign(
    LSEV_DELTA_COLUMN=lambda df_: (df_[OFFER_PRICE_COLUMN] / df_[LSEV_COLUMN] - 1) * 100,
    PPA_DELTA_COLUMN=lambda df_: (df_[OFFER_PRICE_COLUMN] / df_[PPA_COLUMN] - 1) * 100
)

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


with pd.ExcelWriter(conf.get_output_path(), engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
    console.print(f"Añadiendo strats al fichero: {conf.get_filename()}")
    # Write the first pivot table to an Excel file, starting at cell B3 of the 'PivotTable' sheet
    pivot_table1.to_excel(writer, sheet_name=strat_sheet, startrow=2, startcol=1)
    # Write the second pivot table to the same Excel file, starting a few rows below the first pivot table
    startrow = pivot_table1.shape[0] + 6  # 3 rows gap between the two tables
    pivot_table2.to_excel(writer, sheet_name=strat_sheet, startrow=startrow, startcol=1)

data_y_size, data_x_size = pivot_table2.shape

console.print(f"Creando charts al fichero: {conf.get_filename()}")
create_custom_chart(
    conf.get_output_path(),
    strat_sheet,
    [3, data_x_size, startrow + 1, startrow + 1 + data_y_size],
    "bar",
    3,
    f"K15"
)
