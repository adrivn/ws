from conf.settings import DIR_PIPE, DIR_OUTPUT
from conf.functions import Reference, get_missing_values_by_id
from rich.console import Console
from conf.functions import create_style, create_custom_chart
from update_offers import write_output
from datetime import datetime
import os
import duckdb
import pandas as pd

console = Console()
files = [p for p in DIR_PIPE.rglob("*") if p.suffix in [".xlsx", ".xls"]]
sorted_files = sorted([f for f in files], key=os.path.getmtime)
latest_pipe_file = sorted_files[-1]

console.print("Obteniendo datos externos...")

db = duckdb.connect("./basedatos_wholesale.db")

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
ASSET_TYPE_COLUMN = "Tipo Inmueble Agrupado Coral"
OFFER_PROB_COLUMN = "Offer probability"
LSEV_DELTA_COLUMN = "Delta % LSEV"
YEAR_DEED_COLUMN = "Year Planned EP"
QUARTER_DEED_COLUMN = "Q Planned EP"
MONTH_SALE_COLUMN = "Month Sale EP"

now_filename = datetime.strftime(datetime.now(), "%Y%m%d")
FILENAME = f"#{now_filename}_WS_Pipeline.xlsx"
DATA_SHEET_NAME = "Pipeline 2023"
STRAT_SHEET_NAME = "Strats"

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
        LSEV_DELTA_COLUMN=lambda df_: df_[OFFER_PRICE_COLUMN] / df_[LSEV_COLUMN] - 1
    )
)

console.print("Creando strats...")

# Create a pivot table from the DataFrame
pivot_table1 = pd.pivot_table(
    merged,
    index=[OFFER_PROB_COLUMN, ASSET_TYPE_COLUMN],
    columns=[YEAR_DEED_COLUMN, QUARTER_DEED_COLUMN],
    values=LSEV_COLUMN,
    aggfunc="sum",
)
pivot_table2 = pd.pivot_table(
    merged,
    index=[MONTH_SALE_COLUMN],
    values=[OFFER_PRICE_COLUMN, LSEV_COLUMN],
    aggfunc={OFFER_PRICE_COLUMN: "sum", LSEV_COLUMN: "sum"},
).assign(
    LSEV_DELTA_COLUMN=lambda df_: (df_[OFFER_PRICE_COLUMN] / df_[LSEV_COLUMN] - 1) * 100
)

header_start = 6
rows = merged.shape[0]
main_styles = create_style("./conf/styles.json")
custom_styles = {
    "default": [
        f"A{header_start}:AO{header_start + rows}",
    ],
    "header": [
        f"A{header_start}:AO{header_start}",
    ],
    "percents": [
        f"AO{header_start + 1}:AO{header_start + rows}",
    ],
    "dates": [
        "B2",
        f"N{header_start + 1}:N{header_start + rows}",
        f"R{header_start + 1}:R{header_start + rows}",
        f"V{header_start + 1}:V{header_start + rows}",
        f"Y{header_start + 1}:Y{header_start + rows}",
        f"AC{header_start + 1}:AC{header_start + rows}",
        f"AF{header_start + 1}:AF{header_start + rows}",
        f"AH{header_start + 1}:AH{header_start + rows}",
    ],
    "data": [
        f"H{header_start + 1}:M{header_start + rows}",
        f"AK{header_start + 1}:AM{header_start + rows}",
    ],
    "input": ["B3"],
    "title": ["A1"],
    "subtitle": ["A2:A3"],
}

# Create a Pandas Excel writer using openpyxl as the engine
write_output(
    DIR_OUTPUT / FILENAME,
    DATA_SHEET_NAME,
    merged,
    main_styles,
    custom_styles,
    header_start,
    "Pipeline Data"
)


with pd.ExcelWriter(DIR_OUTPUT / FILENAME, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
    console.print(f"Guardando fichero Excel: {FILENAME}")
    # Write the first pivot table to an Excel file, starting at cell B3 of the 'PivotTable' sheet
    pivot_table1.to_excel(writer, sheet_name=STRAT_SHEET_NAME, startrow=2, startcol=1)
    # Write the second pivot table to the same Excel file, starting a few rows below the first pivot table
    startrow = len(pivot_table1) + 6  # 3 rows gap between the two tables
    pivot_table2.to_excel(writer, sheet_name=STRAT_SHEET_NAME, startrow=startrow, startcol=1)

data_y_size, data_x_size = pivot_table2.shape

create_custom_chart(
    DIR_OUTPUT / FILENAME,
    STRAT_SHEET_NAME,
    [2, data_x_size + 1, startrow + 1, startrow + 1 + data_y_size],
    "bar",
    3,
    f"K15"
)
