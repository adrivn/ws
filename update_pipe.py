from conf.settings import DIR_INPUT_LOCAL, DIR_OUTPUT, DIR_PARQUET
from conf.functions import unpack_single_item_lists
from pathlib import Path
from rich.console import Console
import os
import duckdb
import pandas as pd

console = Console()
files = [p for p in Path(DIR_INPUT_LOCAL / "pipe_files").rglob('*') if p.suffix in [".xlsx", ".xls"]]
sorted_files = sorted([f for f in files], key=os.path.getmtime)
latest_pipe_file = sorted_files[-1]

console.print("Obteniendo datos externos...")

db = duckdb.connect()
db.execute(f"CREATE VIEW master_tape AS SELECT * FROM parquet_scan('{DIR_PARQUET}/master_tape.parquet')")
db.execute(f"CREATE VIEW offers_data AS SELECT * FROM parquet_scan('{DIR_PARQUET}/offers.parquet')")
db.execute(f"CREATE VIEW sales2023 AS SELECT * FROM parquet_scan('{DIR_PARQUET}/sales2023.parquet')")

console.print(f"Cargando fichero de pipe: {latest_pipe_file}")
pipe_data = pd.read_excel(latest_pipe_file, sheet_name="PIPE")
agg_data_offers = db.execute(
    """select o.offerid, 
    array_agg(distinct o.ur_current::int) as unique_urs,
    array_agg(distinct o.commercialdev::int) as commercialdev,
    array_agg(distinct o.jointdev::int) as jointdev,
    array_agg(distinct o.offerstatus) as offer_status,
    max(s.saledate) as actual_sale_date, 
    max(s.commitmentdate) as commitment_date, 
    sum(m.lsev_dec19) as lsev_offer, 
    sum(m.ppa) as ppa_offer, 
    string_agg(distinct m.direccion_territorial, '|') as dts 
    from offers_data o
    left join master_tape m
    on o.ur_current = m.ur_current
    left join sales2023 s
    on o.ur_current = s.ur_5000
    group by all"""
).df()

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

FILENAME = "Pandas_Pivot.xlsx"
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
    .pipe(lambda x: x.applymap(unpack_single_item_lists))
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

# Create a Pandas Excel writer using openpyxl as the engine
with pd.ExcelWriter(DIR_OUTPUT / FILENAME, engine="openpyxl") as writer:
    console.print(f"Guardando fichero Excel: {FILENAME}")
    merged.to_excel(writer, sheet_name=DATA_SHEET_NAME, index=False)
    # Write the first pivot table to an Excel file, starting at cell B3 of the 'PivotTable' sheet
    pivot_table1.to_excel(writer, sheet_name=STRAT_SHEET_NAME, startrow=2, startcol=1)
    # Write the second pivot table to the same Excel file, starting a few rows below the first pivot table
    startrow = len(pivot_table1) + 6  # 3 rows gap between the two tables
    pivot_table2.to_excel(writer, sheet_name="Strats", startrow=startrow, startcol=1)
