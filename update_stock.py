from conf.settings import DIR_OUTPUT
from conf.functions import create_style
from update_offers import write_output
from datetime import datetime
from rich.console import Console
import duckdb

con = Console()

con.print("Accessing data...")
with open("./queries/stock_query.sql", encoding="utf8") as sql_file:
    query = sql_file.read()

db = duckdb.connect("./basedatos_wholesale.db")
agg_data = db.execute(query).df()

header_start = 6
rows = agg_data.shape[0]

main_styles = create_style("./conf/styles.json")
custom_styles = {
    "default": [
        f"A{header_start}:AK{header_start + rows}",
    ],
    "header": [
        f"A{header_start}:AK{header_start}",
    ],
    "percents": [
        f"D{header_start + 1}:F{header_start + rows}",
    ],
    "dates": [
        f"Q{header_start + 1}:S{header_start + rows}",
    ],
    "data": [
        f"T{header_start + 1}:V{header_start + rows}",
        f"Y{header_start + 1}:AB{header_start + rows}",
        f"AD{header_start + 1}:AG{header_start + rows}",
        f"AI{header_start + 1}:AK{header_start + rows}",
    ],
    "input": ["B3"],
    "title": ["A1"],
    "subtitle": ["A2:A3"],
}

con.print("Creating Excel output file...")
now_sheet = datetime.strftime(datetime.now(), "%d-%m-%Y")
now_filename = datetime.strftime(datetime.now(), "%Y%m%d")
write_output(
    DIR_OUTPUT / f"#{now_filename}_Wholesale_Stock.xlsx",
    f"Wholesale Stock {now_sheet}",
    agg_data,
    main_styles,
    custom_styles,
    header_start,
    "Stock Data"
)
# agg_data.to_excel(DIR_OUTPUT / f"#{now_filename}_Wholesale_Stock.xlsx", index=False, sheet_name=f"Wholesale Stock {now_sheet}")
