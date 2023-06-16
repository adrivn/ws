from conf.settings import DIR_OUTPUT
from datetime import datetime
import duckdb

print("Accessing data...")
with open("./queries/stock_query.sql", encoding="utf8") as sql_file:
    query = sql_file.read()

db = duckdb.connect("./basedatos_wholesale.db")
agg_data = db.execute(query).df()

print("Creating Excel output file...")
now_sheet = datetime.strftime(datetime.now(), "%d-%m-%Y")
now_filename = datetime.strftime(datetime.now(), "%Y%m%d")
agg_data.to_excel(DIR_OUTPUT / f"#{now_filename}_Wholesale_Stock.xlsx", index=False, sheet_name=f"Wholesale Stock {now_sheet}")
