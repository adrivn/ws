from conf.settings import DIR_OUTPUT, DIR_INPUT_LOCAL, DIR_PARQUET
from datetime import datetime
import duckdb

db = duckdb.connect()
db.execute(f"CREATE VIEW master_tape AS SELECT * FROM parquet_scan('{DIR_OUTPUT}/master_tape.parquet')")
db.execute(f"CREATE VIEW offers_data AS SELECT * FROM parquet_scan('{DIR_OUTPUT}/offers.parquet')")
db.execute(f"CREATE VIEW ws_segregated AS SELECT * FROM parquet_scan('{DIR_OUTPUT}/disaggregated_assets.parquet')")

print("Accessing data...")
with open("./queries/aggregate_query.sql", encoding="utf8") as sql_file:
    query = sql_file.read()

agg_data = db.execute(query).df()

print("Creating Excel output file...")
now_rendered = datetime.strftime(datetime.now(), "%d-%m-%Y")
agg_data.to_excel(DIR_OUTPUT / "stock_file.xlsx", index=False, sheet_name=f"Wholesale Stock {now_rendered}")
