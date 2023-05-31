from pathlib import Path
from datetime import datetime
import duckdb

directorio = Path("//EURFL01/advisors.hal/non-hudson/Coral Homes/CoralHudson/6. Stock/Parquet").absolute()

db = duckdb.connect()
db.execute(f"CREATE VIEW master_tape AS SELECT * FROM parquet_scan('{directorio}/master_tape.parquet')")
db.execute(f"CREATE VIEW offers_data AS SELECT * FROM parquet_scan('{directorio}/offers.parquet')")
db.execute(f"CREATE VIEW ws_segregated AS SELECT * FROM parquet_scan('{directorio}/disaggregated_assets.parquet')")

print("Accessing data...")
with open("./const/aggregate_query.sql", encoding="utf8") as sql_file:
    query = sql_file.read()

agg_data = db.execute(query).df()

print("Creating Excel output file...")
now_rendered = datetime.strftime(datetime.now(), "%d-%m-%Y")
agg_data.to_excel("stock_file.xlsx", index=False, sheet_name=f"Wholesale Stock {now_rendered}")
