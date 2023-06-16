import duckdb
from conf.settings import DIR_PARQUET

def main():
    bd = duckdb.connect("basedatos_wholesale.db")
    bd.execute(f"CREATE OR REPLACE TABLE master_tape AS SELECT * FROM parquet_scan('{DIR_PARQUET}/master_tape.parquet')")
    bd.execute(f"CREATE OR REPLACE TABLE sales2023 AS SELECT * FROM parquet_scan('{DIR_PARQUET}/sales2023.parquet')")
    bd.execute(f"CREATE OR REPLACE TABLE offers_data AS SELECT * FROM parquet_scan('{DIR_PARQUET}/offers.parquet')")
    bd.execute(f"CREATE OR REPLACE TABLE ws_segregated AS SELECT * FROM parquet_scan('{DIR_PARQUET}/disaggregated_assets.parquet')")

if __name__ == "__main__":
    main()
