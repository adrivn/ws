import duckdb
from pathlib import Path
from conf.settings import DIR_PARQUET, DATABASE_FILE

files = Path(DIR_PARQUET).glob("*.parquet")

def main():
    with duckdb.connect(DATABASE_FILE.as_posix()) as bd:
        for parquet_file in files:
            print(f"Updating database: {parquet_file.stem} ")
            bd.execute(f"CREATE OR REPLACE TABLE {parquet_file.stem} AS SELECT * FROM parquet_scan('{parquet_file}')")

if __name__ == "__main__":
    main()
