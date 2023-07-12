import os
import psycopg
import argparse
import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq


# Default connection values
host = os.environ["DBHOST"]
user = os.environ["DBUSER"]
pwd = os.environ["DBPWD"]
port = os.environ["DBPORT"]


def pgsql_to_parquet(host, database, user, password, table, output_file):
    conn_string = f"host='{host}' dbname='{database}' user='{user}' password='{password}'"

    try:
        # Establish a connection to the database
        conn = psycopg.connect(conn_string)

        # Create a new cursor object
        cur = conn.cursor()

        # Execute SQL command to fetch all records from the table
        cur.execute(f"SELECT * FROM {table}")

        # Fetch all rows from cursor
        rows = cur.fetchall()

        # Get the column names from the cursor object
        col_names = [desc[0] for desc in cur.description]

        # Create a Pandas DataFrame from the results
        df = pd.DataFrame(rows, columns=col_names)

        # Convert DataFrame to Apache Arrow Table
        arrow_table = pa.Table.from_pandas(df)

        # Write out the Table to Parquet format
        pq.write_table(arrow_table, output_file)

        print(f"Data from table {table} has been written to .parquet file successfully.", output_file, sep="\n")

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        # Close the cursor and connection
        cur.close()
        conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--table", required=True, help="Name of the table in the PostgreSQL database.")
    parser.add_argument("--filename", required=False, help="Name of the output file.", default=None)
    parser.add_argument("--path", required=False, help="Path to the output .parquet file.", default="N:/CoralHudson/6. Stock/Parquet/")
    args = parser.parse_args()

    if args.filename == None:
        actual_filename = (args.path + args.table + ".parquet")
    else:
        actual_filename = (args.path + args.filename + ".parquet")

    # Here, I'm using dummy values for host, database, user, and password.
    # Please replace them with your actual database credentials.
    pgsql_to_parquet(host, 'postgres', user, pwd, args.table, actual_filename)


if __name__ == "__main__":
    main()
