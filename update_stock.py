import duckdb
from rich.console import Console

from conf.functions import create_style
from conf.settings import stockconf as conf
from conf.settings import styles_file
from update_offers import write_output

con = Console()

# Instanciar las variables de ficheros, carpetas y otros
stylesheet = create_style(styles_file)

con.print("Accessing data...")

lote_queries = []
datos = {}

with duckdb.connect(conf.db_file) as db:
    # for tipo_agregacion in ["Wholesale", "Wholesale - from Retail", "No longer in Wholesale"]:
    # # Iterate over each query
    #     with open("./queries/stock_query.sql", encoding="utf8") as sql_file:
    #         query = sql_file.read()
    #         if not query.strip():
    #             continue
    #
    #     # Replace placeholders with parameters
    #     query = query.replace("{{placeholder}}", "'" + tipo_agregacion + "'")
    #
    #     # Execute query
    #     print(f"Filtering for channel: {tipo_agregacion}")
    #     query_as_df = db.execute(query).df()
    #     lote_queries.append(query_as_df)
    #     sheet_data = {tipo_agregacion: query_as_df}
    #     datos |= sheet_data

    con.print("Creating/updating stock table in database...")
    with open("./queries/stock_data.sql", encoding="utf8") as stock_table_query:
        query = stock_table_query.read()
    db.execute(f"CREATE OR REPLACE TABLE stock AS {query}")
    con.print("✅ Done!")
    # Añadido hasta que podamos partir el perimetro correctamente (ex-WS, etc.)
    query_as_df = db.sql("select * from stock").df()
    query_as_df.updated_at = query_as_df.updated_at.dt.tz_localize(None)
    sheet_data = {"Wholesale": query_as_df}
    datos |= sheet_data

rows = datos["Wholesale"].shape[0]


custom_styles = {
    "default": [
        f"A{conf.header_start}:AQ{conf.header_start + rows}",
    ],
    "header": [
        f"A{conf.header_start}:AQ{conf.header_start}",
    ],
    "percents": [
        f"F{conf.header_start + 1}:H{conf.header_start + rows}",
    ],
    "dates": [
        f"R{conf.header_start + 1}:T{conf.header_start + rows}",
        f"AQ{conf.header_start + 1}:AQ{conf.header_start + rows}",
    ],
    "data": [
        f"U{conf.header_start + 1}:W{conf.header_start + rows}",
        f"Z{conf.header_start + 1}:AC{conf.header_start + rows}",
        f"AE{conf.header_start + 1}:AH{conf.header_start + rows}",
        f"AJ{conf.header_start + 1}:AL{conf.header_start + rows}",
    ],
    "input": ["B3"],
    "title": ["A1"],
    "subtitle": ["A2:A3"],
}

con.print("Creating Excel output file...")

write_output(
    conf.get_output_path(),
    datos,
    stylesheet,
    custom_styles,
    conf.header_start,
    conf.sheet_name,
    # reuse_latest_file=True,
    # autofit=False
)
