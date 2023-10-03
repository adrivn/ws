from conf.settings import stockconf as conf, styles_file
from conf.functions import create_style
from update_offers import write_output
from rich.console import Console
import duckdb

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
    query_as_df = db.execute(query).df()
    lote_queries.append(query_as_df)
    sheet_data = {"Wholesale": query_as_df}
    datos |= sheet_data

rows = datos["Wholesale"].shape[0]


custom_styles = {
    "default": [
        f"A{conf.header_start}:AK{conf.header_start + rows}",
    ],
    "header": [
        f"A{conf.header_start}:AK{conf.header_start}",
    ],
    "percents": [
        f"D{conf.header_start + 1}:F{conf.header_start + rows}",
    ],
    "dates": [
        f"Q{conf.header_start + 1}:S{conf.header_start + rows}",
    ],
    "data": [
        f"T{conf.header_start + 1}:V{conf.header_start + rows}",
        f"Y{conf.header_start + 1}:AB{conf.header_start + rows}",
        f"AD{conf.header_start + 1}:AG{conf.header_start + rows}",
        f"AI{conf.header_start + 1}:AK{conf.header_start + rows}",
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
# agg_data.to_excel(DIR_OUTPUT / f"#{now_filename}_Wholesale_Stock.xlsx", index=False, sheet_name=f"Wholesale Stock {now_sheet}")
