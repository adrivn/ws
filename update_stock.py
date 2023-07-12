from conf.settings import stock_conf, database_file, styles_file
from conf.functions import create_style
from update_offers import write_output
from datetime import datetime
from rich.console import Console
import duckdb

con = Console()

# TODO: 1) Crear multiples pestañas, una para stock vivo, otra para stock vendido, otra para stock que ya no esta en WS

# Instanciar las variables de ficheros, carpetas y otros
header_start = stock_conf.get("header_start")
directory = stock_conf.get("directory")
output_sheet = stock_conf.get("sheet_name")
output_dir = stock_conf.get("output_dir")
output_file = "".join(["#", stock_conf.get("output_date"), stock_conf.get("output_file")])
stylesheet = create_style(styles_file)

con.print("Accessing data...")

lote_queries = []
datos = {
}

with duckdb.connect(database_file) as db:
    for tipo_agregacion in ["Wholesale", "Wholesale - from Retail", "No longer in Wholesale"]:
    # Iterate over each query
        with open("./queries/stock_query.sql", encoding="utf8") as sql_file:
            query = sql_file.read()
            if not query.strip():
                continue

        # Replace placeholders with parameters
        query = query.replace("{{placeholder}}", "'" + tipo_agregacion + "'")

        # Execute query
        print(f"Filtering for channel: {tipo_agregacion}")
        query_as_df = db.execute(query).df()
        lote_queries.append(query_as_df)
        sheet_data = {tipo_agregacion: query_as_df}
        datos |= sheet_data

rows = datos["Wholesale"].shape[0]


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

write_output(
    output_dir / output_file,
    datos,
    stylesheet,
    custom_styles,
    header_start,
    output_sheet,
    # reuse_latest_file=True,
    # autofit=False
)
# agg_data.to_excel(DIR_OUTPUT / f"#{now_filename}_Wholesale_Stock.xlsx", index=False, sheet_name=f"Wholesale Stock {now_sheet}")
