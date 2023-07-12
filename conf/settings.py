from pathlib import Path
import datetime

# Configuracion de carpetas y directorios
BASE_DIR = Path().cwd()
CONF_DIR = BASE_DIR / "conf"
DIR_PARQUET = Path("N:/CoralHudson/6. Stock/Parquet").absolute()
DIR_OUTPUT_LOCAL = Path("_output/").absolute()
DIR_INPUT_LOCAL = Path("_attachments/").absolute()
DIR_OFFERS = Path("N:/CoralHudson/1. AM/8. Wholesale Channel/Ofertas recibidas SVH")
DIR_PIPE = Path("N:/CoralHudson/1. AM/8. Wholesale Channel/WS PIPE 2023")
DATABASE_FILE = Path("N:/CoralHudson/1. AM/8. Wholesale Channel/.code/database.db")
DATABASE_FILE_LOCAL = Path("./basedatos_wholesale.db")

database_file = DATABASE_FILE_LOCAL.as_posix()
cell_address_file = Path(CONF_DIR / "cell_addresses.json")
sap_mapping_file = Path(CONF_DIR / "sap_columns_mapping.json")
styles_file = Path(CONF_DIR / "styles.json")

stock_conf = {
    "directory": DIR_OFFERS,
    "output_dir": DIR_OFFERS.parent / ".outputs",
    "output_date": datetime.datetime.strftime(datetime.datetime.now(), "%Y%m%d"),
    "output_file": "_WS_Stock.xlsx",
    "sheet_name": "Wholesale Stock",
    "header_start": 6,
}

offers_conf = {
    "directory": DIR_OFFERS,
    "output_dir": DIR_OFFERS.parent / ".outputs",
    "output_date": datetime.datetime.strftime(datetime.datetime.now(), "%Y%m%d"),
    "output_file": "_Coral_Homes_Offers_Data.xlsx",
    "sheet_name": "Offers Data",
    "header_start": 6,
}

pipe_conf = {
    "directory": DIR_PIPE,
    "output_dir": DIR_PIPE.parent / ".outputs",
    "output_date": datetime.datetime.strftime(datetime.datetime.now(), "%Y%m%d"),
    "output_file": "_WS_Pipeline.xlsx",
    "sheet_name": "Pipeline 2023",
    "strats_sheet_name": "Strats",
    "header_start": 6,
}
