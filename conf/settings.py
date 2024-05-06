from pathlib import Path, WindowsPath
from dataclasses import dataclass
import datetime

# Configuracion de carpetas y directorios
BASE_DIR = Path().cwd()
CONF_DIR = BASE_DIR / "conf"
QUERIES_DIR = BASE_DIR / "queries"
DIR_PARQUET = Path("N:/CoralHudson/6. Stock/Parquet").absolute()
DIR_OUTPUT_LOCAL = Path("_output/").absolute()
DIR_INPUT_LOCAL = Path("_attachments/").absolute()
DIR_OFFERS = Path("N:/CoralHudson/1. AM/8. Wholesale Channel/Ofertas recibidas SVH")
DIR_PIPE = Path("N:/CoralHudson/1. AM/8. Wholesale Channel/WS PIPE 2024")
DATABASE_FILE = Path("N:/CoralHudson/1. AM/8. Wholesale Channel/.code/database.db")
DATABASE_FILE_LOCAL = Path("./basedatos_wholesale.db")

database_file = DATABASE_FILE_LOCAL.as_posix()
cell_address_file = Path(CONF_DIR / "cell_addresses.json")
sap_mapping_file = Path(CONF_DIR / "sap_columns_mapping.json")
styles_file = Path(CONF_DIR / "styles.json")


@dataclass
class FileSettings:
    directory: str
    output_dir: str
    output_file: str
    sheet_name: str
    strat_sheet: str
    header_start: int
    areas_to_style: dict = None
    styles_file: WindowsPath = Path(CONF_DIR) / "styles.json"
    db_file: str = DATABASE_FILE_LOCAL.as_posix()
    db_schema: str = "ws"

    def get_filename(self):
        return (
            "#"
            + datetime.datetime.strftime(datetime.datetime.now(), "%Y%m%d")
            + self.output_file
        )

    def get_output_path(self):
        return self.output_dir / self.get_filename()


stockconf = FileSettings(
    DIR_OFFERS,
    DIR_OFFERS.parent / ".outputs",
    "_WS_Stock.xlsx",
    "Wholesale Stock",
    "Strats",
    6,
)

offersconf = FileSettings(
    DIR_OFFERS,
    DIR_OFFERS.parent / ".outputs",
    "_Coral_Homes_Offers_Data.xlsx",
    "Offers Data",
    "Strats",
    6,
)

pipeconf = FileSettings(
    DIR_PIPE,
    DIR_OFFERS.parent / ".outputs",
    "_WS_Pipeline.xlsx",
    "Pipeline 2023",
    "Strats",
    6,
)
