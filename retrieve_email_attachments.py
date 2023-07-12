from pathlib import Path
from rich.console import Console
import datetime
import argparse
import re
import glob
import win32com.client

console = Console()

def retrieve_attachments(file_type: str, top_level_dir: str, number_of_months: int = 1):
    # Conseguir todos los ficheros del directorio
    # Listar todos los ficheros del directorio de destino, y filtrar contra los encontrados
    console.print(f"Checking existing {file_type} files...")
    match top_level_dir, file_type:
        case "coralhudson", "pipe":
            base_dir = Path("N:/CoralHudson/1. AM/8. Wholesale Channel/")
            ficheros_existentes = [Path(f).name for f in glob.glob("N:/CoralHudson/1. AM/8. Wholesale Channel/WS PIPE*/[!~$]*.*")]
            subject_pattern = r"\d{6,8} Pipe 20"
            filename_pattern = r"^\d{6,8} Pipe 20"
        case "currentdir", "pipe":
            base_dir = Path("./_attachments/pipe_files/") 
            ficheros_existentes = [f.name for f in base_dir.rglob("*.*")]
            subject_pattern = r"\d{6,8} Pipe 20"
            filename_pattern = r"^\d{6,8} Pipe 20"
        case "coralhudson", "offers":
            base_dir = Path("N:/CoralHudson/1. AM/8. Wholesale Channel/Ofertas recibidas SVH/")
            ficheros_existentes = [Path(f).name for f in glob.glob(base_dir.as_posix() + "/2023/**/[!~$]*.*")]
            subject_pattern = r"OF CH "
            filename_pattern = r"^\d{6,8}[_ ]+OF_"
        case "currentdir", "offers":
            base_dir = Path("./_attachments/offer_files/") 
            ficheros_existentes = [f.name for f in base_dir.rglob("*.*")]
            subject_pattern = r"OF CH "
            filename_pattern = r"^\d{6,8}[_ ]+OF_"

    # Filtrado por fecha
    time_range = datetime.date.today() - datetime.timedelta(days = 30 * int(number_of_months))
    time_string = time_range.strftime("%m/%d/%Y")
    console.print(f"Buscando mensajes recibidos desdes el {time_string}")
    extensions = [".xlsx", ".xls", ".xlsb", ".xlsm"]

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # "6" refers to the inbox
    sFilter = f"[ReceivedTime] >= '{time_string}'"
    messages = inbox.Items.Restrict(sFilter)

    if messages.count == 0:
        print("No messages found.")
        return

    message_count = 0
    duplicados = 0
    guardados = 0
    for message in messages:
        if re.search(subject_pattern.lower(), message.Subject.lower()):
            message_count += 1
            received_time = message.ReceivedTime
            console.print(
                f"[{message_count}] Found email:",
                message.Subject,
                "received on",
                received_time,
            )
            attachments = message.Attachments
            for attachment in attachments:
                if re.search(
                    filename_pattern.lower(), attachment.FileName.lower()
                ) and any(attachment.FileName.endswith(ext) for ext in extensions):
                    # Filtrar aquellos que ya estan encontrados
                    if attachment.FileName in ficheros_existentes:
                        console.print("\tAttachment ignored. Already in folder.")
                        duplicados += 1
                    else:
                        console.print("\tAttachment found:", attachment.FileName, style="bold red")
                        guardados += 1
                        # Get the email's received time and format it as a directory path
                        year_folder = received_time.strftime("%Y")
                        day_folder = received_time.strftime("%Y%m%d")

                        match file_type:
                            case "pipe":
                                full_dir = base_dir / f"WS PIPE {year_folder}"
                            case "offers":
                                full_dir = base_dir / year_folder / day_folder


                        # Create the directory if it doesn't exist
                        full_dir.mkdir(parents=True, exist_ok=True)

                        # Save the attachment in the directory
                        console.print(
                            f"\tAttempting to save down to {(full_dir / attachment.FileName).absolute()}"
                        )
                        attachment.SaveASFile(str((full_dir / attachment.FileName).absolute()))
    console.print(f"Guardados {guardados} ficheros en el directorio base: {base_dir}.")
    console.print(f"Omitidos {duplicados} ficheros por ya encontrarse en la carpeta.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--months",
        type=int,
        default=1,
        choices=range(1,12),
        help="Number of months to look back to (3 months back, 5 months back). Default: 1"
    )
    parser.add_argument(
        "--file_type", 
        required=True,
        choices=["pipe", "offers"],
        help="Type of file to search: pipeline or offers"
    )
    parser.add_argument(
        "--path",
        default="coralhudson",
        choices=["coralhudson", "currentdir"],
        help="Where to save down the files to.",
    )
    args = parser.parse_args()

    retrieve_attachments(
        args.file_type,
        args.path,
        args.months
    )
