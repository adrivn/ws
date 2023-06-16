from datetime import datetime
from pathlib import Path
from rich.console import Console
import argparse
import re
import win32com.client

console = Console()

def retrieve_attachments(subject_pattern: str, filename_pattern: str, extensions: list, file_type: str, top_level_dir: str):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # "6" refers to the inbox

    messages = inbox.Items

    start_count = 0
    for message in messages:
        if re.search(subject_pattern.lower(), message.Subject.lower()):
            start_count += 1
            received_time = message.ReceivedTime
            console.print(
                f"[{start_count}] Found email:",
                message.Subject,
                "received on",
                received_time,
            )
            attachments = message.Attachments
            for attachment in attachments:
                if re.search(
                    filename_pattern.lower(), attachment.FileName.lower()
                ) and any(attachment.FileName.endswith(ext) for ext in extensions):
                    console.print("\tAttachment found:", attachment.FileName, style="bold red")

                    # Get the email's received time and format it as a directory path
                    year_folder = received_time.strftime("%Y")
                    day_folder = received_time.strftime("%Y%m%d")

                    match top_level_dir, file_type:
                        case "coralhudson", "pipe":
                            base_dir = Path("//EURFL01/advisors.hal/non-hudson/Coral Homes/CoralHudson/1. AM/8. Wholesale Channel/")
                            full_dir = base_dir / f"WS PIPE {year_folder}"
                        case "currentdir", "pipe":
                            base_dir = Path("./_attachments/pipe_files/") 
                            full_dir = base_dir / f"WS PIPE {year_folder}"
                        case "coralhudson", "offers":
                            base_dir = Path("//EURFL01/advisors.hal/non-hudson/Coral Homes/CoralHudson/1. AM/8. Wholesale Channel/Ofertas recibidas SVH/")
                            full_dir = base_dir / year_folder / day_folder
                        case "currentdir", "offers":
                            base_dir = Path("./_attachments/offer_files/") 
                            full_dir = base_dir / year_folder / day_folder

                    # Create the directory if it doesn't exist
                    full_dir.mkdir(parents=True, exist_ok=True)

                    # Save the attachment in the directory
                    console.print(
                        f"\tAttempting to save down to {(full_dir / attachment.FileName).absolute()}"
                    )
                    attachment.SaveASFile(str((full_dir / attachment.FileName).absolute()))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--file_type", 
        required=True,
        choices=["pipe", "offers"],
        help="Type of file to search: pipeline or offers"
    )
    parser.add_argument(
        "--path",
        required=True,
        default="coralhudson",
        choices=["coralhudson", "currentdir"],
        help="Where to save down the files to.",
    )
    args = parser.parse_args()

    match args.file_type:
        case "pipe":
            retrieve_attachments(
                "\d{6,8} Pipe 20",
                "^\d{6,8} Pipe 20",
                [".xlsx", ".xls", ".xlsb", ".xlsm"],
                args.file_type,
                args.path
            )
        case "offers":
            retrieve_attachments(
                "OF CH ",
                "^\d{6,8}[_ ]+OF_",
                [".xlsx", ".xls", ".xlsb", ".xlsm"],
                args.file_type,
                args.path
            )
        case _:
            raise ValueError(
                f"You must specify either 'pipe' or 'offers' in the type, because {args.type} is not allowed."
            )
