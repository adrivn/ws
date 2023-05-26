from datetime import datetime
from pathlib import Path
from rich.console import Console
import argparse
import re
import win32com.client

console = Console()

def retrieve_attachments(subject_pattern, filename_pattern, extensions, top_level_dir):
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
                    sub_dir = received_time.strftime("%Y/%m/%d")

                    if top_level_dir.startswith("./"):
                        # Relative directory
                        full_dir = Path.cwd() / top_level_dir[2:] / sub_dir
                    else:
                        # Absolute directory
                        full_dir = Path(top_level_dir) / sub_dir

                    # Create the directory if it doesn't exist
                    full_dir.mkdir(parents=True, exist_ok=True)

                    # Save the attachment in the directory
                    console.print(
                        f"\tAttempting to save down to {str(full_dir / attachment.FileName)}"
                    )
                    attachment.SaveASFile(str(full_dir / attachment.FileName))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--type", required=True, help="Type of file to search: pipeline or offers"
    )
    parser.add_argument(
        "--path",
        required=True,
        help="Relative or absolute path to save down the files to.",
    )
    args = parser.parse_args()

    match args.type:
        case "pipe":
            retrieve_attachments(
                "\d{6,8} Pipe 20",
                "^\d{6,8} Pipe 20",
                [".xlsx", ".xls", ".xlsb"],
                args.path,
            )
        case "offers":
            retrieve_attachments(
                "OF CH ", "$\d{6,8}_OF_", [".xlsx", ".xls", ".xlsb"], args.path
            )
        case _:
            raise ValueError(
                f"You must specify either 'pipe' or 'offers' in the type, because {args.type} is not allowed."
            )
