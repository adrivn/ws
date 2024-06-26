import datetime
import subprocess
import sys

from conf.functions import timing


def display_menu(input_dict):
    dict_title = input_dict.get("title", "Choose an updater")
    message = "Type option and press enter. You can type 'Q' to quit the menu."
    separator_up = f"\n{dict_title:=^40}\n"
    separator_bot = f"\n{message:=^40}\n"
    menu = (
        separator_up
        + "\n".join(f"[{k}]\t{v[0]}" for k, v in input_dict.items())
        + separator_bot
        + " >>> "
    )
    choice = input(menu)

    if choice.lower() == "q":
        print("Goodbye! 👋")
        sys.exit()

    if choice.isdigit() and int(choice) in input_dict:
        selected_option = input_dict[int(choice)]
        option_text, rest = selected_option[0], selected_option[1:]

        if isinstance(rest[0], dict):
            submenu = rest[0]
            print(f"\n{option_text} Submenu:")
            display_menu(submenu)
        elif callable(rest[0]):
            function_to_run, kwargs_dict = rest
            if kwargs_dict:
                function_to_run(**kwargs_dict)
            else:
                function_to_run()

            display_menu(ws_menu)
        else:
            print("Invalid option, please try again.")
    else:
        print("Invalid input, please enter a valid option.")


@timing
def blank_runner(script: str, other_params: list = None):
    if other_params:
        return subprocess.run([sys.executable, script, *other_params])
    else:
        return subprocess.run([sys.executable, script])


# Define the menus

current_year = datetime.datetime.now().year

ws_menu = {
    1: (
        "Obtener ofertas desde el email",
        blank_runner,
        {
            "script": "./retrieve_email_attachments.py",
            "other_params": ["--file_type", "offers", "--months", "1"],
        },
    ),
    2: (
        "Obtener pipeline desde el email",
        blank_runner,
        {
            "script": "./retrieve_email_attachments.py",
            "other_params": ["--file_type", "pipe", "--months", "1"],
        },
    ),
    3: (
        "Crear / Actualizar fichero ofertas",
        {
            1: (
                f"Escanear todas las ofertas de {current_year} y crear fichero",
                blank_runner,
                {
                    "script": "./update_offers.py",
                    "other_params": [
                        "--update",
                        "--write",
                        f"--year={current_year}",
                    ],
                },
            ),
            2: (
                f"Escanear sólo las nuevas ofertas de {current_year} y crear fichero",
                blank_runner,
                {
                    "script": "./update_offers.py",
                    "other_params": [
                        "--update",
                        "--refresh",
                        "--write",
                        f"--year={current_year}",
                    ],
                },
            ),
            3: (
                "Crear nuevo fichero",
                blank_runner,
                {"script": "./update_offers.py", "other_params": ["--write"]},
            ),
        },
    ),
    4: (
        "Crear / Actualizar fichero pipeline",
        blank_runner,
        {"script": "./update_pipe.py"},
    ),
    5: (
        "Crear / Actualizar fichero stock",
        blank_runner,
        {"script": "./update_stock.py"},
    ),
}


# Run the menu
if __name__ == "__main__":
    display_menu(ws_menu)
