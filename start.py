import subprocess
import sys

def display_menu(input_dict):
    dict_title = input_dict.get("title", "Choose an updater")
    message = "Type option and press enter. You can type 'Q' to quit the menu."
    separator_up = f"\n{dict_title:=^40}\n"
    separator_bot = f"\n{message:=^40}\n"
    menu = (
        separator_up
        + "\n".join(
            f"[{k}]\t{v[0]}" for k, v in input_dict.get("menu").items()
        )
        + separator_bot
        + " >>> "
    )
    choice = input(menu)

    if choice.lower() == "q":
        return 0

    try:
        return int(choice)
    except ValueError as e:
        print(f"Exception: {e} happened. Invalid choice.")
        return "Error"

def run(input_dict):
    choice = display_menu(input_dict)
    # Display menu and respond to user choices
    match choice:
        case 0:
            print("Goodbye! 👋")
            sys.exit()
        case "Error":
            print("Invalid choice. Please try again.")
            return run(input_dict)
        case _:
            print(f"Chosen option: {choice}")
            result = input_dict.get("menu").get(choice)

    # Treat function(s) and kwargs as iterables
    functions = result[1] if isinstance(result[1], list) else [result[1]]
    kwargs_list = result[2] if isinstance(result[2], list) else [result[2]]
    # Loop over functions and kwargs together
    for function, kwargs in zip(functions, kwargs_list):
        if callable(function):
            if kwargs:
                function(**kwargs)
            else:
                function()
        else:
            raise TypeError(f"The item {function.__name__} is not callable")

    return run(input_dict)


def blank_runner(script: str, other_params: list = None):
    if other_params:
        return subprocess.run([sys.executable, script, *other_params])
    else:
        return subprocess.run([sys.executable, script])

# Define the menus

ws_menu = {
    "title": "Coral Homes Wholesale Auto Ops",
    "menu": {
        1: ("Obtener ofertas desde el email", blank_runner, {"script": "./retrieve_email_attachments.py", "other_params": ["--file_type", "offers", "--months", "5"] }),
        2: ("Obtener pipeline desde el email", blank_runner, {"script": "./retrieve_email_attachments.py", "other_params": ["--file_type", "pipe", "--months", "5"] }),
        3: ("Crear / Actualizar fichero ofertas", blank_runner, {"script": "./update_offers.py", "other_params": ["--update", "True", "--current", "True"] }),
        4: ("Crear / Actualizar fichero pipeline", blank_runner, {"script": "./update_pipe.py" }),
        5: ("Crear / Actualizar fichero stock", blank_runner, {"script": "./update_stock.py" }),
    },
}


# Run the menu
if __name__ == "__main__":
    run(ws_menu)
