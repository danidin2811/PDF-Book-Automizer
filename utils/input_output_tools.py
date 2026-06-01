from src.logic.interface_controller import AppInterface


def print_red(text):
    """Prints text in red color."""
    print("\033[31m" + text + "\033[0m")


def print_green(text):
    """Prints text in green color."""
    print("\033[32m\n" + text + "\033[0m")


def yes_or_no(prompt):
    while True:
        choice = input(prompt).strip().lower()
        if choice in {"y", "yes"}:
            return True
        elif choice in {"n", "no"}:
            return False
        print_red("Invalid input. Please enter 'y' for yes or 'n' for no.")


def wait_for_ready_signal(prompt):
    """Confirms system requirements are met before starting."""
    print(prompt)
    input("Press Enter to continue: ")
    print("Enter pressed")


def ask_offset(interface: AppInterface):
    """
    Prompts the user to specify an offset for book page numbering.
    Returns:
        int: The offset entered by the user. Always an integer value (can be positive, negative, or zero).
    """
    if interface.is_gui and hasattr(interface.ui, 'get_offset_value'):
        try:
            return int(interface.ui.get_offset_value())
        except (ValueError, TypeError):
            pass

    while True:
        try:
            offset_str = interface.ask_string(
                "Page Offset Check",
                "Is there an offset in the book's pages?\n"
                "Please enter the amount of offset pages as a number (positive, negative, or 0 for none):"
            )
            if offset_str is None or offset_str == "":
                return 0

            offset = int(offset_str.strip())
            interface.print_info(f"You entered offset: {offset}")

            if interface.ask_yes_no("Confirm Offset", "Is this correct?"):
                return offset

            interface.print_info("Re-enter the offset.")
        except ValueError:
            interface.print_error("\nInvalid input. Please enter a valid integer.\n")