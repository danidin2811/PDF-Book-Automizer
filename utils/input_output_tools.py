def print_red(text):
    """Prints text in red color."""
    print("\033[31m" + text + "\033[0m")

def print_green(text):
    """Prints text in green color."""
    print("\033[32m" + '\n' + text + "\033[0m")

def yes_or_no(prompt):
    while True:
        choice = input(prompt).strip().lower()
        if choice in {"y", "yes"}:
            return True
        elif choice in {"n", "no"}:
            return False
        print_red("Invalid input. Please enter 'y' for yes or 'n' for no.")


def get_page_range_ui(section: str, total_pages: int):
    """UI function to get ranges from user."""
    while True:
        try:
            start = int(input(f"Enter start page for {section.upper()}: "))
            end = int(input(f"Enter end page for {section.upper()}: "))

            if 1 <= start <= end <= total_pages:
                return start, end

            print_red(f"Invalid range. Total pages: {total_pages}")

        except ValueError:
            print_red("Please enter numbers only.")


def wait_for_ready_signal(prompt):
    """Confirms system requirements are met before starting."""

    print(prompt)
    input("Press Enter to continue: ")


def ask_offset():
    """
    Prompts the user to specify an offset for book page numbering.
    Returns:
        int: The offset entered by the user. Always an integer value (can be positive, negative, or zero).
    """
    while True:
        try:
            offset = int(input(
                "\nIs there an offset in the book's pages?\n"
                "Please enter the amount of offset pages as a number (positive, negative, or 0 for none): "
            ).strip())

            print(f"You entered offset: {offset}")
            confirm = input("Is this correct? (y/n): ").strip().lower()
            if confirm == 'y':
                return offset
            else:
                print("Re-enter the offset.")
        except ValueError:
            print_red("\nInvalid input. Please enter a valid integer.\n")