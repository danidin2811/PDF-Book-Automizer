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
    checklist = (
        "\nPRE-PROCESSING CHECKLIST:\n"
        "1. Close the Excel tracking table\n"
        "2. Ensure the numeric JPG cover is in the source folder\n"
        "3. Ensure the JPG filename matches the DanaCode\n\n"
    )
    print(prompt)
    input("Press Enter to continue: ")