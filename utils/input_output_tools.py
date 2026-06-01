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


def get_all_page_ranges_cli(sections: list, total_pages: int, interface) -> dict:
    """
    CLI-specific loop that prompts for all section ranges sequentially,
    validates them against total pages, and returns the full map once submitted.
    """
    while True:
        ranges = {}
        cancelled = False
        interface.print_info(f"\n--- Enter Page Ranges (Total Book Pages: {total_pages}) ---")

        for section in sections:
            if interface.is_canceling:
                return {}

            interface.print_info(f"\n[ Section: {section.upper()} ]")
            while True:
                try:
                    start_str = interface.input_prompt(f"Enter start page for {section.upper()}: ")
                    if start_str is None:
                        cancelled = True;
                        break

                    end_str = interface.input_prompt(f"Enter end page for {section.upper()}: ")
                    if end_str is None:
                        cancelled = True;
                        break

                    start = int(start_str.strip())
                    end = int(end_str.strip())

                    # Your exact validation logic
                    if 1 <= start <= end <= total_pages:
                        ranges[section] = (start, end)
                        break  # Move to next section

                    interface.print_error(f"Invalid range! Must be between 1 and {total_pages}, and start <= end.")
                except ValueError:
                    interface.print_error("Please enter numbers only.")

            if cancelled:
                break

        if cancelled or interface.is_canceling:
            return {}

        # Summary and final unified submission step
        interface.print_info("\n--- Review Your Ranges ---")
        for sec, (s, e) in ranges.items():
            interface.print_info(f"  {sec.upper()}: Pages {s} to {e}")

        confirm = interface.input_prompt("\nSubmit all ranges? (y/n): ")
        if confirm and confirm.strip().lower() == 'y':
            return ranges

        interface.print_info("Let's re-enter the data.")


def wait_for_ready_signal(prompt):
    """Confirms system requirements are met before starting."""

    print(prompt)
    input("Press Enter to continue: ")
    print("Enter pressed")


def ask_offset(interface):
    """
    Prompts the user to specify an offset for book page numbering.
    Returns:
        int: The offset entered by the user. Always an integer value (can be positive, negative, or zero).
    """
    # Check if the GUI layer can supply this value directly
    if interface.is_gui and hasattr(interface.ui, 'get_offset_value'):
        try:
            return int(interface.ui.get_offset_value())
        except (ValueError, TypeError):
            # If the GUI has a bad value or it's unconfigured, fall back to safe zero or terminal
            pass

    while True:
        try:
            offset_str = interface.input_prompt(
                "\nIs there an offset in the book's pages?\n"
                "Please enter the amount of offset pages as a number (positive, negative, or 0 for none): "
            )
            if offset_str is None:  # Handle cancellation if supported
                return 0

            offset = int(offset_str.strip())
            interface.print_info(f"You entered offset: {offset}")

            # Using interface's built-in confirmation wrapper if available, else standard terminal
            if hasattr(interface, 'confirm_choice'):
                if interface.confirm_choice("Is this correct?"):
                    return offset
            else:
                confirm = input("Is this correct? (y/n): ").strip().lower()
                if confirm == 'y':
                    return offset
                interface.print_info("Re-enter the offset.")
        except ValueError:
            interface.print_error("\nInvalid input. Please enter a valid integer.\n")