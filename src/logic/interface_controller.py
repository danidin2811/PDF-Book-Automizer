# REMOVE the global top-level import: from utils.input_output_tools import yes_or_no

class AppInterface:
    """Manages abstraction layers between CLI (Terminal) and GUI (CustomTkinter)."""
    def __init__(self, ui=None):
        self.ui = ui  # If None, the system dynamically defaults to CLI mode
        self.is_gui = ui is not None

    @property
    def is_canceling(self) -> bool:
        """
        Dynamically forwards the cancellation status check.
        If running in GUI mode, tracks the actual window state.
        Otherwise, defaults safely to False for CLI loops.
        """
        if self.is_gui and hasattr(self.ui, 'is_canceling'):
            return self.ui.is_canceling
        return False

    def ask_string(self, title: str, prompt: str) -> str | None:
        """Abstracts picking up a string from either UI overlays or the terminal console."""
        if self.is_gui:
            return self.ui.async_ask_string(title, prompt)
        else:
            try:
                val = input(f"{prompt} ")
                return val.strip()
            except (KeyboardInterrupt, EOFError):
                return None

    def ask_yes_no(self, title: str, question: str) -> bool:
        """Abstracts true/false execution decisions across operational paradigms."""
        if self.is_gui:
            return self.ui.async_ask_yes_no(title, question)
        else:
            # Local runtime import completely breaks the circular dependency loop!
            from utils.input_output_tools import yes_or_no
            return yes_or_no(f"{question} (y/n): ")

    def print_error(self, message: str):
        """Routes error messages to the correct stream output."""
        if self.is_gui:
            self.ui.log(f"[WARN/ERROR] {message}")
        else:
            print(f"\033[91mError: {message}\033[0m")

    def print_success(self, message: str):
        """Routes success messages to the correct stream output."""
        if self.is_gui:
            self.ui.log(f"[SUCCESS] {message}")
        else:
            print(f"\033[32mSuccess: {message}\033[0m")

    def print_info(self, message: str):
        """Displays data structural confirmation results across layers."""
        if self.is_gui:
            self.ui.log(message)
        else:
            print(message)

    def print_header(self, title: str):
        """Dedicated visual formatting wrapper for critical system checkpoints."""
        divider = "-" * 40
        if self.is_gui:
            self.ui.log(divider)
            self.ui.log(title.strip("\n"))
            self.ui.log(divider)
        else:
            print(divider)
            print(title)
            print(divider)

    def ask_checkpoint(self, title: str, action_message: str):
        """
        Pauses execution until the user manually completes an external task
        (e.g., renaming a folder, checking Adobe Acrobat, saving a CSV file).
        """
        if self.is_gui:
            return self.ui.async_blocking_checkpoint(title, action_message)
        else:
            print(f"\n[ACTION REQUIRED] {title}")
            print("-" * (len(title) + 18))
            print(action_message)
            input("\nPress Enter once you have completed this step to continue... ")
            print("Proceeding...")

    def request_all_page_ranges(self, sections: list, total_pages: int) -> dict:
        """
        Gathers valid page ranges for all specified sections at once.
        Returns a dictionary mapping section name -> (start, end).
        """
        if self.is_gui and hasattr(self.ui, 'get_all_ranges_from_form'):
            return self.ui.get_all_ranges_from_form(sections, total_pages)

        return get_all_page_ranges_cli(sections, total_pages, self)


def get_all_page_ranges_cli(sections: list, total_pages: int, interface: AppInterface) -> dict:
    """
    CLI-specific loop that prompts for all section ranges sequentially,
    validates them against total pages, and returns the full map once submitted.
    """
    while True:
        ranges = {}
        cancelled = False
        interface.print_header(f"Enter Page Ranges (Total Book Pages: {total_pages})")

        for section in sections:
            if interface.is_canceling:
                return {}

            interface.print_info(f"\n[ Section: {section.upper()} ]")
            while True:
                try:
                    start_str = interface.ask_string("Range Input", f"Enter start page for {section.upper()}:")
                    if start_str is None or start_str == "":
                        cancelled = True
                        break

                    end_str = interface.ask_string("Range Input", f"Enter end page for {section.upper()}:")
                    if end_str is None or end_str == "":
                        cancelled = True
                        break

                    start = int(start_str)
                    end = int(end_str)

                    if 1 <= start <= end <= total_pages:
                        ranges[section] = (start, end)
                        break

                    interface.print_error(f"Invalid range! Must be between 1 and {total_pages}, and start <= end.")
                except ValueError:
                    interface.print_error("Please enter numbers only.")

            if cancelled:
                break

        if cancelled or interface.is_canceling:
            return {}

        interface.print_header("Review Your Ranges")
        for sec, (s, e) in ranges.items():
            interface.print_info(f"  {sec.upper()}: Pages {s} to {e}")

        if interface.ask_yes_no("Confirm Ranges", "Submit all ranges?"):
            return ranges

        interface.print_info("Let's re-enter the data.\n")