from utils.input_output_tools import get_all_page_ranges_cli


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
            return input(f"{prompt} ").strip()

    def ask_yes_no(self, title: str, question: str) -> bool:
        """Abstracts true/false execution decisions across operational paradigms."""
        if self.is_gui:
            return self.ui.async_ask_yes_no(title, question)
        else:
            # Use your pre-existing terminal validation utility
            from utils.input_output_tools import yes_or_no
            return yes_or_no(f"{question} (y/n): ")

    def print_error(self, message: str):
        """Routes error messages to the correct stream output."""
        if self.is_gui:
            self.ui.log(f"[WARN/ERROR] {message}")
        else:
            print(f"\033[91mError: {message}\033[0m")

    def print_info(self, title: str, folder_name: str):
        """Displays data structural confirmation results across layers."""
        message = f"Display Title: {title}\nFolder Name:   {folder_name}"
        if self.is_gui:
            self.ui.log("-" * 30)
            self.ui.log(message)
            self.ui.log("-" * 30)
        else:
            print("-" * 30)
            print(message)
            print("-" * 30)

    def ask_checkpoint(self, title: str, action_message: str):
        """
        Pauses execution until the user manually completes an external task
        (e.g., renaming a folder, checking Adobe Acrobat, saving a CSV file).
        """
        if self.is_gui:
            # Invokes the thread-safe overlay banner with a confirmation button
            return self.ui.async_blocking_checkpoint(title, action_message)

        else:
            # Fallback to the standard blocking terminal prompt
            print(f"\n[ACTION REQUIRED] {title}")
            print("-" * (ffff := len(title) + 18))
            print(action_message)
            input("\nPress Enter once you have completed this step to continue... ")
            print("Proceeding...")

    def request_all_page_ranges(self, sections: list, total_pages: int) -> dict:
        """
        Gathers valid page ranges for all specified sections at once.
        Returns a dictionary mapping section name -> (start, end).
        """
        if self.is_gui and hasattr(self.ui, 'get_all_ranges_from_form'):
            # The GUI can display a single window containing fields for all sections
            return self.ui.get_all_ranges_from_form(sections, total_pages)

        # Fallback for CLI loop
        return get_all_page_ranges_cli(sections, total_pages, self)