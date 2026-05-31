class AppInterface:
    """Manages abstraction layers between CLI (Terminal) and GUI (CustomTkinter)."""
    def __init__(self, ui=None):
        self.ui = ui  # If None, the system dynamically defaults to CLI mode
        self.is_gui = ui is not None

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