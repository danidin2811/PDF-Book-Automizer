import re
import logging
from src.constants import SMALL_WORDS, VALID_TITLE_REGEX
from utils.input_output_tools import print_red, yes_or_no
from src.logic.interface_controller import AppInterface

def is_valid_english_title(title: str) -> bool:
    """
    Validates that the title contains English alphanumeric characters.
    Allows specific punctuation: - ' " , . ? !
    """

    # Check if empty or only whitespace
    if not title.strip():
        return True  # Handled as a bypass elsewhere

    # Check against allowed characters
    if not re.match(VALID_TITLE_REGEX, title):
        return False

    # Ensure there is at least one letter or number (not just punctuation)
    return any(char.isalnum() for char in title)


def to_title_case(title: str) -> str:
    """
    Converts an English string to Title Case while keeping prepositions lowercase.
    """

    clean_title = title.replace('-', ' ').strip()
    words = clean_title.split()

    formatted_words = [
        word.capitalize() if word.lower() not in SMALL_WORDS or i == 0
        else word.lower()
        for i, word in enumerate(words)
    ]
    return ' '.join(formatted_words)


def is_snake_case(text: str) -> bool:
    """
    Checks if a string consists only of lowercase letters, numbers, and underscores.
    """
    # Regex: ^ (start), [a-z0-9_] (allowed chars), + (one or more), $ (end)
    pattern = r'^[a-z0-9_]+$'
    return bool(re.match(pattern, text))


def to_snake_case(title: str, interface: AppInterface) -> str | None:
    """
    Converts an English string to web-safe snake_case with user intervention.
    Now fully abstracts input and decisions through the interface layer.
    """

    def auto_convert(text):
        text = re.sub(r'[^a-zA-Z0-9\s-]', '', text)
        text = re.sub(r'[\s-]+', '_', text)
        return text.strip('_').lower()

    snake_title = auto_convert(title)

    while True:
        # Check for user cancellation in GUI mode
        if interface.is_gui and interface.ui.is_canceling:
            return None

        # 1. Check Length Optimization
        if len(snake_title) > 40:
            interface.print_error(
                f"The generated snake title '{snake_title}' is too long ({len(snake_title)} chars). FlipHTML5 limit is 40.")

            snake_title = interface.ask_string(
                "Title Length Recovery",
                f"Generated title too long ({len(snake_title)} chars). Please enter a shorter version:"
            )
            if snake_title is None:
                return None
            snake_title = snake_title.strip()
            continue

        # 2. Check Structural Validity Format
        if not is_snake_case(snake_title):
            interface.print_error(f"'{snake_title}' is not valid snake_case.")

            # Abstract choice handling using the updated Interface pattern
            if interface.ask_yes_no("Auto-Format Prompt",
                                    "Would you like to auto-format the invalid layout parameters?"):
                snake_title = auto_convert(snake_title)
                continue
            else:
                snake_title = interface.ask_string("Manual Correction", "Please fix the format manually:")
                if snake_title is None:
                    return None
                snake_title = snake_title.strip()
                continue

        # 3. Final Confirmation Reached Successfully
        break

    return snake_title


def get_book_metadata(raw_title: str, interface: AppInterface) -> dict:
    """Orchestrates conversion. Asks user for title preference and generates folder metadata."""
    if not raw_title.strip():
        logging.info("Bypass triggered: Empty input provided.")
        return {"display_title": "", "folder_name": ""}

    if not is_valid_english_title(raw_title):
        return {"display_title": "Invalid Input", "folder_name": "invalid_input"}

    display_title = to_title_case(raw_title)

    # Forward the interface controller layer down into the input loops
    folder_name = to_snake_case(display_title, interface)

    return {
        "display_title": display_title,
        "folder_name": folder_name
    }


def normalize_book_title(interface: AppInterface) -> dict | None:
    """
    Dual-mode wrapper that persists until a valid English title is provided.
    Works dynamically over CLI streams or Asynchronous GUI overlay panels.
    """
    while True:
        # 1. Ask for input abstractly through the controller
        user_input = interface.ask_string("Book Naming", "Enter book title in English:")

        # Check for user cancellation or closing the window frame in GUI mode
        if interface.is_gui and interface.ui.is_canceling:
            return None
        if user_input is None:
            return None

        user_input = user_input.strip()

        # 2. Check for empty input immediately
        if not user_input:
            interface.print_error("Input cannot be empty. Please enter an English title.")
            continue

        # 3. Process and Validate via the metadata tool
        metadata = get_book_metadata(user_input, interface)

        # 4. Check if the metadata tool approved the input configuration matrix
        if metadata["display_title"] not in ["Error", "Invalid Input"]:
            interface.print_info(f"Display title = {metadata['display_title']}\nFolder name = {metadata['folder_name']}")
            return metadata  # Return the valid dictionary to the caller

        # 5. Fallback message if the characters were non-English or invalid
        interface.print_error("Invalid input detected. Please use English alphanumeric characters.")


if __name__ == "__main__":
    normalize_book_title()