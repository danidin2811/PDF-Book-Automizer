import re
import logging
from src.constants import SMALL_WORDS, VALID_TITLE_REGEX
from utils.input_output_tools import print_red

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')


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


def to_snake_case(title: str) -> str:
    """
    Converts an English string to web-safe snake_case for folder naming.
    """

    # Remove all non-alphanumeric (except spaces/hyphens for temporary splitting)
    text = re.sub(r'[^a-zA-Z0-9\s-]', '', title)
    text = re.sub(r'[\s-]+', '_', text)
    return text.strip('_').lower()


def get_book_metadata(raw_title: str) -> dict:
    """
    Orchestrates conversion. Includes validation for English alphanumeric content.
    """

    if not raw_title.strip():
        logging.info("Bypass triggered: Empty input provided.")
        return {"display_title": "", "folder_name": ""}

    if not is_valid_english_title(raw_title):
        logging.error("Invalid Input: Title must contain English alphanumeric characters.")
        return {"display_title": "Invalid Input", "folder_name": "invalid_input"}

    return {
        "display_title": to_title_case(raw_title),
        "folder_name": to_snake_case(raw_title)
    }


def normalize_book_title() -> dict | None:
    """
    CLI wrapper that persists until a valid English title is provided.
    Bypass (empty input) is no longer allowed.
    """
    while True:
        user_input = input("Enter book title in English: ").strip()

        # 1. Check for empty input immediately
        if not user_input:
            print_red("Error: Input cannot be empty. Please enter an English title.")
            continue

        # 2. Process and Validate via the metadata tool
        metadata = get_book_metadata(user_input)

        # 3. Check if the metadata tool actually liked the input
        if metadata["display_title"] not in ["Error", "Invalid Input"]:
            print("-" * 30)
            print(f"Display Title: {metadata['display_title']}")
            print(f"Folder Name:   {metadata['folder_name']}")
            print("-" * 30)
            return metadata  # Return the valid dictionary to the caller

        # 4. Fallback message if the characters were non-English or invalid
        print_red("Invalid input detected. Please use English alphanumeric characters.")


if __name__ == "__main__":
    normalize_book_title()