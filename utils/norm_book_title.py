import re
import logging
from src.constants import SMALL_WORDS, VALID_TITLE_REGEX
from utils.input_output_tools import print_red, yes_or_no


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


def to_snake_case(title: str) -> str:
    """
    Converts an English string to web-safe snake_case with user intervention.
    """

    def auto_convert(text):
        # The conversion logic
        text = re.sub(r'[^a-zA-Z0-9\s-]', '', text)
        text = re.sub(r'[\s-]+', '_', text)
        return text.strip('_').lower()

    snake_title = auto_convert(title)

    while True:
        # 1. Check Length
        if len(snake_title) > 40:
            print_red(f"[!] The generated snake title {snake_title} is too long ({len(snake_title)} chars). fliphtml limit is 40.")
            snake_title = input("Please enter a shorter version: ").strip()
            continue

        # 2. Check Format
        if not is_snake_case(snake_title):
            print_red(f"[!] '{snake_title}' is not valid snake_case.")

            # Offer to auto-fix the manually entered text
            if yes_or_no("Would you like to auto-format the title? "):
                snake_title = auto_convert(snake_title)
                print(f"Auto-formatted to: {snake_title}")
                continue  # Re-validate the newly formatted title

            else:
                snake_title = input("Please fix the format manually: ").strip()
                continue

        # 3. Final Confirmation
        break

    return snake_title


def get_book_metadata(raw_title: str) -> dict:
    """
    Orchestrates conversion. Asks user for title preference and always generates a snake_case folder name.
    """

    if not raw_title.strip():
        logging.info("Bypass triggered: Empty input provided.")
        return {"display_title": "", "folder_name": ""}

    if not is_valid_english_title(raw_title):
        logging.error("Invalid Input: Title must contain English alphanumeric characters.")
        return {"display_title": "Invalid Input", "folder_name": "invalid_input"}

    prompt = (f"You entered: '{raw_title}'.\n"
              f"Do you want to CONVERT this to Title Case? (y=convert, n=keep original) ")

    if yes_or_no(prompt):
        display_title = to_title_case(raw_title)

    else:
        display_title = raw_title

    # 2. Always normalize the folder name for the filesystem
    folder_name = to_snake_case(display_title)

    print("-" * 30)
    print(f"Display Title: {display_title}")
    print(f"Folder Name:   {folder_name}")
    print("-" * 30)

    return {
        "display_title": display_title,
        "folder_name": folder_name
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