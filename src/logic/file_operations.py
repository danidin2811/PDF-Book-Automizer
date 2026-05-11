import os
import re
import shutil
from pathlib import Path

from utils.input_output_tools import print_red, print_green, wait_for_ready_signal


def validate_csv_path(path_str: str) -> tuple[bool, str]:
    """Validates if a string is a valid path to an existing CSV."""

    # Strip both types of quotes that might come from drag-and-drop
    clean_path = path_str.strip('"').strip("'")
    path = Path(clean_path)

    if not path.exists():
        return False, f"CSV file path {path} does not exist."
    if path.suffix.lower() != '.csv':
        return False, f"File {path} is not a CSV."

    return True, str(path)


def validate_pdf_path(path_str: str) -> tuple[bool, str]:
    """Validates if a string is a valid path to an existing PDF."""

    path = Path(path_str.strip('"'))

    if not path.exists():
        return False, f"File path {path} does not exist."
    if not path.is_file():
        return False, f"Path {path} is not a file."
    if path.suffix.lower() != '.pdf':
        return False, f"File {path} is not a PDF."

    return True, ""


def move_cover_image(source_dir: Path, dest_dir: Path) -> str:
    """
    Moves a numeric JPG (DanaCode) from source to destination.

    Returns the DanaCode string if found, otherwise returns an empty string.
    """

    # 1. Validation for network drive reliability
    if not source_dir.exists():
        print_red(f"Source directory does not exist: {source_dir}")
        return ''

    dest_dir.mkdir(parents=True, exist_ok=True)

    # 2. Pattern for 3+ digits (DanaCode standard)
    dana_pattern = re.compile(r"^(\d{3,})\.jpg$")

    # 3. Iterating with corrected parentheses
    for file_path in source_dir.iterdir():
        if file_path.is_file():  # Pro tip: ensure it's a file, not a sub-folder
            match = dana_pattern.match(file_path.name)
            if match:
                dana_code = match.group(1)
                try:
                    shutil.move(str(file_path), str(dest_dir / file_path.name))
                    print(f"Moved DanaCode: {dana_code}")
                    return dana_code
                except PermissionError:
                    print_red(f"File {file_path.name} is in use.")

    return ''


def check_file_size(file_path: str, limit_mb: int = 500) -> bool:
    """
    Checks if a PDF exceeds the limit and advises on the new naming convention.
    """

    if not os.path.exists(file_path):
        return False

    limit_bytes = limit_mb * 1024 * 1024
    file_size_bytes = os.path.getsize(file_path)
    file_size_mb = file_size_bytes / (1024 * 1024)

    if file_size_bytes > limit_bytes:
        print(f"\n[!] ALERT: File is {file_size_mb:.2f} MB (Limit: {limit_mb} MB).")

        # Construct the suggested new filename
        original_path = Path(file_path)
        new_name = f"DOWNSIZED {original_path.name}"

        wait_for_ready_signal(
            "ACTION REQUIRED: The PDF is too large for FlipHTML5.\n"
            "1. Open the PDF in Acrobat\n"
            "2. Use 'Save as Other' -> 'Reduced Size PDF'\n"
            f"3. Save the new file as save: {new_name}\n"
        )

        wait_for_ready_signal(
            "ACTION REQUIRED:\n"
            f"Before proceeding, please make sure that the new DOWNSIZED file is less than {limit_mb} MB\n"
        )


    print_green(f"Check passed: {file_size_mb:.2f} MB.")
    return True