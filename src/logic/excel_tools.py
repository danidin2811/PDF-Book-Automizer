import os
import subprocess
import logging
from pathlib import Path
import psutil
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from src.constants import BOOK_TRACKER_EXCEL_FILE_PATH
from src.logic.file_operations import validate_csv_path
from utils.input_output_tools import print_red
import csv


logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')


def get_new_toc_entries(csv_path):
    """Extracts TOC data with detailed error logging."""
    new_entries = []

    if not os.path.exists(csv_path):
        print(f"DEBUG: File not found at {csv_path}")
        return []

    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        # Use DictReader to handle header variations
        reader = csv.DictReader(f)

        for line_num, row in enumerate(reader, start=2):  # Header is line 1
            try:
                # 1. Level Check
                level_raw = row.get("level", "").strip()
                if not level_raw:
                    raise ValueError("Missing 'level' field")
                level = int(level_raw)

                # 2. Title Check
                title = row.get("title", "").strip()
                if not title:
                    raise ValueError("Missing 'title' field")

                # 3. Page Number Check (Handling both 'page_number' and 'page number')
                page_val = row.get("page_number") or row.get("page number")

                # Allow empty page numbers for level 1 (sections/parts)
                if not page_val or not page_val.strip():
                    if level == 1:
                        page = 0
                    else:
                        raise ValueError(f"Missing page number for level {level} entry")
                else:
                    # Clean the page string in case there are non-digit characters
                    clean_page = ''.join(c for c in page_val if c.isdigit())
                    if not clean_page:
                        raise ValueError(f"Invalid page format: '{page_val}'")
                    page = int(clean_page)

                new_entries.append({
                    "level": level,
                    "title": title,
                    "page": page
                })

            except ValueError as ve:
                print(f"Row {line_num} Data Error: {ve} | Content: {row}")
            except Exception as e:
                print(f"Row {line_num} Unexpected Error: {e}")

    return new_entries


def process_toc_extraction(initial_csv_path):
    current_path = initial_csv_path

    while True:
        is_valid, result_or_error = validate_csv_path(current_path)

        if not is_valid:
            print(f"\n[!] Path Error: {result_or_error}")
            current_path = input("Please enter the correct path to the .csv file: ")
            continue  # Go back to start of loop to check the new path

        # Step 2: Try to extract entries
        # result_or_error is now the cleaned path string
        entries = get_new_toc_entries(result_or_error)

        if entries:
            print("-" * 30)
            print(f"✅ Success! Loaded {len(entries)} entries.")
            print("-" * 30)
            return entries

        # Step 3: Handle empty/invalid file content
        print("\n" + "!" * 40)
        print("[!] CRITICAL: No valid entries found in the file.")
        print(f"File Path: {result_or_error}")
        print("Ensure headers are exactly: level, title, page_number")
        print("!" * 40 + "\n")

        choice = input("Would you like to try again? (y = retry file, p = change path, n = exit): ").lower()

        if choice == 'p':
            current_path = input("Enter new file path: ")
        elif choice != 'y':
            print("Exiting TOC extraction.")
            return None



def get_lock_status(filepath: Path) -> str:
    """
    Determines the specific nature of a file lock.

    Returns:
        str: 'local' if Excel is running on this machine,
             'remote' if the file is locked but Excel is not running locally,
             'none' if the file is accessible.
    """

    # 1. Check if Excel is running locally
    local_excel_active = False
    for process in psutil.process_iter(['name']):
        try:
            if process.info['name'] and 'EXCEL' in process.info['name'].upper():
                local_excel_active = True
                break
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    # 2. Check if the file itself is locked
    file_is_accessible = True
    if filepath.exists():
        try:
            # Atomic rename check for network/local locks
            os.rename(str(filepath), str(filepath))
        except (OSError, PermissionError):
            file_is_accessible = False

    # 3. Distinguish the cases
    if not file_is_accessible:
        return 'local' if local_excel_active else 'remote'

    return 'none'

def update_book_tracker(dana_code: str, folder_name: str) -> bool:
    """
    Safely updates the Excel tracker with the normalized folder name.
    """
    if not dana_code or not folder_name:
        logging.error("Invalid DanaCode or FolderName provided for Excel update.")
        return False

    try:
        # data_only=False ensures formulas in your sheet are preserved
        workbook = load_workbook(BOOK_TRACKER_EXCEL_FILE_PATH, data_only=False)
        sheet = workbook.active

        target_code = str(dana_code).strip()

        for row in sheet.iter_rows(min_col=2, max_col=2):
            cell = row[0]
            if cell.value and str(cell.value).strip() == target_code:
                sheet.cell(row=cell.row, column=12).value = folder_name # Column L (index 12)
                workbook.save(BOOK_TRACKER_EXCEL_FILE_PATH)
                return True

        logging.warning(f"DanaCode {target_code} not found in Column B.")
        return False

    except PermissionError:
        logging.error(f"Permission denied: {BOOK_TRACKER_EXCEL_FILE_PATH.name} is locked.")
        return False
    except FileNotFoundError:
        logging.error(f"Excel file missing at: {BOOK_TRACKER_EXCEL_FILE_PATH}")
        return False
    except InvalidFileException:
        logging.error("The tracker file is corrupted or not a valid .xlsx file.")
        return False
    except Exception as e:
        logging.error(f"An unexpected error occurred during Excel update: {e}")
        return False


def open_tracker_in_excel() -> None:
    """Opens the tracker using the system default handler for Excel files."""
    try:
        # Use str() for subprocess compatibility
        subprocess.run(['start', 'excel', str(BOOK_TRACKER_EXCEL_FILE_PATH)], shell=True, check=True)
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to open Excel: {e}")


def run_excel_update_workflow(dana_code: str, folder_name: str) -> bool:
    """
    Orchestrates the Excel update with specific feedback on lock types.
    """
    while True:
        status = get_lock_status(BOOK_TRACKER_EXCEL_FILE_PATH)

        if status == 'none':
            break

        if status == 'local':
            print_red("The tracker is open on your computer.")
            print("Please save and close your local Excel window.")

        else:
            print_red("The tracker is locked by another user or another computer.")
            print("Please wait for them to finish or ask them to close the file.")

        user_choice = input("\nPress Enter to retry, or type 'c' to cancel: ").strip().lower()
        if user_choice == 'c':
            return False  # Properly returns bool

    # Proceed with the update
    success = update_book_tracker(dana_code, folder_name)
    if success:
        open_tracker_in_excel()

    return success


def get_password_from_excel(book_title):
    import pandas as pd

    df = pd.read_excel(BOOK_TRACKER_EXCEL_FILE_PATH) # Load the Excel file

    # Define your column headers
    folder_col = "שם תיקייה בתיקיית העלאה לאמזון"
    password_col = "סיסמא"

    # Search for the row where the folder column matches book_title
    # We use .strip() to avoid errors from accidental spaces in the Excel cells
    result = df[df[folder_col].astype(str).str.strip() == book_title.strip()]

    if not result.empty:
        # Get the first match and return the password
        password = result.iloc[0][password_col]
        return str(password)

    else:
        print(f"Warning: No password found for {book_title} in Excel.")
        return None