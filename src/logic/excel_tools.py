import os
import subprocess
import logging
from pathlib import Path
import psutil
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from src.constants import BOOK_TRACKER_EXCEL_FILE_PATH
from utils.input_output_tools import print_red

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')


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