import psutil
import logging
from pathlib import Path

def is_excel_running() -> bool:
    """
    Checks the system process list for an active Excel instance.

    Returns:
        bool: True if Excel is detected, False otherwise.
    """

    for process in psutil.process_iter(['name']):
        try:
            if process.info['name'] and 'EXCEL' in process.info['name'].upper():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False

def delete_file(file_path: Path) -> bool:
    """
    Safely deletes a file from the system.
    """
    try:
        if file_path.exists():
            file_path.unlink()
            return True

        logging.warning(f"Delete failed: {file_path} does not exist.")
        return False

    except PermissionError:
        logging.error(f"Delete failed: {file_path.name} is currently in use.")
        return False

    except Exception as e:
        logging.error(f"Error deleting file: {e}")
        return False

def clean_up_folder_after_processing(folder_path):
    from utils.input_output_tools import wait_for_ready_signal

    wait_for_ready_signal(
        "MANUAL ACTION REQUIRED:\n"
        "Please close all open files of the book.\n"
    )