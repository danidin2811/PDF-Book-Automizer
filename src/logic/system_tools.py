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


def clean_up_folder_after_processing(folder_path: str):
    """
    Organizes files into 'flip', deletes temps, and moves folder to archive.
    Retries locally if files are locked by other processes.
    """

    from utils.input_output_tools import wait_for_ready_signal, print_red, print_green, yes_or_no
    import shutil

    print("in clean_up_folder_after_processing")

    wait_for_ready_signal(
        "\nMANUAL ACTION REQUIRED:\n"
        "1. Close Adobe Acrobat and Excel.\n"
        "2. Ensure no files from this folder are open in any application.\n"
    )

    print("Back to clean")

    folder = Path(folder_path)

    while True:  # Outer loop: Retries the entire cleanup process
        try:
            flip_folder = folder / "flip"
            flip_folder.mkdir(exist_ok=True)

            flip_suffixes = ("_fin.pdf", "_pre.pdf", "_chap.pdf", "_eng.pdf", "_con.pdf")

            # Stage 1: Move and Delete individual files
            for item in list(folder.iterdir()):  # Use list() to avoid iterator issues during move/delete
                if item.is_dir() or not item.exists():
                    continue

                if item.name.lower().endswith(flip_suffixes):
                    shutil.move(str(item), str(flip_folder / item.name))
                    print(f"Moved {item.name} to flip")

                else:
                    item.unlink()
                    print(f"Deleted: {item.name}")

            # Stage 2: Move the entire folder to Archive
            from src.constants import READY_TO_UPLOAD_TO_AMAZON_FOLDER
            dest_path = Path(READY_TO_UPLOAD_TO_AMAZON_FOLDER) / folder.name

            shutil.move(str(folder), str(dest_path))

            print_green(f"Successfully archived folder to: {dest_path}")
            return dest_path / "flip"  # Success exit

        except (PermissionError, OSError) as e:
            print_red(f"\n[!] Cleanup Blocked: {e}")
            print("A file or folder is still open in another program (Acrobat, Excel, or File Explorer).")

            if not yes_or_no("Would you like to close the programs and RETRY cleanup?"):
                print_red("Cleanup aborted. The folder remains in its current location.")
                return None

            print("Retrying cleanup...\n")