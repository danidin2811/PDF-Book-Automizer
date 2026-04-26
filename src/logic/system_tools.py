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
    Organizes specific PDF sections into a 'flip' folder, deletes temporary files,
    and moves the entire book folder to a final archive destination.
    """

    from utils.input_output_tools import wait_for_ready_signal, print_red, print_green
    import shutil

    folder = Path(folder_path)
    flip_folder = folder / "flip"

    wait_for_ready_signal(
        "\nMANUAL ACTION REQUIRED:\n"
        "1. Close Adobe Acrobat and Excel.\n"
        "2. Ensure no files from this folder are open in any application.\n"
    )

    flip_folder.mkdir(exist_ok=True) # 1. Create the 'flip' sub-folder

    flip_suffixes = ("_fin.pdf", "_pre.pdf", "_chap.pdf", "_eng.pdf", "_con.pdf") # 2. Define the suffixes to move

    # 3. Scan and process files
    for item in folder.iterdir():
        if item.is_dir(): # Skip the 'flip' folder itself and any other sub-directories
            continue

        if item.name.lower().endswith(flip_suffixes): # Check if the file needs to be moved the 'flip' folder
            try:
                shutil.move(str(item), str(flip_folder / item.name))
                print(f"Moved {item.name} to flip")

            except Exception as e:
                print_red(f"Could not move {item.name}: {e}")

        else: # Delete all other files (like toc.csv, logs, or temporary snippets)
            try:
                item.unlink()
                print(f"Deleted: {item.name}")
            except Exception as e:
                print_red(f"Could not delete {item.name}: {e}")

    # 4. Move the entire folder to the archive destination
    try:
        from src.constants import READY_TO_UPLOAD_TO_AMAZON_FOLDER

        dest_path = Path(READY_TO_UPLOAD_TO_AMAZON_FOLDER) / folder.name
        shutil.move(str(folder), str(dest_path))
        print_green(f"Successfully archived folder to: {dest_path}")

        return dest_path

    except Exception as e:
        print_red(f"Failed to move entire folder to archive: {e}")

