import os
from pathlib import Path
from typing import Optional
import shutil

from src.gemini.gemini_prompt import handle_gemini_toc_transcription
from src.constants import COVERS_FOLDER
from src.logic.excel_tools import process_toc_extraction, find_danacode_row
from src.logic.pdf_tools import get_pdf_page_count, extract_pdf_sections, handle_english_section_logic
from src.logic.file_operations import validate_pdf_path, move_cover_image
from utils.input_output_tools import *
from utils.input_output_tools import wait_for_ready_signal
from src.logic.pdf_tools import append_to_existing_toc
from utils.open_pdfs_side_by_side import open_pdfs_side_by_side_acrobat


def get_input_pdf_path() -> Path:
    """
        Retrieves and validates a PDF path from the user.

        Handles Windows 11 drag-and-drop quote cleaning and ensures file validity.

        Returns:
            Path: Validated object pointing to the target PDF.
        """

    prompt = "\nEnter the path of the PDF file (or drag and drop it here): "

    while True:
        raw_input = input(prompt).strip().replace('"', '')

        is_valid, error_message = validate_pdf_path(raw_input)

        if is_valid:
            return Path(raw_input)

        # Provide feedback and loop back
        print_red(f"Error: {error_message}")
        prompt = "\nPlease try again. Drag and drop the PDF file and press Enter: "

def setup_working_directory() -> tuple[Path, Path, str]:
    """
        Prompts for the PDF path and derives the folder context.
        Returns: (input_pdf_path, source_folder, folder_name)
    """
    print("now in setup_working_directory")
    input_pdf_path = get_input_pdf_path()
    source_folder = input_pdf_path.parent  # get the source folder of the PDF file
    folder_name = str(source_folder.name)

    return input_pdf_path, source_folder, folder_name

def run_cover_workflow(source_folder: Path, destination_folder: Path) -> Optional[str]:
    """
    Orchestrates the movement of a book cover JPG based on its DanaCode.

    This function continuously attempts to locate and move a numeric JPG file from the source to the destination.
    If the file is missing, it prompts the user to retry or exit.

    Args:
        source_folder (Path): The directory containing the raw PDF and JPG.
        destination_folder (Path): The central archival folder for covers.

    Returns:
        Optional[str]: The extracted DanaCode string if successful;
                      None if the user chooses to cancel.
    """

    while True:
        # Attempt the silent logic operation
        dana_code = move_cover_image(source_folder, destination_folder)

        if dana_code:
            # Note: No emojis used in professional output
            print_green(f"Successfully processed DanaCode: {dana_code}")
            return dana_code

        # Error handling with user feedback
        print_red(f"Error: No numeric JPG found in {source_folder.name}")

        if not yes_or_no("Would you like to try again? (y/n): "):
            print("Operation cancelled by user.")
            return None

def run_extraction_workflow(input_pdf_path: Path, source_folder: Path, folder_name: str) -> tuple[bool, Optional[Path]]:
    """
    Handles the physical file copying and section extraction logic.
    Returns True to continue, False to stop the main process.
    """

    if not yes_or_no("Do you want to extract section PDFs? "):
        return True, None

    fin_file_path = source_folder / f"{folder_name}_fin.pdf"
    book_title = source_folder.name

    copy_success = False
    while not copy_success:
        try:
            shutil.copy2(input_pdf_path, fin_file_path)
            print_green(f"Created working file: {fin_file_path.name}")
            copy_success = True

        except PermissionError:
            print_red(f"Permission denied: {input_pdf_path.name} is locked.")
            if not yes_or_no("Please close the PDF and try again? "):
                return False, None

        except Exception as e:
            print_red(f"Failed to create fin file: {e}")
            return False, None

    total_pages = get_pdf_page_count(fin_file_path)
    if not total_pages:
        return False, None

    ranges = {}
    for sec in ['con', 'pre', 'chap']:
        ranges[sec] = get_page_range_ui(sec, total_pages)

    has_english = yes_or_no("Does the book have an English section? ")
    if has_english:
        ranges['english'] = get_page_range_ui('english', total_pages)

    # LOCAL RETRY LOOP for the extraction logic
    while True:
        success = extract_pdf_sections(book_title, fin_file_path, ranges, source_folder)

        if success:
            if has_english:
                handle_english_section_logic(source_folder, folder_name)
            break  # Exit retry loop

        if not yes_or_no("Extraction failed due to file lock. Retry? "):
            return False, None

    con_file_path = source_folder / f"{book_title}_con.pdf"

    if con_file_path.exists():
        print_green(f"Path captured: {con_file_path}")

    return True, con_file_path

def add_toc_to_pdf(con_file_path, folder_name, input_pdf_path, source_folder) -> bool:
    """
        Handles TOC transcription and applies bookmarks with a local retry loop.
        Returns True if successful, False if the user chooses to skip/cancel.
    """

    print_green(f"Ready for transcription: {con_file_path.name}")

    handle_gemini_toc_transcription(source_folder, con_file_path)
    print("Back to add_toc_to_pdf")
    csv_path = os.path.join(source_folder, "toc.csv")
    output_pdf_path = os.path.join(source_folder, f"{Path(folder_name).stem}_fin.pdf")

    toc_entries = None

    while not toc_entries:
        toc_entries = process_toc_extraction(csv_path)

        if not toc_entries:
            if not yes_or_no("TOC entries are missing/invalid. Fix the CSV and try again? "):
                return False

    while True:
        success = append_to_existing_toc(input_pdf_path, output_pdf_path, toc_entries)

        if success:
            print_green("TOC applied successfully.")
            return True

        print_red("\n[!] TOC Append Failed (Check if PDF is open in Acrobat).")
        if not yes_or_no("Fix the issue and retry writing bookmarks? "):
            return False

def process_pdf() -> tuple[str, Path, int | None] | None:
    checklist = (
        "\nPRE-PROCESSING CHECKLIST:\n"
        "1. Close the Excel tracking table\n"
        "2. Ensure the numeric JPG cover is in the source folder\n"
        "3. Ensure the JPG filename matches the DanaCode\n\n"
    )

    wait_for_ready_signal(checklist)
    print("Back in process_pdf")
    input_pdf_path, source_folder, folder_name = setup_working_directory() # 1. Setup paths

    # 2. Process Cover and Excel
    danacode = run_cover_workflow(source_folder, COVERS_FOLDER)

    if not danacode:
        print_red("Process halted: Cover error.")
        return None  # return None on failure

    success, row_index = find_danacode_row(danacode)

    while not success:
        print_red(f"Error: DanaCode '{danacode}' not found or Excel is locked.")

        if not yes_or_no("Would you like to fix the Excel and try again? "):
            print("User cancelled. Exiting process.")
            return None

        success, row_index = find_danacode_row(danacode)

    print_green(f"Danacode {danacode} found in row: {row_index}. Ready for updates.")

    # 3. Handle PDF Extraction
    success, con_file_path = run_extraction_workflow(input_pdf_path, source_folder, folder_name)
    if not success:
        print_red("Extraction workflow failed.")
        return None

    if con_file_path:
        add_toc_to_pdf(con_file_path, folder_name, input_pdf_path, source_folder)

    fin_file_path = os.path.join(source_folder, f"{Path(folder_name).stem}_fin.pdf")

    open_pdfs_side_by_side_acrobat(str(con_file_path), str(fin_file_path))

    return fin_file_path, source_folder, row_index

if __name__ == "__main__":
    process_pdf()