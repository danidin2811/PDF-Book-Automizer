import logging
import os
from src.constants import COVERS_FOLDER
from pathlib import Path
from typing import Optional
import shutil

from src.logic.pdf_tools import get_pdf_page_count, extract_pdf_sections, handle_english_section_logic
from src.logic.excel_tools import run_excel_update_workflow
from src.logic.file_operations import validate_pdf_path, move_cover_image
from src.logic.system_tools import delete_file
from utils.input_output_tools import print_green, print_red, yes_or_no


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

def get_page_range_ui(section: str, total_pages: int):
    """UI function to get ranges from user."""
    while True:
        try:
            start = int(input(f"Enter start page for {section.upper()}: "))
            end = int(input(f"Enter end page for {section.upper()}: "))

            if 1 <= start <= end <= total_pages:
                return start, end

            print_red(f"Invalid range. Total pages: {total_pages}")

        except ValueError:
            print_red("Please enter numbers only.")

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

def process_pdf():
    # 1. Setup paths
    input_pdf_path = get_input_pdf_path()
    source_folder = input_pdf_path.parent # get the source folder of the PDF file
    folder_name = str(source_folder.name)

    # 2. Process Cover and Excel
    danacode = run_cover_workflow(source_folder, COVERS_FOLDER)

    if not danacode:
        print_red("Process halted: Cover error.")

    if not run_excel_update_workflow(danacode, folder_name):
        print_red("Process halted: Excel error.")
        return

        # 3. Handle PDF Extraction
    if yes_or_no("Do you want to extract section PDFs? "):
        fin_file_path = source_folder / f"{folder_name}_fin.pdf"

        try:
            shutil.copy2(input_pdf_path, fin_file_path)
            print_green(f"Created working file: {fin_file_path.name}")
        except Exception as e:
            print_red(f"Failed to create fin file: {e}")
            return

        total_pages = get_pdf_page_count(fin_file_path)

        if total_pages:
            ranges = {}
            for sec in ['con', 'pre', 'chap']:
                ranges[sec] = get_page_range_ui(sec, total_pages)

            # Extract English if it exists
            if yes_or_no("Does the book have an English section? "):
                ranges['english'] = get_page_range_ui('english', total_pages)

                extract_pdf_sections(folder_name, fin_file_path, ranges, source_folder)

                if handle_english_section_logic(source_folder, folder_name):
                    print_green(f"Successfully processed English section for {folder_name}")

            else:
                extract_pdf_sections(fin_file_path, ranges, source_folder)

if __name__ == "__main__":
    process_pdf()