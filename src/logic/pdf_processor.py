import os
from pathlib import Path
from typing import Optional
import shutil

import pyperclip

from src.gemini.gemini_prompt import handle_gemini_toc_transcription
from src.constants import COVERS_FOLDER
from src.logic.excel_tools import process_toc_extraction, find_danacode_row
from src.logic.interface_controller import AppInterface
from src.logic.pdf_tools import get_pdf_page_count, extract_pdf_sections, handle_english_section_logic
from src.logic.file_operations import validate_pdf_path, move_cover_image
from utils.input_output_tools import *
from utils.input_output_tools import wait_for_ready_signal
from src.logic.pdf_tools import append_to_existing_toc
from utils.open_pdfs_side_by_side import open_pdfs_side_by_side_acrobat


def verify_and_rename_folder(source_folder: Path, target_folder_name: str, interface: AppInterface) -> bool:
    """
    Checks if the actual parent folder name matches the generated target book name.
    If it matches, it automatically skips the checkpoint. If it doesn't match,
    it copies the name to the clipboard and handles the abstract blocking checkpoint.

    Returns:
        bool: True if the process should continue, False if a GUI cancellation occurred.
    """
    # 1. Normalize and extract the parent directory's actual string name
    actual_parent_name = source_folder.name.strip().lower()

    if actual_parent_name == target_folder_name.strip().lower():
        # Using print_info or a clean log router so it prints cleanly to Terminal or UI Log Box
        interface.print_info(f"Folder matches target tracking layout name ('{target_folder_name}'). Skipping rename checkpoint.")
        return True

    # 2. If they do not match, proceed with the system clipboard action
    try:
        pyperclip.copy(target_folder_name)
    except Exception as e:
        interface.print_error(f"Failed to copy folder name to clipboard: {e}")

    # 3. Fire the abstracted blocking checkpoint
    interface.ask_checkpoint(
        "Rename Folder",
        f"Please rename the book folder to: '{target_folder_name}'\n(The text has been automatically copied to your system clipboard)."
    )

    # 4. Handle thread execution termination checks safely if running under UI instances
    if interface.is_gui and interface.ui.is_canceling:
        return False

    return True

def get_input_pdf_path(interface: AppInterface) -> Path | None:
    """
    Retrieves and validates a PDF path from the user.
    Supports instant retrieval from GUI configurations or fallback to interactive CLI loops.

    Returns:
        Path: Validated object pointing to the target PDF, or None if cancelled.
    """
    # --- 1. GUI EXECUTION PARADIGM ---
    if interface.is_gui:
        # Pull the pre-selected path directly from the UI state variable
        raw_input = interface.ui.source_file.get().strip().replace('"', '').replace("'", "")

        is_valid, error_message = validate_pdf_path(raw_input)
        if is_valid:
            return Path(raw_input)

        # If the GUI path is invalid, log the issue and exit gracefully
        interface.print_error(f"Invalid PDF Source Path: {error_message}")
        return None

    # --- 2. CLI EXECUTION PARADIGM ---
    prompt = "Enter the path of the PDF file (or drag and drop it here):"

    while True:
        # Abstracted via ask_string, which uses standard terminal input() in CLI mode
        raw_input = interface.ask_string("PDF Selection", prompt)

        if raw_input is None:
            return None

        raw_input = raw_input.strip().replace('"', '').replace("'", "")

        is_valid, error_message = validate_pdf_path(raw_input)
        if is_valid:
            return Path(raw_input)

        # Provide feedback and update prompt structure for next loop iteration
        interface.print_error(error_message)
        prompt = "Please try again. Drag and drop the PDF file and press Enter:"

def setup_working_directory(interface: AppInterface) -> tuple[Path, Path]:
    """
        Prompts for the PDF path and derives the folder context.
        Returns: (input_pdf_path, source_folder, folder_name)
    """
    input_pdf_path = get_input_pdf_path(interface)
    source_folder = input_pdf_path.parent  # get the source folder of the PDF file

    return input_pdf_path, source_folder

def run_cover_workflow(source_folder: Path, destination_folder: Path, interface) -> Optional[str]:
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

    checklist_msg = (
        "1. Close the Excel tracking table\n"
        "2. Ensure the numeric JPG cover is in the source folder\n"
        "3. Ensure the JPG filename matches the DanaCode\n\n"
    )

    interface.ask_checkpoint("Complete Checklist", checklist_msg)

    while True:
        # Attempt the silent logic operation
        dana_code = move_cover_image(source_folder, destination_folder, interface)

        if dana_code:
            # Note: No emojis used in professional output
            print(f"Successfully moved DanaCode: {dana_code}")
            return dana_code

        # Error handling with user feedback
        print_red(f"Error: No numeric JPG found in {source_folder.name}")

        if not interface.ask_yes_no("Try Again?", "Would you like to try again?: "):
            print("Operation cancelled by user.")
            return None


def run_extraction_workflow(input_pdf_path, source_folder, folder_name, interface):
    """
    Handles the physical file copying and section extraction logic.
    Supports asynchronous GUI parameters context fallback to CLI operations.
    Returns True to continue, False to stop the main process.
    """

    extract_sections = interface.ask_yes_no("Extract Sections", "Do you want to extract section PDFs? ")
    if not extract_sections:
        return True, None

    fin_file_path = source_folder / f"{folder_name}_fin.pdf"
    book_title = source_folder.name

    # 2. Duplicate PDF to Working Workspace
    copy_success = False
    while not copy_success:
        try:
            shutil.copy2(input_pdf_path, fin_file_path)
            interface.print_info(f"Created working file: {fin_file_path.name}")
            copy_success = True
        except PermissionError:
            retry = interface.ask_yes_no("Permission Denied", f"Permission denied: {input_pdf_path.name} is locked.\nPlease close the PDF. Retry? ")
            if not retry:
                return False, None

            else:
                interface.print_error(f"Permission denied: {input_pdf_path.name} is locked.")
                if not interface.ask_yes_no("File Open", "Do you want to try to close the PDF file and try again? "):
                    return False, None
        except Exception as e:
            interface.print_error(f"Failed to create fin file: {e}")
            return False, None

    total_pages = get_pdf_page_count(fin_file_path, interface)
    if not total_pages:
        return False, None

    sections = ['con', 'pre', 'chap']

    has_english = interface.ask_yes_no("Language Check", "Does the book have an English section?")
    if has_english:
        sections.append('english')

    ranges = interface.request_all_page_ranges(sections, total_pages)

    if not ranges or interface.is_canceling:
        interface.print_error("Configuration aborted by user when asking for ranges.")
        return False

    while True:
        success = extract_pdf_sections(book_title, fin_file_path, ranges, source_folder, interface)

        if success:
            if has_english:
                handle_english_section_logic(source_folder, folder_name, interface)
            break  # Exit loop safely

        # Error Recovery Handling
        if not interface.ask_yes_no("File Lock Error", "Extraction failed due to file lock. Retry?"):
            return False, None

    con_file_path = source_folder / f"{book_title}_con.pdf"
    if con_file_path.exists():
        interface.print_info(f"Path captured: {con_file_path}")

    return True, con_file_path


def add_toc_to_pdf(con_file_path, folder_name, input_pdf_path, source_folder, interface) -> bool:
    """
    Handles TOC transcription and applies bookmarks with a local retry loop.
    Returns True if successful, False if the user chooses to skip/cancel or hits abort.
    """
    import os
    from pathlib import Path

    interface.print_success(f"Ready for transcription: {con_file_path.name}")

    # Pass the unified interface context down the line
    handle_gemini_toc_transcription(source_folder, con_file_path, interface)
    interface.print_info("Back to add_toc_to_pdf")

    csv_path = os.path.join(source_folder, "toc.csv")
    output_pdf_path = os.path.join(source_folder, f"{Path(folder_name).stem}_fin.pdf")

    toc_entries = None

    while not toc_entries:
        # Check if the user initiated an app abort via the GUI window interface
        if interface.is_canceling:
            interface.print_error("Operation canceled by user.")
            return False

        toc_entries = process_toc_extraction(csv_path)

        if not toc_entries:
            # Route confirmation prompting dynamically through the interface helper methods
            retry = interface.confirm_choice("TOC entries are missing/invalid. Fix the CSV and try again?") if hasattr(
                interface, 'confirm_choice') else yes_or_no(
                "TOC entries are missing/invalid. Fix the CSV and try again? ")
            if not retry:
                return False

    while True:
        if interface.is_canceling:
            interface.print_error("Operation canceled by user.")
            return False

        success = append_to_existing_toc(input_pdf_path, output_pdf_path, toc_entries, interface)

        if success:
            interface.print_success("TOC applied successfully.")
            return True

        interface.print_error("\n[!] TOC Append Failed (Check if PDF is open in Acrobat).")

        retry_write = interface.confirm_choice("Fix the issue and retry writing bookmarks?") if hasattr(interface,'confirm_choice') else yes_or_no("Fix the issue and retry writing bookmarks? ")
        if not retry_write:
            return False

def process_pdf(input_pdf_path, source_folder, folder_name, interface) -> Optional[tuple]:
    """
    Executes the book processing pipeline.
    If a ui context is provided, requests inputs asynchronously through the GUI layout layer.
    Fallback to traditional CLI prompts if no UI context is supplied.
    """

    danacode = run_cover_workflow(source_folder, COVERS_FOLDER, interface)

    if not danacode:
        print_red("Process halted: Cover error.")
        return None  # return None on failure

    success, row_index = find_danacode_row(danacode)

    while not success:
        interface.print_error(f"Error: DanaCode '{danacode}' not found or Excel is locked.")

        if not interface.ask_yes_no("Try Again?", "Would you like to fix the Excel and try again? "):
            interface.print_info("User cancelled", "Exiting process.")

        success, row_index = find_danacode_row(danacode)

        interface.print_info("Danacode found", f"Danacode {danacode} found in row: {row_index}. Ready for updates.")

    success, con_file_path = run_extraction_workflow(input_pdf_path, source_folder, folder_name, interface)
    if not success:
        interface.print_error("Extraction workflow failed.")
        return None

    if con_file_path:
        add_toc_to_pdf(con_file_path, folder_name, input_pdf_path, source_folder, interface)

    fin_file_path = os.path.join(source_folder, f"{Path(folder_name).stem}_fin.pdf")

    open_pdfs_side_by_side_acrobat(str(con_file_path), str(fin_file_path))

    return fin_file_path, source_folder, row_index

if __name__ == "__main__":
    process_pdf()