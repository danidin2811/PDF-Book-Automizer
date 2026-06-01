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
        interface.print_info("System Check",
                             f"Folder matches target tracking layout name ('{target_folder_name}'). Skipping rename checkpoint.")
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

def setup_working_directory(interface: AppInterface) -> tuple[Path, Path, str]:
    """
        Prompts for the PDF path and derives the folder context.
        Returns: (input_pdf_path, source_folder, folder_name)
    """
    print("now in setup_working_directory")
    input_pdf_path = get_input_pdf_path(interface)
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
            print(f"Successfully moved DanaCode: {dana_code}")
            return dana_code

        # Error handling with user feedback
        print_red(f"Error: No numeric JPG found in {source_folder.name}")

        if not yes_or_no("Would you like to try again? (y/n): "):
            print("Operation cancelled by user.")
            return None


def run_extraction_workflow(input_pdf_path, source_folder, folder_name, ui=None):
    """
    Handles the physical file copying and section extraction logic.
    Supports asynchronous GUI parameters context fallback to CLI operations.
    Returns True to continue, False to stop the main process.
    """

    # 1. Determine if we skip extraction
    if ui is not None:
        # If the GUI context exists, check if user cancelled or unselected extraction
        if ui.is_canceling:
            return False, None

        # We handle this logic before process_pdf via GUI screens,
        # so if this function is hit, it means the user already clicked "Yes".
        extract_sections = True

    else: # Legacy Terminal Fallback behavior
        extract_sections = yes_or_no("Do you want to extract section PDFs? ")
        if not extract_sections:
            return True, None

    fin_file_path = source_folder / f"{folder_name}_fin.pdf"
    book_title = source_folder.name

    # 2. Duplicate PDF to Working Workspace
    copy_success = False
    while not copy_success:
        try:
            shutil.copy2(input_pdf_path, fin_file_path)
            print(f"Created working file: {fin_file_path.name}")  # Will route to GUI log box automatically
            copy_success = True
        except PermissionError:
            if ui is not None:
                retry = ui.async_ask_yes_no("Permission Denied",
                                            f"Permission denied: {input_pdf_path.name} is locked.\nPlease close the PDF. Retry?")
                if ui.is_canceling or not retry:
                    return False, None
            else:
                print_red(f"Permission denied: {input_pdf_path.name} is locked.")
                if not yes_or_no("Please close the PDF and try again? "):
                    return False, None
        except Exception as e:
            print(f"Failed to create fin file: {e}")
            return False, None

    total_pages = get_pdf_page_count(fin_file_path)
    if not total_pages:
        return False, None

    # 3. RANGE DICTIONARY POPULATION (GUI Matrix vs CLI Loops)
    ranges = {}
    has_english = False

    if ui is not None:
        gui_ranges = getattr(ui, 'current_ranges_data', {})

        # Restructure your GUI output dictionary to match what extract_pdf_sections expects
        for sec in ['con', 'pre', 'chap']:
            if f"{sec}_start" in gui_ranges:
                ranges[sec] = (gui_ranges[f"{sec}_start"], gui_ranges[f"{sec}_end"])

        if "eng_start" in gui_ranges:
            has_english = True
            ranges['english'] = (gui_ranges["eng_start"], gui_ranges["eng_end"])
    else:
        # --- LEGACY CLI LAYER POPULATION ---
        for sec in ['con', 'pre', 'chap']:
            ranges[sec] = get_page_range_ui(sec, total_pages)

        has_english = yes_or_no("Does the book have an English section? ")
        if has_english:
            ranges['english'] = get_page_range_ui('english', total_pages)

    # 4. EXECUTE EXTRACTION WORKFLOW RETRY LOOPS
    while True:
        success = extract_pdf_sections(book_title, fin_file_path, ranges, source_folder)

        if success:
            if has_english:
                handle_english_section_logic(source_folder, folder_name)
            break  # Exit loop safely

        # Error Recovery Handling
        if ui is not None:
            retry = ui.async_ask_yes_no("Extraction Interrupted", "Extraction failed due to file lock. Retry?")
            if ui.is_canceling or not retry:
                return False, None
        else:
            if not yes_or_no("Extraction failed due to file lock. Retry? "):
                return False, None

    con_file_path = source_folder / f"{book_title}_con.pdf"
    if con_file_path.exists():
        print(f"Path captured: {con_file_path}")

    return True, con_file_path

def add_toc_to_pdf(con_file_path, folder_name, input_pdf_path, source_folder, ui) -> bool:
    """
        Handles TOC transcription and applies bookmarks with a local retry loop.
        Returns True if successful, False if the user chooses to skip/cancel.
    """

    print_green(f"Ready for transcription: {con_file_path.name}")

    handle_gemini_toc_transcription(source_folder, con_file_path, ui)
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

def process_pdf(input_pdf_path, source_folder, folder_name, interface) -> Optional[tuple]:
    """
    Executes the book processing pipeline.
    If a ui context is provided, requests inputs asynchronously through the GUI layout layer.
    Fallback to traditional CLI prompts if no UI context is supplied.
    """

    checklist_msg = (
        "1. Close the Excel tracking table\n"
        "2. Ensure the numeric JPG cover is in the source folder\n"
        "3. Ensure the JPG filename matches the DanaCode\n\n"
    )

    interface.ask_checkpoint("Complete Checklist", checklist_msg)

    danacode = run_cover_workflow(source_folder, COVERS_FOLDER)

    if not danacode:
        print_red("Process halted: Cover error.")
        return None  # return None on failure

    success, row_index = find_danacode_row(danacode)

    while not success:
        print_red(f"Error: DanaCode '{danacode}' not found or Excel is locked.")

        if ui is not None:
            retry = ui.async_ask_yes_no("Excel Table Missing Record", "Would you like to fix the Excel and try again?")
            if ui.is_canceling or not retry:
                print("User cancelled or closed application pipeline context. Exiting process.")
                return None
        else:
            if not yes_or_no("Would you like to fix the Excel and try again? "):
                print("User cancelled. Exiting process.")
                return None

        success, row_index = find_danacode_row(danacode)

        print_green(f"Danacode {danacode} found in row: {row_index}. Ready for updates.")

    success, con_file_path = run_extraction_workflow(input_pdf_path, source_folder, folder_name, ui=ui)
    if not success:
        print_red("Extraction workflow failed.")
        return None

    if con_file_path:
        add_toc_to_pdf(con_file_path, folder_name, input_pdf_path, source_folder, ui=ui)

    fin_file_path = os.path.join(source_folder, f"{Path(folder_name).stem}_fin.pdf")

    open_pdfs_side_by_side_acrobat(str(con_file_path), str(fin_file_path))

    return fin_file_path, source_folder, row_index

if __name__ == "__main__":
    process_pdf()