import pyperclip

from src.fliphtml5.flip_html_automation import fliphtml5_automation
from src.logic.file_operations import check_file_size
from src.logic.system_tools import clean_up_folder_after_processing
from utils.norm_book_title import normalize_book_title
from src.logic.pdf_processor import process_pdf, setup_working_directory, verify_and_rename_folder
from utils.input_output_tools import wait_for_ready_signal, yes_or_no, print_red
from src.logic.excel_tools import run_excel_update_workflow
from src.logic.interface_controller import AppInterface

def main():
    interface = AppInterface(ui=None)

    # 1. Run setup checks to pull down paths
    input_pdf_path, source_folder, folder_name_placeholder = setup_working_directory(interface)
    if input_pdf_path is None:
        return

    # 2. Capture the validated english metadata matrix
    book_titles = normalize_book_title(interface)

    if book_titles:
        target_folder_name = book_titles['folder_name']

        # 3. Hand off to the exact same logic processor!
        verify_and_rename_folder(source_folder, target_folder_name, interface)

    book_folder_path = None
    fin_file_path = ''
    book_row_index_in_table = 0

    while book_folder_path is None:
        try:
            result = process_pdf(input_pdf_path, source_folder, folder_name, interface)

            if result is None:  # If process_pdf returned None on early failure
                if not yes_or_no("\n[!] Error encountered. RETRY? "):
                    return
                continue

            fin_file_path, book_folder_path, book_row_index_in_table = result

        except Exception as e:
            print_red(f"An error occurred during processing: {e}")
            if not yes_or_no("Unexpected error. Try again? "):
                return

    wait_for_ready_signal(
        "The PDF file has been successfully processed and the TOC has been added.\n"
        "----------------------------------------------------------------------\n"
        "MANUAL INSPECTION REQUIRED:\n"
        "Open the updated PDF and verify all levels, page numbers and titles in the bookmarks tab.\n"
        "----------------------------------------------------------------------\n"
    )

    wait_for_ready_signal(
        "MANUAL ACTION REQUIRED:\n"
        "Great, you've finished inspecting the TOC, to proceed, please add front and back covers as needed.\n"
    )
    print("Back to main")
    wait_for_ready_signal(
        "MANUAL ACTION REQUIRED:\n"
        "Before proceeding, make sure to remove 'Blank Page' bookmarks after adding the front and back covers.\n"
    )
    print("Back to main")

    check_file_size(fin_file_path)

    folder_in_amazon = clean_up_folder_after_processing(str(book_folder_path))

    fliphtml5_automation(folder_in_amazon, book_titles['display_title'], book_row_index_in_table)

    run_excel_update_workflow(book_row_index_in_table, book_titles['folder_name'])

    print("Workflow complete!")

if __name__ == "__main__":
    main()