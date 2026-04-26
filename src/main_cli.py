from src.logic.system_tools import clean_up_folder_after_processing
from utils.norm_book_title import normalize_book_title
from logic.pdf_processor import process_pdf
from utils.input_output_tools import wait_for_ready_signal

def main():
    metadata = normalize_book_title()

    if metadata:
        book_folder_name = metadata['folder_name']
        wait_for_ready_signal(f"ACTION REQUIRED: Rename the book folder to: {book_folder_name}")

    else:
        print("Using existing folder name.")

    try:
        book_folder_path = process_pdf()

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

        wait_for_ready_signal(
            "MANUAL ACTION REQUIRED:\n"
            "Before proceeding, make sure to remove 'Blank Page' bookmarks after adding the front and back covers.\n"
        )

        folder_in_amazon = clean_up_folder_after_processing(book_folder_path)



    except Exception as e:
        print(f"An error occurred during processing: {e}")

if __name__ == "__main__":
    main()