from src.logic.system_tools import clean_up_folder_after_processing
from utils.norm_book_title import normalize_book_title
from logic.pdf_processor import process_pdf
from utils.input_output_tools import wait_for_ready_signal, yes_or_no


def main():
    metadata = normalize_book_title()

    if metadata:
        book_folder_name = metadata['folder_name']
        wait_for_ready_signal(f"ACTION REQUIRED: Rename the book folder to: {book_folder_name}")

    book_folder_path = None

    while book_folder_path is None:
        try:
            book_folder_path = process_pdf()

            if book_folder_path is None:
                print("\n[!] The process encountered an error (Check Excel or Cover).")

                if not yes_or_no("Would you like to fix the error and RETRY? "):
                    print("Exiting script...")
                    return  # Exit the whole program

                print("Restarting process...\n")

        except Exception as e:
            print(f"An error occurred during processing: {e}")
            if not yes_or_no("Unexpected error. Try again?"):
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

    wait_for_ready_signal(
        "MANUAL ACTION REQUIRED:\n"
        "Before proceeding, make sure to remove 'Blank Page' bookmarks after adding the front and back covers.\n"
    )

    folder_in_amazon = clean_up_folder_after_processing(book_folder_path)

    print("Workflow complete!")

if __name__ == "__main__":
    main()