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
        process_pdf()

        wait_for_ready_signal(
            "The PDF file has been successfully processed and the TOC has been added.\n"
            "----------------------------------------------------------------------\n"
            "MANUAL INSPECTION REQUIRED:\n"
            "1. Open the updated PDF.\n"
            "2. Verify all levels and page numbers in the bookmarks tab.\n"
            "3. Ensure 'toc.csv' matches the PDF structure.\n"
            "----------------------------------------------------------------------\n"
            "Once confirmed, press Enter to finish: "
        )

    except Exception as e:
        print(f"An error occurred during processing: {e}")

if __name__ == "__main__":
    main()