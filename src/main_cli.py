from utils.norm_book_title import normalize_book_title
from logic.pdf_processor import process_pdf
from utils.input_output_tools import wait_for_ready_signal

if __name__ == "__main__":
    metadata = normalize_book_title()

    if metadata:
        book_folder_name = metadata['folder_name']
        wait_for_ready_signal(f"Please change the folder name of the book folder to {book_folder_name}")

    else:
        # Logic for when the user bypassed the title change
        print("Using existing folder name.")

    process_pdf()