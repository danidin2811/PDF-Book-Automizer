import logging
from pathlib import Path
from typing import Optional, Dict, Tuple
from PyPDF2 import PdfReader, PdfWriter
import os

def append_to_existing_toc(input_pdf_path, output_pdf_path, new_toc_entries):
    """Clones existing PDF bookmarks and appends new ones from the CSV entries."""
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    writer.append_pages_from_reader(reader)

    def copy_outlines(outlines, parent=None):
        last_added = None
        for item in outlines:
            if isinstance(item, list):
                copy_outlines(item, parent=last_added)
            else:
                last_added = writer.add_outline_item(
                    title=item.title,
                    page_number=reader.get_destination_page_number(item),
                    parent=parent
                )

    if reader.outline:
        copy_outlines(reader.outline)

    bookmark_stack = {}
    for entry in new_toc_entries:
        level = entry["level"]
        title = entry["title"]
        page_index = entry["page"] - 1

        if page_index < 0 or page_index >= len(reader.pages):
            continue

        parent = bookmark_stack.get(level - 1)
        bookmark = writer.add_outline_item(
            title=title,
            page_number=page_index,
            parent=parent,
        )

        bookmark_stack[level] = bookmark
        keys_to_delete = [l for l in bookmark_stack if l > level]
        for l in keys_to_delete:
            del bookmark_stack[l]

    with open(output_pdf_path, "wb") as f:
        writer.write(f)


def reverse_pdf_pages(input_path: Path) -> bool:
    """
    Reverses the page order of a PDF and saves it with an '_eng' suffix.

    This is primarily used for the English sections of books processed at the Institute.

    Args:
        input_path (Path): The path to the PDF file to be reversed.

    Returns:
        bool: True if the file was created successfully, False otherwise.
    """

    # Define the output path using Pathlib's with_stem
    # Example: 'book.pdf' -> 'book_eng.pdf'
    output_path = input_path.with_stem(f"{input_path.stem}_eng")

    try:
        reader = PdfReader(input_path)
        writer = PdfWriter()

        # Reverse pages efficiently
        for page in reversed(reader.pages):
            writer.add_page(page)

        # Write the file using a context manager
        with open(output_path, "wb") as output_pdf:
            writer.write(output_pdf)

        logging.info(f"Successfully reversed: {output_path.name}")
        return True

    except (FileNotFoundError, PermissionError) as e:
        logging.error(f"File access error during reversal: {e}")
        return False

    except Exception as e:
        logging.error(f"Unexpected error reversing {input_path.name}: {e}")
        return False


def get_pdf_page_count(file_path: Path) -> Optional[int]:
    """
    Returns the total number of pages in a PDF file.
    """

    try:
        with open(file_path, 'rb') as pdf_file:
            return len(PdfReader(pdf_file).pages)

    except (FileNotFoundError, PermissionError) as e:
        logging.error(f"Access error for {file_path.name}: {e}")
        return None

    except Exception as e:
        logging.error(f"Unexpected error reading PDF: {e}")
        return None


def extract_pdf_sections(book_title: str, source_path: Path, ranges: Dict[str, Tuple[int, int]], output_dir: Path):
    """
    Core logic to slice a PDF into multiple section files.
    """

    reader = PdfReader(source_path)

    for section_name, (start, end) in ranges.items():
        writer = PdfWriter()

        for i in range(start - 1, end): # Adjusting for 0-indexed PyPDF2 pages
            writer.add_page(reader.pages[i])

        if section_name != "english":
            output_filename = output_dir / f"{book_title}_{section_name}.pdf"
        else:
            output_filename = output_dir / f"{book_title}.pdf"

        with open(output_filename, "wb") as output_pdf:
            writer.write(output_pdf)


def handle_english_section_logic(source_folder: Path, folder_name: str) -> bool:
    """
    Handles the post-extraction logic for the English section:
    Reverses the pages and renames the file to the official format.
    """
    temp_eng_file = source_folder / f"{folder_name}.pdf"

    if not temp_eng_file.exists():
        logging.error(f"English extraction failed: {temp_eng_file} not found.")
        return False

    # Perform reversal
    if reverse_pdf_pages(temp_eng_file):
        try:
            os.remove(temp_eng_file)
            return True

        except Exception as e:
            logging.error(f"Error during English file cleanup: {e}")
            return False

    return False