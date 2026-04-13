from PyPDF2 import PdfReader, PdfWriter
import os

def reverse_pages(file_path):
    print("reverse pages")
    try:
        # Read the PDF file
        reader = PdfReader(file_path)
        writer = PdfWriter()

        # Reverse the pages and add them to the writer
        for page in reversed(reader.pages):
            writer.add_page(page)

        output_path = os.path.splitext(file_path)[0] + "_eng.pdf"
        with open(output_path, "wb") as output_pdf:
            writer.write(output_pdf)

        return True

    except Exception as e:
        return False

def extract_pages_logic():
