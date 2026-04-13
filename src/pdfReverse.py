from removeBleedBox import print_green
import argparse
from PyPDF2 import PdfReader, PdfWriter
import os

def print_red(text):
    """Prints text in red color."""
    print("\033[31m" + text + "\033[0m")

def reversePages(file_path):
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

        print_green(f"Reversed english file successfully")

    except Exception as e:
        print_red(f"\nAn error occurred: {e}")

def main():
    file_path = input("Please enter the full path of the PDF file ").strip('"').strip()
    if not file_path.endswith(".pdf"):
        print_red("Please provide a valid PDF file.")
        return
    
    reversePages(file_path)

if __name__ == "__main__":
    main()
