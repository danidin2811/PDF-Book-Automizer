from PyPDF2 import PdfReader, PdfWriter
import os
from removeBleedBox import print_green

def debugPrint(text):
    print(text)

def yesOrNo(prompt):
    while True:
        choice = input(prompt).strip().lower()
        if choice in {"y", "yes"}:
            return True
        elif choice in {"n", "no"}:
            return False
        print_red("Invalid input. Please enter 'y' for yes or 'n' for no.")


def extract_pages(input_pdf_path, ranges, folderName):
    """
    Extracts specific page ranges from an input PDF and saves them as separate PDF files.

    Parameters:
    ----------
    input_pdf_path : str - The full path to the input PDF file from which pages will be extracted.

    ranges : dict
        A dictionary where keys are section names (e.g., "con", "chap") and values are tuples of (start_page, end_page), both inclusive.
        Page numbers are 1-based.

        Example:
        {
            "con": (1, 10),
            "chap": (11, 20)
        }

    folderName : str
        A base name used to generate output file names. If the section is "english", the file is saved
        as `<folderName>.pdf`. For other sections, files are saved as `<folderName>_<section>.pdf`.

    Behavior:
    --------
    - For each entry in the `ranges` dictionary:
        - Pages in the specified range are extracted (inclusive).
        - A new PDF is created containing only those pages.
        - The output PDF is saved in the same directory as the input file.
    - A success message is printed for each section after extraction.

    Notes:
    -----
    - Page ranges are converted from 1-based to 0-based indexing internally.
    - Existing files with the same names will be overwritten without warning.
    """

    reader = PdfReader(input_pdf_path)
    
    for section, page_range in ranges.items():
        writer = PdfWriter()
        
        # Extract pages based on the range (inclusive)
        for page_num in range(page_range[0] - 1, page_range[1]):
            writer.add_page(reader.pages[page_num])
        
        # Save the extracted pages as a new PDF
        if section == "english":
            output_pdf_path = os.path.join(os.path.dirname(input_pdf_path), f"{folderName}.pdf")
        else:
            output_pdf_path = os.path.join(os.path.dirname(input_pdf_path), f"{folderName}_{section}.pdf")
            
        with open(output_pdf_path, "wb") as output_pdf:
            writer.write(output_pdf)
        
        print_green(f"Successfully extracted {section}")

def print_red(text):
    """Prints text in red color."""
    print("\033[31m" + text + "\033[0m")

def get_page_range(section, total_pages, withCover):
    while True:
        try:
            start = int(input(f"\nEnter the start page for {section.upper()}: "))
            if not (1 <= start <= total_pages):
                print_red(f"\nPage number must be between 1 and {total_pages}.")
                continue  # Restart loop
            
            end = int(input(f"\nEnter the end page for {section.upper()}: "))
            if not (start <= end <= total_pages):
                print_red(f"\nPage number must be between {start} and {total_pages}.")
                continue  # Restart loop
            
            print(f"\nYou entered start page: {start}, end page: {end} for section: {section}")
            confirm = yesOrNo("Are these correct? (y/n): ")
            
            if confirm:
                return (start + 2, end + 2) if withCover else (start, end)
            
            print_red("Re-entering page range...\n")
        except ValueError:
            print_red("\nPlease enter a valid number.")

def get_pdf_page_count(file_path):
    """Returns the total number of pages in a PDF file."""
    try:
        with open(file_path, 'rb') as pdf_file:
            pdf_reader = PdfReader(pdf_file)
            return len(pdf_reader.pages)
    except Exception as e:
        print_red(f"Error reading PDF file: {e}")
        return None

def extractSectionPages(file_path, folderName, withCover = False):
    total_pages = get_pdf_page_count(file_path)
    
    # Get ranges for each section
    ranges = {}
    for section in ['con', 'pre', 'chap', 'english']:
        if section == 'english':
            thereIsEnglish = yesOrNo("Does the book have an English section? ")
            if thereIsEnglish:
                ranges[section] = get_page_range(section, total_pages, withCover)
                
        else:
            ranges[section] = get_page_range(section, total_pages, withCover)
    
    extract_pages(file_path, ranges, folderName)
    
    return thereIsEnglish
    
def main():
    input_pdf_path = input("Enter the path of the PDF file: ").replace('"','')
    extractSectionPages(input_pdf_path, os.path.basename(os.path.dirname(input_pdf_path)))

if __name__ == "__main__":
    main()