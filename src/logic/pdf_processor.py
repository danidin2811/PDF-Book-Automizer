import re
import os
import shutil
from src.constants import COVERS_FOLDER
from PyPDF2 import PdfReader

def print_red(text):
    """Prints text in red color."""
    print("\033[31m" + text + "\033[0m")

def print_green(text):
    """Prints text in green color."""
    print("\033[32m" + '\n' + text + "\033[0m")

def validate_pdf_file(input_pdf_path):
    """Validates the given PDF file path."""

    if not os.path.isfile(input_pdf_path):
        return False, "Invalid file path. Please check and try again."
    if not input_pdf_path.lower().endswith('.pdf'):
        return False, "This file is not a PDF."
    return True, None

def get_input_pdf():
    input_pdf_path = input("\nEnter the path of the PDF file: ").replace('"', '')

    while True:
        is_valid, error = validate_pdf_file(input_pdf_path)
        if is_valid:
            break
        print_red(error)
        input_pdf_path = input("\nDrag and drop the PDF file here and press Enter: ").strip('"')

    return input_pdf_path

def yesOrNo(prompt):
    while True:
        choice = input(prompt).strip().lower()
        if choice in {"y", "yes"}:
            return True
        elif choice in {"n", "no"}:
            return False
        print_red("Invalid input. Please enter 'y' for yes or 'n' for no.")

def moveJpgFile(sourceFolder, destinationFolder):
    while True:
        if not os.path.exists(destinationFolder):
            os.makedirs(destinationFolder)

        for fileName in os.listdir(sourceFolder):
            if re.match(r"^\d{4}\.jpg$", fileName):  # Matches filenames like "1234.jpg"
                sourcePath = os.path.join(sourceFolder, fileName)
                destinationPath = os.path.join(destinationFolder, fileName)
                shutil.move(sourcePath, destinationPath)
                print_green(f"Moved: {fileName} successfully")
                return fileName.removesuffix(".jpg")  # Exit the function after successfully moving the file
        else:
            print_red("\njpg file of front cover with danacode file name wasn't found\n")
            choice = yesOrNo("Would you like to try again? (y/n): ")

            if not choice:
                print("\nok")
                return None


def get_pdf_page_count(file_path):
    """Returns the total number of pages in a PDF file."""

    try:
        with open(file_path, 'rb') as pdf_file:
            return len(PdfReader(pdf_file).pages)
    except Exception as e:
        print_red(f"Error reading PDF file: {e}")
        return None

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

def extractSectionPages(file_path, folderName, withCover=False):
    total_pages = get_pdf_page_count(file_path)
    thereIsEnglish = ''

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


def delete_file(filePath, fileName):
    print("delete file")
    try:
        if os.path.exists(filePath):
            os.remove(filePath)
            print_green(f"{fileName} deleted successfully.\n")
        else:
            print_red(f"\nThe file {filePath} does not exist.")
    except Exception as e:
        print_red(f"\nError deleting the file: {e}")

def extract_pages(finFileName, input_pdf_path, folderName):
    if yesOrNo("\nDo you want to extract PDFs for con, pre, chap and eng? (y/n): "):
        thereIsEnglish = extractSectionPages(os.path.join(folderName, finFileName), folderName, True)

        if thereIsEnglish == 'y':
            engFile = os.path.join(os.path.dirname(input_pdf_path), f"{folderName}.pdf")
            reverse_pages(engFile)
            delete_file(engFile, 'temp english file')

def process_pdf():
    input_pdf_path = get_input_pdf()

    source_folder = os.path.dirname(input_pdf_path) # get the source folder of the PDF file
    folder_name = os.path.basename(source_folder)

    updateExcelCellAndOpenExcel(moveJpgFile(source_folder, COVERS_FOLDER), folder_name)

    fin_file_name = os.path.join(os.path.dirname(input_pdf_path), f"{folder_name}_fin.pdf")

    extract_pages(fin_file_name, input_pdf_path, folder_name)

if __name__ == "__main__":
    process_pdf()