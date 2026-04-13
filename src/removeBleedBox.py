import subprocess
import os
import itertools
from PyPDF2 import PdfReader, PdfWriter

def print_green(text):
    """Prints text in green color."""
    print("\033[32m" + '\n' + text + "\033[0m")

def print_red(text):
    """Prints text in red color."""
    print("\033[31m" + text + "\033[0m")

def askTypeOfBleed():
    """Asks the user what is the type of the bleed."""
    while True:
        response = input("\nWhat type of bleed does the file have?\nPlease enter 0 for none, 1 for regular, 2 for double, 3 for customized: ").strip().lower()
        if response in ('0', '1', '2', '3'):
            return response
        print_red("\nPlease enter 0, 1, 2 or 3.")

def customizedCrop(writer, reader):
    print("Please enter the amount to crop for each side:")

    def get_crop_value(side):
        while True:
            try:
                return int(input(f"Enter crop size for {side.upper()}: "))
            except ValueError:
                print(f"Invalid input. Please enter an integer value for {side.upper()}.")

    crop_left = get_crop_value("left")
    crop_right = get_crop_value("right")
    crop_top = get_crop_value("top")
    crop_bottom = get_crop_value("bottom")

    print("Please wait, cropping the document...")
    for page in reader.pages:
        page.mediabox.lower_left = (
            page.mediabox.lower_left[0] + crop_left,
            page.mediabox.lower_left[1] + crop_bottom,
        )
        page.mediabox.upper_right = (
            page.mediabox.upper_right[0] - crop_right,
            page.mediabox.upper_right[1] - crop_top,
        )
        writer.add_page(page)


def remove_bleed(file_path):
    try:
        reader = PdfReader(file_path)
        writer = PdfWriter()

        single_bleed_crop = {'left': 16, 'right': 15, 'top': 21, 'bottom': 22}
        double_bleed_crop = {'left': 36, 'right': 35, 'top': 35, 'bottom': 35}
        
        bleedType = askTypeOfBleed()
            
        if bleedType == '0':
            print("Please wait, processing...")
            for page in reader.pages:
                writer.add_page(page)
        
        elif bleedType in ('1', '2'):
            print("Please wait, removing bleed...")
            crop_values = double_bleed_crop if bleedType == '2' else single_bleed_crop

            for page in reader.pages:
                page.mediabox.lower_left = (
                    page.mediabox.lower_left[0] + crop_values['left'],
                    page.mediabox.lower_left[1] + crop_values['bottom'],
                )
                page.mediabox.upper_right = (
                    page.mediabox.upper_right[0] - crop_values['right'],
                    page.mediabox.upper_right[1] - crop_values['top'],
                )
                writer.add_page(page)
        
        elif bleedType == '3':
            customizedCrop(writer, reader)

        base_name, ext = os.path.splitext(file_path)
        new_file_name = base_name + "_no_bleed_box.pdf"

        with open(new_file_name, "wb") as output_pdf:
            writer.write(output_pdf)

        print_green(f"Finished bleed box removal")
        # subprocess.Popen([r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe", new_file_name])
        return new_file_name

    except Exception as e:
        print(f"Error: An error occurred - {e}")

def main():
    print("PDF Bleed Remover")
    print("=================")
    
    while True:
        file_path = input("\nDrop the PDF file here or enter the path of the PDF file (or type 'exit' to quit): ").strip().replace('"','')
        
        if file_path.lower() == 'exit':
            print("Exiting the program.")
            break
        elif os.path.isfile(file_path) and file_path.lower().endswith('.pdf'):
            print(f"Success: Processed PDF saved as {remove_bleed(file_path)}")
            break
        else:
            print_red("\nInvalid file. Please provide a valid PDF file.")

if __name__ == "__main__":
    main()
