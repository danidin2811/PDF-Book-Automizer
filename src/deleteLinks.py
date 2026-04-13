import PyPDF2
import os

def validate_pdf_file(input_pdf_path):
    """Validates the given PDF file path."""
    
    if not os.path.isfile(input_pdf_path):
        return False, "Invalid file path. Please check and try again."
    if not input_pdf_path.lower().endswith('.pdf'):
        return False, "This file is not a PDF."
    return True, None

def deleteLinks(input_pdf_path):
    output_pdf_path = os.path.join(os.path.dirname(input_pdf_path), f" noLinks.pdf")
    input_reader = PyPDF2.PdfReader(input_pdf_path)
    
    # Initialize PDF writer
    pdf_writer = PyPDF2.PdfWriter()

    # Add pages to the writer
    for page in input_reader.pages:
        pdf_writer.add_page(page)

    pdf_writer.remove_links()
     
    # Write to the output file
    with open(output_pdf_path, "wb") as output_file:
        pdf_writer.write(output_file)
        

def main():
    # Drag and drop the input file
    input_pdf_path = input("Drag and drop the PDF file here and press Enter: ").strip('"')
    
    while True:
        is_valid, error = validate_pdf_file(input_pdf_path)
        if is_valid:
            break
        print_red(error)
        input_pdf_path = input("\nDrag and drop the PDF file here and press Enter: ").strip('"')
        
    deleteLinks(input_pdf_path)

if __name__ == "__main__":
    main()