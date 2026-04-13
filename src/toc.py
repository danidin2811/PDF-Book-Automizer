from tqdm import tqdm
from removeBleedBox import print_green
import re
import PyPDF2
import os
import subprocess
import string

from deleteLinks import deleteLinks
from tocCorrection import correct_line

def char_to_int(char):
    """Converts a character (e.g., 'a', 'b', ...) to an integer (e.g., 1, 2, ...)."""
    return ord(char.lower()) - ord('a') + 1

def debugPrint(text):
    print(text)

def clean_leading_dots(s: str) -> str:
    # Match cases where the string starts with two dots, allowing spaces around them
    match = re.match(r'^\s*\.\s*\.', s)
    if match:
        # Remove all leading dots and spaces until a non-dot, non-space character appears
        s = re.sub(r'^[.\s]+', '', s)
    return s

def reverseTitle(title):
    if '(' in title or ')' in title:
      words = re.findall(r'\([^\)]*\)|\S+', title.strip())

      # Process each word to reverse internal parentheses content if needed
      for i, word in enumerate(words):
          if word.startswith('(') and word.endswith(')'):
              inside = word[1:-1].strip().split()
              inside.reverse()
              words[i] = '(' + ' '.join(inside) + ')'

      # Reverse the whole word list
      words.reverse()
      
    else:
    	words = title.strip().split()
    	words.reverse()

    return ' '.join(words)

def extract_toc(pdf_reader, toc_start_page, toc_end_page, total_pages, offsets, reverseOrder, second_volume_marker=None):
    """
    Extracts TOC entries from the specified pages.
    Handles single or two-volume books with different offsets.
    """
    toc = []
    current_offset = offsets[0]  # Default to the first volume's offset
    marker_found = False  # Track if the marker for the second volume is found
    
    for page_num in range(toc_start_page - 1, toc_end_page):
        text = pdf_reader.pages[page_num].extract_text()
        if text:  # Check if the page has text
            for line in text.splitlines():  # Split the text into lines
                debugPrint(line)
                if second_volume_marker and not marker_found:
                    if second_volume_marker in line:
                        current_offset = offsets[1]
                        marker_found = True
                        print(f"Second volume marker '{second_volume_marker}' detected.")


        # First regex pattern: page number first
        matches = re.findall(r"^\s*(\d+)\s+(.*?)$", text, re.MULTILINE)

        # If no matches, try the second regex pattern: page number last
        if not matches:
            matches = re.findall(r"^(.*?)\s+(\d+)\s*$", text, re.MULTILINE)
            
        for match in matches:
            if len(match) == 2:  # Ensure match has both title and page
                try:
                    if match[0].isdigit():  # First pattern (page number first)
                        page_number = int(match[0]) + current_offset
                        title = match[1].strip()
                    else:  # Second pattern (title first)
                        page_number = int(match[1]) + current_offset
                        title = match[0].strip()
                        
                    debugPrint(f"{title}")

                    # Process the title for better readability
                    
                    title = clean_leading_dots(title)
                    title = correct_line(title)
                    
                    if "|" in title:
                        parts = title.split("|")
                        new_title = parts[1].strip() + " | " + parts[0].strip()
                    elif "-" in title:
                        parts = title.split("-")
                        new_title = parts[1].strip() + "-" + parts[0].strip()
                    elif " - " in title:
                        parts = title.split(" - ")
                        new_title = parts[1].strip() + " - " + parts[0].strip()
                    else:
                        new_title = title.strip()
                        new_title = new_title.replace('  ', ' ')
                        new_title = new_title.replace(' ', ' ')

                    # Append valid TOC entries
                    if page_number < total_pages:
                        if reverseOrder:
                            new_title = reverseTitle(new_title)
                        debugPrint(new_title)

                        toc.append((new_title, page_number))
                except ValueError:
                    print(f"Skipping invalid entry: {match}")

    return toc


def askIfBeforeOrAfterFrontCover():
    while True:
        user_input = input("Is the pdf file before or after adding the front cover?\nPlease enter either a for after or b for before: ").strip().lower()
        if user_input in ('a', 'b'):
            return 1 if user_input == 'a' else -1
            
        else:
            print_red("\nPlease enter either a for after or b for before")


def askOffset():
    """
    Prompts the user to specify an offset for book page numbering.
    Returns:
        int: The offset entered by the user. Always an integer value (can be positive, negative, or zero).
    """
    while True:
        try:
            user_input = int(input(
                "\nIs there an offset in the book's pages?\n"
                "Please enter the amount of offset pages as a number (positive, negative, or 0 for none): "
            ).strip())
            print(f"You entered offset: {user_input}")
            confirm = input("Is this correct? (y/n): ").strip().lower()
            if confirm == 'y':
                return user_input
            else:
                print("Re-enter the offset.")
        except ValueError:
            print_red("\nInvalid input. Please enter a valid integer.\n")


def add_bookmarks(pdf_writer, toc, pagesOffset):#, bookmarks_list):
    """
    Adds bookmarks to the PDF based on the TOC.
    """
    parent_bookmarks = {}
    
    coverOffset = askIfBeforeOrAfterFrontCover()
    print("Adding bookmarks, please wait...")
    
    for bookmark in tqdm(toc, total=len(toc), desc="Processing bookmarks"):
        title, page_num = bookmark
        page_index = page_num + coverOffset #+ pagesOffset  # PyPDF2 uses zero-based indexing
        # print(f"{title} at page {page_index}")
        parent_bookmarks[1] = pdf_writer.add_outline_item(title, page_index)
        # if level == 1:
            # parent_bookmarks[level] = pdf_writer.add_outline_item(title, page_index)
        # else:
            # parent = parent_bookmarks.get(level - 1)
            # if parent:
                # pdf_writer.add_outline_item(title, page_index, parent=parent)
    
    print('\n')


def print_red(text):
    """Prints text in red color."""
    print("\033[31m" + text + "\033[0m")


def get_pdf_page_count(file_path):
    """Returns the total number of pages in a PDF file."""
    
    try:
        with open(file_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            return len(pdf_reader.pages)
    except Exception as e:
        print_red(f"Error reading PDF file: {e}")
        return None
        

def readBookmarksList():
    """
    Reads a structured list of bookmark levels and their frequencies from user input.

    The user is prompted to enter multiple lines in the format:
        X Y
    Where:
        - X is a single alphabetical character representing the bookmark level 
          ('a' maps to 1, 'b' to 2, etc.).
        - Y is an integer representing how many times this level of bookmarks 
          should be repeated.

    Input ends when the user enters a single '-' character.

    Returns:
        A list of integers, where each integer represents a bookmark level. 
        The list contains each level repeated according to the specified frequency.
    """
    
    inputBookmarksList = []

    print("Enter the bookmarks structure:")
    print("Each line should be in the format: X Y")
    print("    - X is a single letter (e.g., 'a', 'b', etc.) representing the level of the bookmark.")
    print("    - Y is an integer specifying how many bookmarks of this level should be added.")
    print("    - Type '-' when you are done entering your data.\n")

    while True:
        line = input("Enter a line (or '-' to finish): ").strip()  # Read input and strip whitespace
        if line == '-':  # Check for the end of input
            break
        
        # print(line)
        # Split the input line into level and count
        try:
            level_char, count_str = map(str.strip, line.split())
            level = char_to_int(level_char)
            count = int(count_str)
            inputBookmarksList.extend([level] * count)

        except ValueError as e:
            print_red(f"Error parsing line '{line}': {e}. Please ensure it's in the format 'X Y'.")

    return inputBookmarksList


def validate_pdf_file(input_pdf_path):
    """Validates the given PDF file path."""
    
    if not os.path.isfile(input_pdf_path):
        return False, "Invalid file path. Please check and try again."
    if not input_pdf_path.lower().endswith('.pdf'):
        return False, "This file is not a PDF."
    return True, None


def validate_toc_pages(total_pages):
    """
    Prompts the user to enter valid TOC start and end pages.
    
    Args:
        total_pages (int): The total number of pages in the PDF.

    Returns:
        tuple: A tuple containing the valid TOC start and end pages (start_page, end_page).
    """
    while True:
        try:
            toc_start_page = int(input("Enter the starting page number of the TOC: "))
            if 1 <= toc_start_page <= total_pages:
                break
            print_red(f"Page number must be between 1 and {total_pages}.")
        except ValueError:
            print_red("Please enter a valid number.")

    while True:
        try:
            toc_end_page = int(input("Enter the ending page number of the TOC: "))
            if toc_start_page <= toc_end_page <= total_pages:
                print(f"You entered start page: {toc_start_page}, end page: {toc_end_page}")
                confirm = input("Are these correct? (y/n): ").strip().lower()
                if confirm == 'y':
                    return toc_start_page, toc_end_page
                else:
                    print("Re-enter the TOC pages.")
                    return validate_toc_pages(total_pages)
            print_red(f"Page number must be between {toc_start_page} and {total_pages}.")
        except ValueError:
            print_red("Please enter a valid number.")


def askToRemove(toc, diff):
    print("More TOC entries were detected than bookmarks provided.")
    input("Please press any key")
    print("These are the TOC entries that were detected:")
    input("Please press any key")
    
    for i, e in enumerate(toc, start=1):
        print(f"{i}. {e}")
    
    print("\nPlease choose which TOC entries to remove (by number).")
    
    while diff > 0 and toc:
        print("Enter the number of the entry to remove:")
        try:
            a = int(input().strip()) - 1  # Convert to zero-based index
            if 0 <= a < len(toc):
                removed = toc.pop(a)
                print(f"Removed: {removed}")
                diff -= 1
            else:
                print("Invalid number. Please choose a valid number.")
        except ValueError:
            print("Invalid input. Please enter a number.")
    
    if diff > 0:
        print(f"Could not remove all entries. {diff} entries remain.")
    elif not toc:
        print("All TOC entries have been removed.")
    else:
        print("Entries successfully removed.")


def processBookForBookmarks(input_pdf_path, total_pages, folderName):
    print("")
    output_pdf_path = os.path.join(os.path.dirname(input_pdf_path), f"{folderName}_toc.pdf")
    
    with open(input_pdf_path, "rb") as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        pdf_writer = PyPDF2.PdfWriter()

        for page in pdf_reader.pages:
            pdf_writer.add_page(page)
    
    try:
        while True:
            user_input = input("Does the book have two volumes? (y/n): ").strip().lower()
            if user_input in {"y", "n"}:
                is_two_volumes = user_input == 'y'
                break
            print_red("\nInvalid input. Please enter either 'y' for yes or 'n' for no.\n")

        if is_two_volumes:
            print("\nFIRST VOLUME")
            offset1 = askOffset()
            print("\nSECOND VOLUME")
            offset2 = askOffset()
            second_volume_marker = input(
                "\nWhat text indicates the start of the second volume's TOC? ").strip()
            offsets = (offset1, offset2)
        else:
            offsets = (askOffset(),)
            second_volume_marker = None

        while True:
            user_input = input("Do you need to reverse the word order of the title? (y/n): ").strip().lower()
            if user_input in {"y", "n"}:
                reverseOrder = user_input == 'y'
                break
            print_red("\nInvalid input. Please enter either 'y' for yes or 'n' for no.\n")
        
        while True:
            toc_start_page, toc_end_page = validate_toc_pages(total_pages)
            toc = extract_toc(
                pdf_reader, toc_start_page, toc_end_page, total_pages, offsets, reverseOrder, second_volume_marker)
            
            if not toc:
                print_red("No TOC detected.\nPerhaps there was some mistake in the input\n")
                while True:
                    choice = input("Would you like to try again? (y/n): ").strip().lower()
                    if choice in {"y", "n"}:
                        break
                    print_red("\nInvalid input. Please enter either 'y' for yes or 'n' for no.\n")
                
                if choice == "n":
                    print("Exiting without processing TOC.")
                    return
                else:
                    print()
            else:
                print_green("TOC successfully detected")
                break

        add_bookmarks(pdf_writer, toc, offsets[0])  # First offset is used for the base logic

        print("Please wait, writing the new PDF file...")
        with open(output_pdf_path, "wb") as output_pdf_file:
            pdf_writer.write(output_pdf_file)

        print_green(f"Bookmarks added successfully!")
        
        return output_pdf_path

    except Exception as e:
        print_red(f"An error occurred: {e}")
        return None
        
def constructToc(bookmarks, pdf_reader, level=1):
    """
    Recursively extracts the table of contents from the bookmarks.

    Args:
        bookmarks (list): The bookmarks to process.
        pdf_reader (PyPDF2.PdfReader): The PDF reader object.
        level (int): The current bookmark level (default is 1).

    Returns:
        list: A list of tuples with (title, page_number, level).
    """
    extractedToc = []

    for item in bookmarks:
        try:
            if isinstance(item, list):
                # Recursively process nested bookmarks and extend the result
                extractedToc.extend(constructToc(item, pdf_reader, level + 1))
            else:
                # Extract title
                title = getattr(item, 'title', None)
                if not title:
                    print_red("Bookmark without a title encountered. Skipping...")
                    continue

                # Resolve page number
                page_number = None
                if hasattr(item, 'page') and item.page:
                    try:
                        page_number = pdf_reader.get_page_number(item.page) + 1  # Convert to 1-based index
                    except Exception as e:
                        print_red(f"Error resolving page number for {title}: {e}")
                        continue

                # Append the current item to the extracted TOC
                extractedToc.append((title, page_number, level))

        except Exception as e:
            print_red(f"Error processing bookmark: {e}")

    return extractedToc


def copyToc(input_pdf_path, folderName, originalPdfWithToc=None):
    """
    Copies the table of contents (TOC) from the original PDF to a new PDF.
    
    Args:
        input_pdf_path (str): Path to the input PDF.
        folderName (str): Name of the folder where the output will be saved.
        originalPdfWithToc (str, optional): Path to the original PDF with TOC. Defaults to None.
    """
    print("")
    output_pdf_path = os.path.join(os.path.dirname(input_pdf_path), f"{folderName}_fin.pdf")

    def process_pdf(pdf_reader, toc_reader=None):
        """
        Handles TOC construction and writing the new PDF.
        """
        
        # Extract and construct TOC
        toc = toc_reader.outline if toc_reader else pdf_reader.outline

        if toc:
            ready_toc = constructToc(toc, toc_reader or pdf_reader)

            
        else:
            print_red("Couldn't read bookmarks.")
            return

        # Initialize PDF writer
        pdf_writer = PyPDF2.PdfWriter()

        # Add pages to the writer
        for page in pdf_reader.pages:
            pdf_writer.add_page(page)

        pdf_writer.remove_links()
        
        print("Copying bookmarks, please wait...")
        parent_bookmarks = {}

        # Add bookmarks
        for bookmark in ready_toc:
            try:
                title, page_num, level = bookmark
                page_num -= 1  # Adjust page number for zero-based indexing

                if level == 1:
                    parent_bookmarks[1] = pdf_writer.add_outline_item(title, page_num)
                else:
                    parent = parent_bookmarks.get(level - 1)
                    if parent:
                        new_bookmark = pdf_writer.add_outline_item(title, page_num, parent=parent)
                        parent_bookmarks[level] = new_bookmark
            except Exception as e:
                print_red(f"Error processing bookmark {bookmark}: {e}")
            
        # Write to the output file
        with open(output_pdf_path, "wb") as output_file:
            pdf_writer.write(output_file)
# ---------------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------
    try:
        # Process the PDF files
        if originalPdfWithToc:
            with open(originalPdfWithToc, "rb") as toc_file, open(input_pdf_path, "rb") as input_file:
                toc_reader = PyPDF2.PdfReader(toc_file)
                input_reader = PyPDF2.PdfReader(input_file)
                process_pdf(input_reader, toc_reader)
        else:
            with open(input_pdf_path, "rb") as input_file:
                input_reader = PyPDF2.PdfReader(input_file)
                process_pdf(input_reader)
        
        folder = os.path.dirname(input_pdf_path)
        con_file_path = None
        for file in os.listdir(folder):
            if file.endswith("_con.pdf"):
                con_file_path = os.path.join(folder, file)
                break

        if con_file_path:
            # Store the original _con.pdf file name in a variable
            original_con_file_name = os.path.basename(con_file_path)
            
            # Delete links from the '_con' file
            no_links_file = deleteLinks(con_file_path)

            # Delete the '_con' file
            os.remove(con_file_path)

            # Rename the file without links to the original _con.pdf file name
            renamed_con_file_path = os.path.join(folder, original_con_file_name)
            
            if no_links_file:  # Check if no_links_file is NOT None
                os.rename(no_links_file, renamed_con_file_path)
                
        else:
            print_red("Couldn't find _con file to delete links.")
            
        print_green(f"Bookmarks added successfully!")
        subprocess.Popen([r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe", output_pdf_path])
        return output_pdf_path
        
    except Exception as e:
        print_red(f"An error occurred in copy toc: {e}")
        
        
def processBookNoToc(input_pdf_path, folderName):
    total_pages = get_pdf_page_count(input_pdf_path)
    if total_pages is None:
        return
    
    try:
        output_pdf_path = processBookForBookmarks(input_pdf_path, total_pages, folderName)
        return output_pdf_path  # Return the path if successful
        
    except Exception as e:
        print_red(f"An error occurred during TOC processing: {e}")
        return None  # Return None if an exception occurs
        
# Main function
def main():
    # Drag and drop the input file
    input_pdf_path = input("Drag and drop the PDF file here and press Enter: ").strip('"')
    
    while True:
        is_valid, error = validate_pdf_file(input_pdf_path)
        if is_valid:
            break
        print_red(error)
        input_pdf_path = input("\nDrag and drop the PDF file here and press Enter: ").strip('"')
    
    while True:
        tocExists = input("Does the pdf file already has bookmarks? (y/n): ").strip().lower()
        if tocExists in {"y", "n"}:
            break
        print_red("\nInvalid input. Please enter either 'y' for yes or 'n' for no.\n")
    
    if tocExists == 'y':
        copyToc(input_pdf_path, os.path.basename(os.path.dirname(input_pdf_path)))
        
    elif tocExists == 'n':
        processBookNoToc(input_pdf_path, os.path.basename(os.path.dirname(input_pdf_path)))

if __name__ == "__main__":
    main()