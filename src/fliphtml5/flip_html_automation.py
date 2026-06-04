from src.logic.interface_controller import AppInterface

def verify_link_str_len(link_str, interface: AppInterface):
    while len(link_str) > 40: # Keep looping as long as the title is too long
        interface.print_info(f"  [WARNING] Title too long: {len(link_str)} chars.")
        link_str = interface.ask_string("String Too Long","Please enter a new book title (max 40 chars): ")

    return link_str


def fliphtml5_automation(pdf_folder_path, book_titles, row_index, interface:AppInterface):
    from src.fliphtml5 import API_Automation
    from src.logic.file_operations import validate_pdf_path
    from pathlib import Path

    display_title = book_titles.get('display_title')
    folder_title = book_titles.get('folder_name')
    fin_pdf_path = Path(pdf_folder_path) / f"{folder_title}_fin.pdf"
    is_path_valid, error_message = validate_pdf_path(str(fin_pdf_path))

    while not is_path_valid:
        interface.print_error(f"Invalid path at for {display_title}: {error_message}")
        fin_pdf_path = interface.ask_string("Fix File Path", f"Please enter the correct path for '{display_title}': ")
        is_path_valid, error_message = validate_pdf_path(fin_pdf_path)

    # Guaranteed to be a valid path if we break out of the loop above
    upload_success, upload_result = API_Automation.upload_file(fin_pdf_path, interface)

    if not upload_success:
        # This will now only handle actual SERVER/NETWORK upload errors, not bad paths!
        interface.print_error(f"Upload process failed: {upload_result}")
        return False

    uploaded_url = upload_result
    interface.print_success(f"File uploaded successfully to: {uploaded_url}")

    book_description = interface.ask_string("Enter Descrtiption", "Please enter the book description in English, if none - leave empty: ")

    link_str = verify_link_str_len(folder_title, interface)

    my_book_id = API_Automation.create_book(uploaded_url, display_title, book_description, link_str) # Create book using custom definitions and design profiles

    if my_book_id:
        interface.print_info(f"Book context successfully initialized. Target Book ID: {my_book_id}")

        # Wait for backend processing pools to output assets
        if interface.ask_yes_no("Set Password?", "Does the book needs to be protected by a password? "):
            from src.logic.excel_tools import get_password_from_excel
            password_as_a_list = get_password_from_excel(row_index,interface)

            if API_Automation.poll_conversion(my_book_id):
                # Lock book visibility down behind authorization passkeys
                API_Automation.set_book_privacy_with_password(my_book_id, password_as_a_list)
                interface.print_info(f"Set password {password_as_a_list}")

    return True