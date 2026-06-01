from src.logic.interface_controller import AppInterface

def verify_link_str_len(link_str, interface: AppInterface):
    while len(link_str) > 40: # Keep looping as long as the title is too long
        interface.print_info(f"  [WARNING] Title too long: {len(link_str)} chars.")
        link_str = interface.ask_string("String Too Long","Please enter a new book title (max 40 chars): ")

    return link_str


def fliphtml5_automation(pdf_folder_path, display_title, row_index, interface:AppInterface):
    from src.fliphtml5 import API_Automation

    uploaded_url = API_Automation.upload_file(pdf_folder_path)

    if uploaded_url:
        interface.print_info(f"File uploaded successfully to: {uploaded_url}")

        book_description = interface.ask_string("Enter Descrtiption", "Please enter the book description in English, if none - leave empty: ")

        link_str = verify_link_str_len(display_title, interface)

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