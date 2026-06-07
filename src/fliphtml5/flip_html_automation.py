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

    success = False
    my_book_id = None

    # Run an explicit control loop until the task succeeds, or the user cancels
    while not success:
        # Trigger the call and unpack the stateful tuple
        success, result = API_Automation.create_book(uploaded_url, display_title, book_description, link_str)

        if success:
            my_book_id = result  # Now guaranteed to be a valid bookId string
            break

        # --- ERROR HANDLING LAYER ---
        # Case A: An API specification error occurred (result is a dictionary)
        if isinstance(result, dict):
            error_code = result.get("code")

            if error_code == "LINK_ALREADY_EXISTS":
                # Give the user a clear prompt to update the link
                link_str = interface.ask_string(
                    "Link Conflict",
                    f"The URL suffix '{link_str}' is already taken.\nPlease enter a unique alternative name:"
                )
                if not link_str:  # User canceled or entered empty string
                    interface.print_error("Process aborted by user.")
                    return False
                link_str = verify_link_str_len(link_str, interface)
                continue  # Re-evaluate loop with the newly typed link_str

            else:
                # Handle any other backend format validation errors
                interface.print_error(f"FlipHTML5 rejected the configuration layout: {result}")
                return False

        # Case B: A critical system/network failure occurred (result is a string error message)
        else:
            interface.print_error(f"Critical System Fault: {result}")
            if not interface.ask_yes_no("Retry Connection?", "Would you like to try sending the request again?"):
                return False

    # =========================================================
    # GURANTEED SUCCESS ZONE
    # =========================================================
    interface.print_success(f"Book context successfully initialized. Target Book ID: {my_book_id}")

    # Wait for backend processing pools to output assets
    if interface.ask_yes_no("Set Password?", "Does the book need to be protected by a password? "):
        from src.logic.excel_tools import get_password_from_excel
        password_as_a_list = get_password_from_excel(row_index, interface)

        if API_Automation.poll_conversion(my_book_id):
            API_Automation.set_book_privacy_with_password(my_book_id, password_as_a_list)
            interface.print_success(f"Set password {password_as_a_list}")

    my_book_id = API_Automation.create_book(uploaded_url, display_title, book_description, link_str) # Create book using custom definitions and design profiles

    while not isinstance(my_book_id, list):
        if isinstance(my_book_id, dict):
            error_code = my_book_id.get("code")
            error_msg = my_book_id.get("msg")

            if error_code == "'LINK_ALREADY_EXISTS'" and error_msg == "'LINK_ALREADY_EXISTS'":
                link_str = interface.ask_string(f"A book with the link {link_str} already exists", "Please enter a new string for the URL of the book")
                my_book_id = API_Automation.create_book(uploaded_url, display_title, book_description, link_str)

        if isinstance(my_book_id, str):
            interface.print_error(f"There was an error creating the book: {my_book_id}")

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