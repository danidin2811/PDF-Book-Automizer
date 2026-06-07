from src.logic.interface_controller import AppInterface

def verify_link_str_len(link_str, interface: AppInterface):
    while len(link_str) > 40: # Keep looping as long as the title is too long
        interface.print_info(f"  [WARNING] Title too long: {len(link_str)} chars.")
        link_str = interface.ask_string("String Too Long","Please enter a new book title (max 40 chars): ")

    return link_str


def fliphtml5_automation(pdf_folder_path, book_titles, row_index, interface: AppInterface):
    from src.constants import BASE_DESIGN_TEMPLATE, FREE_HEBREW_DESIGN_TEMPLATE, ENGLISH_DESIGN_TEMPLATE
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
        interface.print_error(f"Upload process failed: {upload_result}")
        return False

    uploaded_url = upload_result
    interface.print_success(f"File uploaded successfully to: {uploaded_url}")

    book_description = interface.ask_string("Enter Description", "Please enter the book description in English, if none - leave empty: ")

    link_str = verify_link_str_len(folder_title, interface)

    chosen_config = BASE_DESIGN_TEMPLATE
    clean_choice = ""

    while True:
        user_design_template_choice = interface.ask_string(
            "Select Design Template",
            "Choose a design profile layout:\n"
            "Enter '1' for Paid Hebrew\n"
            "Enter '2' for Paid English\n"
            "Enter '3' for Free Hebrew\n\n"
            "Please enter 1, 2, or 3:"
        )

        if user_design_template_choice is None:
            if interface.ask_yes_no("Abort?", "Do you want to cancel the entire book creation process?"):
                interface.print_error("Process aborted by user.")
                return False
            continue

        clean_choice = user_design_template_choice.strip()

        if clean_choice == "1":
            chosen_config = BASE_DESIGN_TEMPLATE
            break

        elif clean_choice == "2":
            chosen_config = ENGLISH_DESIGN_TEMPLATE
            break

        elif clean_choice == "3":
            chosen_config = FREE_HEBREW_DESIGN_TEMPLATE
            break

        interface.print_error(f"Invalid entry: '{clean_choice}'. You must select 1, 2, or 3.")

    success = False
    my_book_id = None

    while not success:
        success, result = API_Automation.create_book(
            uploaded_url, display_title, book_description, link_str, design_config=chosen_config
        )

        if success:
            my_book_id = result
            break

        if isinstance(result, dict):
            error_code = result.get("code")

            if error_code == "LINK_ALREADY_EXISTS":
                link_str = interface.ask_string(
                    "Link Conflict",
                    f"The URL suffix '{link_str}' is already taken.\nPlease enter a unique alternative name:"
                )
                if not link_str:
                    interface.print_error("Process aborted by user.")
                    return False
                link_str = verify_link_str_len(link_str, interface)
                continue

            else:
                interface.print_error(f"FlipHTML5 rejected the configuration layout: {result}")
                return False
        else:
            interface.print_error(f"Critical System Fault: {result}")
            if not interface.ask_yes_no("Retry Connection?", "Would you like to try sending the request again?"):
                return False

    interface.print_success(f"Book context successfully initialized. Target Book ID: {my_book_id}")

    if clean_choice in ("1", "2"):
        from src.logic.excel_tools import get_password_from_excel
        password_as_a_list = get_password_from_excel(row_index, interface)

        if API_Automation.poll_conversion(my_book_id):
            API_Automation.set_book_privacy_with_password(my_book_id, password_as_a_list)
            interface.print_success(f"Set password {password_as_a_list}")

    return True