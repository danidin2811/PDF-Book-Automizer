import os
import time
from pathlib import Path

from openpyxl.utils.protection import hash_password
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver import ActionChains


def wait_and_type(wait, selector, text, by=By.CSS_SELECTOR):
    print(f"  [DEBUG] Attempting to type into: {selector}")
    element = wait.until(EC.element_to_be_clickable((by, selector)))
    element.clear()
    element.send_keys(text)
    print(f"  [DEBUG] Successfully entered text: {text}")
    return element


def customize_book_link(book_title, wait, driver):
    print("[STEP] Opening settings popover...")
    try:
        settings_toggle = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.convert-setting-btn")))
        actions = ActionChains(driver)
        actions.move_to_element(settings_toggle).click().perform()

        wait.until(lambda d: d.find_element(By.CSS_SELECTOR, ".upload-setting-popper").is_displayed())
        print("  [DEBUG] Menu is now visible.")
    except Exception as e:
        print(f"  [ERROR] Failed to open menu: {e}")
        raise

    print("[STEP] Selecting 'Customize Book Link' from menu...")
    link_btn_xpath = "//div[contains(@class, 'convert-setting-item')]//span[text()='Customize Book Link']"
    try:
        book_link_button = wait.until(EC.visibility_of_element_located((By.XPATH, link_btn_xpath)))
        driver.execute_script("arguments[0].click();", book_link_button)
    except (TimeoutException, NoSuchElementException):
        raise

    if len(book_title) > 40:
        print(f"  [WARNING] Title too long: {len(book_title)} chars.")
        book_title = input("Please enter a new book title (max 40 chars): ")

    wait_and_type(wait, "input.el-input__inner", book_title)

    time.sleep(1)

    print("[STEP] Clicking 'Confirm' for Book Link...")

    # 2. BETTER SELECTOR: The 'confirm-btn' class is the most unique identifier here.
    # We use a broad XPath that looks for 'Confirm' text inside that specific class.
    confirm_xpath = "//div[contains(@class, 'confirm-btn') and contains(., 'Confirm')]"

    try:
        # 3. USE PRESENCE INSTEAD: Sometimes visibility is tricky with overlays.
        # We find it, then use JavaScript to force the click.
        confirm_btn = wait.until(EC.presence_of_element_located((By.XPATH, confirm_xpath)))

        # Scroll into view just in case
        driver.execute_script("arguments[0].scrollIntoView(true);", confirm_btn)

        # Force the click via JavaScript
        driver.execute_script("arguments[0].click();", confirm_btn)
        print("[DEBUG] Book link customization confirmed.")

        # 4. WAIT FOR DIALOG TO DISAPPEAR: To ensure we don't move to the next
        # step while the modal is still closing.
        wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".light_box_con")))

    except Exception as e:
        print(f"[ERROR] Could not confirm Book Link: {e}")
        # FALLBACK: Try clicking by class only if text fails
        try:
            fallback_btn = driver.find_element(By.CSS_SELECTOR, "div.confirm-btn")
            driver.execute_script("arguments[0].click();", fallback_btn)
            print("[DEBUG] Confirmed using fallback selector.")
        except:
            raise e


def set_book_password(book_title, wait, driver):
    from src.logic.excel_tools import get_password_from_excel

    print("[STEP] Opening settings popover for Visibility...")
    try:
        settings_toggle = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.convert-setting-btn")))
        actions = ActionChains(driver)
        actions.move_to_element(settings_toggle).click().perform()
        wait.until(lambda d: d.find_element(By.CSS_SELECTOR, ".upload-setting-popper").is_displayed())
    except Exception as e:
        print(f"  [ERROR] Failed to open menu for Visibility: {e}")
        raise

    print("[STEP] Selecting 'Visibility Settings' from menu...")
    visibility_btn_xpath = "//div[contains(@class, 'convert-setting-item')]//span[text()='Visibility Settings']"
    visibility_btn = wait.until(EC.visibility_of_element_located((By.XPATH, visibility_btn_xpath)))
    driver.execute_script("arguments[0].click();", visibility_btn)

    print(f"  [DEBUG] Fetching password for: {book_title}")
    password = get_password_from_excel(book_title)

    print("  [DEBUG] Selecting 'Private with Password' radio button...")
    # The radio button is inside an El-Radio component. We target the span containing the text.
    radio_xpath = "//label[.//span[text()='Private with Password']]"
    radio_btn = wait.until(EC.presence_of_element_located((By.XPATH, radio_xpath)))
    driver.execute_script("arguments[0].click();", radio_btn)

    # Wait specifically for the input field to become visible after the click
    print("  [DEBUG] Waiting for password input field...")
    password_input = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input.h5_input")))
    password_input.clear()
    password_input.send_keys(password)

    # FlipHTML5 often requires clicking 'Add' or pressing Enter for the password to register
    try:
        add_button = driver.find_element(By.XPATH, "//button[contains(., 'Add')]")
        driver.execute_script("arguments[0].click();", add_button)
        print("  [DEBUG] Clicked 'Add' button.")
    except:
        print("  [DEBUG] 'Add' button not found or not needed.")

    time.sleep(1)

    print("[STEP] Clicking the footer 'Confirm' button...")
    try:
        # 1. Wait for the element to be present
        confirm_btn = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[contains(@class, 'footer_button')]//span[text()='Confirm']")))

        # 2. Scroll to it just in case it's below the fold
        driver.execute_script("arguments[0].scrollIntoView(true);", confirm_btn)

        # 3. Force the click via JavaScript
        driver.execute_script("arguments[0].click();", confirm_btn)
        print("  [DEBUG] Confirm clicked successfully.")

    except Exception as e:
        print(f"  [ERROR] Failed to click footer Confirm: {e}")


def upload_fin_pdf(pdf_folder_path):
    import pyperclip
    from utils.input_output_tools import wait_for_ready_signal

    book_title = Path(pdf_folder_path).parent.name
    fin_pdf_path = Path(pdf_folder_path) / f"{book_title}_fin.pdf"
    pyperclip.copy(fin_pdf_path)

    try:
        print("[STEP] Triggering the Windows Open dialog...")
        wait_for_ready_signal("MANUAL ACTION REQUIRED:\n"
                              "1. Click the 'Upload Files' button\n"
                              "2. Press Ctrl+V in the File name field at the bottom to paste the path"
                              "3. Press Enter to start the file upload\n")

        return True

    except Exception as e:
        print(f"  [ERROR] Failed to trigger dialog: {e}")
        return False


def apply_design_settings(driver, wait, has_password):
    from utils.input_output_tools import yes_or_no

    # 1. Change the string in the input field
    # Assuming this is a title or alias field appearing after conversion
    try:
        print("[STEP] Updating text field...")
        input_field = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input.h5_input")))
        input_field.clear()
        input_field.send_keys("Updated Book Title")  # Replace with your variable
    except Exception as e:
        print(f"  [ERROR] Could not update input field: {e}")

    # 2. Click 'Customize' in the left menu
    try:
        print("[STEP] Clicking 'Customize' menu...")
        customize_btn = wait.until(EC.element_to_be_clickable((By.ID, "left_menu_customize")))
        driver.execute_script("arguments[0].click();", customize_btn)
    except Exception as e:
        print(f"  [ERROR] Customize menu failed: {e}")

    # 3. Click 'My Designs'
    try:
        print("[STEP] Opening 'My Designs'...")
        my_designs_xpath = "//li[@title='My Designs']//div[contains(@class, 'h5_setting_nav_button')]"
        my_designs_btn = wait.until(EC.element_to_be_clickable((By.XPATH, my_designs_xpath)))
        driver.execute_script("arguments[0].click();", my_designs_btn)
        time.sleep(2)  # Wait for templates to render
    except Exception as e:
        print(f"  [ERROR] My Designs failed: {e}")

    # 4. Handle Hebrew vs English Logic
    is_hebrew = yes_or_no("Is the book in Hebrew?")

    # Identify target design based on text label in 'itemName'
    if is_hebrew and has_password:
        target_label = "לא להורדה ולשיתוף"
    elif not is_hebrew and has_password:
        target_label = "English design"
    else:
        # Assuming "פתוח לשיתוף" is the 'no password' default
        target_label = "פתוח לשיתוף"

    print(f"[STEP] Applying design template: {target_label}")

    try:
        # Locate the design container that contains the specific text label
        container_xpath = f"//div[contains(@class, 'designItem')][.//div[text()='{target_label}']]"
        design_container = wait.until(EC.presence_of_element_located((By.XPATH, container_xpath)))

        # Hover is often required to make the 'Apply' button appear in the DOM
        actions = ActionChains(driver)
        actions.move_to_element(design_container).perform()

        # Find the 'Apply' button inside that specific container
        apply_btn = design_container.find_element(By.XPATH, ".//div[contains(@class, 'apply_btn')]")
        driver.execute_script("arguments[0].click();", apply_btn)
        print(f"  [DEBUG] '{target_label}' applied successfully.")

    except Exception as e:
        print(f"  [ERROR] Failed to apply template '{target_label}': {e}")

def wait_for_conversion_and_continue(driver, wait, has_password, long_timeout=600):
    print(f"[STEP] Waiting for conversion to complete (Timeout: {long_timeout}s)...")

    # Create a dedicated wait object for the conversion process
    conversion_wait = WebDriverWait(driver, long_timeout)

    try:
        # The URL changes to 'bookinfo' only after conversion hits 100%
        conversion_wait.until(EC.url_contains("bookinfo"))
        print(f"  [DEBUG] Conversion finished. New URL: {driver.current_url}")
        apply_design_settings(driver,wait,has_password)

    except TimeoutException:
        print(f"  [ERROR] Conversion exceeded {long_timeout} seconds.")
        return False

    # 2. Allow the Editor UI (sidebar and preview) to fully load
    time.sleep(5)

    try:
        print("[STEP] Navigating to 'Customize' settings...")
        customize_xpath = "//div[contains(@class, 'menu-item')]//span[text()='Customize']"
        customize_btn = conversion_wait.until(EC.element_to_be_clickable((By.XPATH, customize_xpath)))
        driver.execute_script("arguments[0].click();", customize_btn)
        print("  [DEBUG] Customize menu opened.")
    except Exception as e:
        print(f"  [ERROR] Could not click 'Customize': {e}")

    return True

def test_fliphtml5_with_profile(pdf_folder_path):
    chrome_options = Options()

    # 1. NEW PROFILE LOGIC: Create a dedicated automation folder
    # This prevents the 'Chrome not reachable' error caused by profile locks
    automation_data = os.path.join(os.environ['LOCALAPPDATA'], r"Google\Chrome\AutomationData")

    if not os.path.exists(automation_data):
        os.makedirs(automation_data)
        print(f"Created new automation profile at: {automation_data}")

    chrome_options.add_argument(f"--user-data-dir={automation_data}")

    # 2. STABILITY ARGUMENTS:
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_experimental_option("detach", True)

    try:
        print("Initializing driver...")
        driver = webdriver.Chrome(options=chrome_options)

        dashboard_url = "https://fliphtml5.com/dashboard/publications/folder/7398072?lang=en"
        driver.get(dashboard_url)

        # CHECK FOR LOGIN:
        if "login" in driver.current_url.lower():
            print("\nACTION REQUIRED: Please log in manually in the opened browser window.")
            print("The script will wait up to 2 minutes for you to reach the dashboard...")

        wait = WebDriverWait(driver, 15)  # Standard wait after login check

        print("[1] Waiting for dashboard upload area...")
        upload_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.upload_area")))
        upload_btn.click()

        print("[2] Waiting for URL to change to upload/edit mode...")
        wait.until(EC.url_contains("edit-book/upload"))
        print(f"  [DEBUG] Current URL: {driver.current_url}")

        book_title = Path(pdf_folder_path).parent.name
        print(f"[3] Processing book title: {book_title}")

        # customize_book_link(book_title, wait, driver)
        print("[4] Book link customized.")

        from utils.input_output_tools import yes_or_no
        has_password = yes_or_no("Does the book needs to be protected by password? ")
        if has_password:
            print("[5] Proceeding to set password...")
            set_book_password(book_title, wait, driver)

        else:
            print("[5] Skipping password protection.")

        if upload_fin_pdf(pdf_folder_path):
            wait_for_conversion_and_continue(driver, wait, has_password)

    except Exception as e:
        # This will now print the full error name which is helpful
        print(f"\n[ERROR] Process failed at: {type(e).__name__}")
        print(f"[ERROR] Details: {e}")

    finally:
        print("Process complete. Browser remains open due to 'detach' option.")


def run_isolated_test():
    chrome_options = Options()
    automation_data = os.path.join(os.environ['LOCALAPPDATA'], r"Google\Chrome\AutomationData")
    chrome_options.add_argument(f"--user-data-dir={automation_data}")
    chrome_options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

    try:
        url = "https://fliphtml5.com/edit-book/38783277/design?lang=en"
        driver.get(url)

        print("Waiting for page to load. Please log in if prompted...")
        time.sleep(5)

        # Testing the function with has_password=True as an example
        apply_design_settings(driver, wait, has_password=True)

    except Exception as e:
        print(f"Test Runner Error: {e}")


if __name__ == "__main__":
    run_isolated_test()