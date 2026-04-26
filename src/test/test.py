import os
import time
from pathlib import Path
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


def set_book_password(book_title, wait):
    from src.logic.excel_tools import get_password_from_excel
    print("[STEP] Opening 'Visibility Settings' dialog...")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Visibility Settings']"))).click()

    print(f"  [DEBUG] Fetching password for: {book_title}")
    password = get_password_from_excel(book_title)

    print("  [DEBUG] Selecting 'Private with Password' radio button...")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//label[.//span[text()='Private with Password']]"))).click()

    wait_and_type(wait, "input.el-h5_input", password)

    print("[STEP] Clicking 'Confirm' for Visibility Settings...")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'footer_button')]//span[text()='Confirm']"))).click()

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

        customize_book_link(book_title, wait, driver)
        print("[4] Book link customized.")

        from utils.input_output_tools import yes_or_no
        if yes_or_no("Does the book needs to be protected by password? "):
            print("[5] Proceeding to set password...")
            set_book_password(book_title, wait)
        else:
            print("[5] Skipping password protection.")

    except Exception as e:
        # This will now print the full error name which is helpful
        print(f"\n[ERROR] Process failed at: {type(e).__name__}")
        print(f"[ERROR] Details: {e}")

    finally:
        print("Process complete. Browser remains open due to 'detach' option.")


if __name__ == "__main__":
    # Your specific Institute network path
    test_path = r"R:\Documents\001אתר האינטרנט ופרויקטים דיגיטליים\הכנת כתבי עת לאתר\הכנת ספרים לאתר\קבצי ספרים מוכנים להעלאה לאמזון\studies_in_the_history_of_eretz_israel\flip"

    print("Starting fused flip upload script...")
    test_fliphtml5_with_profile(test_path)