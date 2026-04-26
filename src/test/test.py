import os
import time
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


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

        # Increased wait time for this manual step
        wait = WebDriverWait(driver, 120)

        # This will now wait until you finish logging in and the dashboard loads
        print("Waiting for dashboard to load...")
        upload_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.upload_area")))

        print("Dashboard detected. Proceeding with click...")
        upload_btn.click()

        # 5. TRANSITION TO UPLOAD PAGE:
        print("Waiting for upload page to load...")
        wait.until(EC.url_contains("edit-book/upload"))

        book_link_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Customize Book Link']")))
        book_link_button.click()

        book_title = Path(pdf_folder_path).parent.name

        if len(book_title) > 40:
            print(f"Warning: Title is {len(book_title)} characters. It will be truncated to 40.")
            book_title = input("Please enter a new book title shorter than 40 chars: ")


        input_field = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input.el-input__inner")))
        input_field.clear() # Clear the field first in case there is default text
        input_field.send_keys(book_title) # Enter the string

        confirm_selector = "div.confirm-btn.h5_setting_white_text_blue_button"
        confirm_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, confirm_selector)))
        confirm_button.click()

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        print("Process complete. Browser remains open due to 'detach' option.")


if __name__ == "__main__":
    # Your specific Institute network path
    test_path = r"R:\Documents\001אתר האינטרנט ופרויקטים דיגיטליים\הכנת כתבי עת לאתר\הכנת ספרים לאתר\קבצי ספרים מוכנים להעלאה לאמזון\studies_in_the_history_of_eretz_israel\flip"

    print("Starting fused flip upload script...")
    test_fliphtml5_with_profile(test_path)