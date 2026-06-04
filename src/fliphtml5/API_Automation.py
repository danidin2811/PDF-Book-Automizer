import os
import json
import time
import base64
import hmac
import hashlib
from datetime import datetime, timezone
from pathlib import Path

import requests
from dotenv import load_dotenv

from src.logic.file_operations import validate_pdf_path
from src.logic.interface_controller import AppInterface

load_dotenv()

# ==========================================
# 1. HEADER GENERATION & SIGNING HELPER
# ==========================================
def generate_fliphtml5_headers(path, query_params=None):
    ACCESS_KEY_ID = os.getenv("Access_ID")
    ACCESS_KEY_SECRET = os.getenv("Access_Key")

    if not ACCESS_KEY_ID or not ACCESS_KEY_SECRET:
        raise ValueError("Missing FlipHTML5 credentials! Ensure your .env file is set up correctly.")

    now = datetime.now(timezone.utc)
    date_str = now.strftime("%a, %d %b %Y %H:%M:%S GMT")

    resource = path
    if query_params:
        sorted_keys = sorted(query_params.keys())
        query_string = "&".join([f"{key}={query_params[key]}" for key in sorted_keys])
        resource += f"?{query_string}"

    sign_string = f"{date_str}\n{resource}"
    key_bytes = ACCESS_KEY_SECRET.encode('utf-8')
    msg_bytes = sign_string.encode('utf-8')

    hashed = hmac.new(key_bytes, msg_bytes, hashlib.sha1).digest()
    signature = base64.b64encode(hashed).decode('utf-8')

    return {
        "Date": date_str,
        "x-yzw-apiversion": "0.1.0",
        "Authorization": f"{ACCESS_KEY_ID}:{signature}"
    }

# ==========================================
# 2. FILE UPLOAD API
# ==========================================
def upload_file(local_file_path:Path, interface:AppInterface) -> tuple[bool, str]:
    """
    Uploads a local file to FlipHTML5.
    Returns: (True, file_url) on success, or (False, error_message) on failure.
    """
    print("\n--- 1. Uploading File ---")
    url = "https://api.fliphtml5.com/api/common/upload-file"
    path = "/api/common/upload-file"

    is_path_valid, validation_msg = validate_pdf_path(str(local_file_path))
    if not is_path_valid:
        return False, validation_msg

    try:
        headers = generate_fliphtml5_headers(path, query_params=None)

        interface.print_info(f"[DEBUG] Opening file binary stream...")
        with open(local_file_path, "rb") as f:
            files = {"file": (os.path.basename(local_file_path), f, "application/pdf")}

            interface.print_info(f"[DEBUG] Sending POST request to FlipHTML5...")
            response = requests.post(url, headers=headers, files=files)

        # Check HTTP Status Codes before parsing JSON
        interface.print_info(f"[DEBUG] HTTP Response Status Code: {response.status_code}")
        response.raise_for_status()

        res_data = response.json()
        interface.print_info(f"[DEBUG] Raw API Response: {res_data}")

        if res_data.get("code") == "OK":
            file_src = res_data["data"]["fileSrc"]
            final_url = "https:" + file_src if file_src.startswith("//") else file_src
            return True, final_url
        else:
            return False, f"[API Error] Upload rejected by server. Payload: {res_data}"

    except PermissionError:
        return False, f"[Permission Error] Access Denied to '{local_file_path}'. File may be open elsewhere."
    except requests.exceptions.RequestException as req_err:
        return False, f"[Network Error] Connection or HTTP issue during upload: {req_err}"
    except Exception as e:
        return False, f"[Unexpected Error] Critical failure: {str(e)}"

# ==========================================
# 3. CREATE BOOK API (with design config)
# ==========================================
def create_book(file_src_url, title, description, link, design_config=None):
    print("\n--- 2. Creating Book Task ---")
    url = "https://api.fliphtml5.com/api/book/create-book-multi"
    path = "/api/book/create-book-multi"

    file_path_json_str = json.dumps([{"link": file_src_url}])

    params = {
        "title": title,
        "description": description,
        "filePath": file_path_json_str,
        "bLink": link,
        "folderId": "7398072"
    }

    # Goal 4: Apply design customizations (passed as a JSON string configuration)
    if design_config:
        params["bookConfig"] = json.dumps(design_config)

    headers = generate_fliphtml5_headers(path, query_params=params)

    try:
        response = requests.post(url, headers=headers, data=params)
        res_data = response.json()
        if res_data.get("code") == "OK":
            return res_data["data"]["bookId"]
        else:
            print("Create Error Response:", res_data)
    except Exception as e:
        print("Create Book error:", str(e))
    return None

# ==========================================
# 4. GET BOOK CONVERSION PROGRESS API
# ==========================================
def poll_conversion(book_id):
    print("\n--- 3. Polling Conversion Progress ---")
    url = "https://api.fliphtml5.com/api/book/get-book-progress"
    path = "/api/book/get-book-progress"

    params = {"bookId": str(book_id)}
    headers = generate_fliphtml5_headers(path, query_params=params)

    max_attempts = 60
    attempt = 0
    while attempt < max_attempts:
        attempt += 1
        try:
            response = requests.post(url, headers=headers, data=params)
            res_data = response.json()

            if res_data.get("code") == "OK":
                status = int(res_data["data"]["convertStatus"])
                progress = res_data["data"]["convertProgress"]
                print(f"Check #{attempt} | Status ID: {status} | Progress: {progress}%")

                if status == 5:
                    print("🎉 Success! Book Conversion Complete!")
                    return True
                elif status in [3, 4, 6]:
                    print("❌ Book conversion failed explicitly on FlipHTML5 side.")
                    return False
            time.sleep(10)
        except Exception as e:
            print("Polling network error:", str(e))
            break
    return False

# ==========================================
# 5. SET BOOK ACCESS CONTROL (with Password option)
# ==========================================
def set_book_privacy_with_password(book_id, passwords_list):
    print("\n--- 4. Setting Privacy Controls ---")
    url = "https://api.fliphtml5.com/api/book/set-book-privacy"
    path = "/api/book/set-book-privacy"

    # Goal 1: Set visibility status to 0 (Private with Password)
    # purviewList maps structural permissions in explicit list matrices
    params = {
        "bookId": str(book_id),
        "isPublic": "0",
        "purviewList": json.dumps({"password": passwords_list})
    }
    headers = generate_fliphtml5_headers(path, query_params=params)

    try:
        response = requests.post(url, headers=headers, data=params)
        print("Privacy Settings Response:", response.json())
    except Exception as e:
        print("Privacy configuration error:", str(e))

# ==========================================
# EXECUTION CONTROLLER
# ==========================================
if __name__ == "__main__":
    target_pdf = r"C:\Users\system1\Desktop\3111025.pdf" # Your local input file path

    my_custom_title = input("Please enter a title: ")
    my_custom_desc = input("Please enter a description: ")
    my_custom_link = my_custom_title

    # Goal 1: Set arbitrary passwords required to access the document
    MY_PASSWORDS = [input("Please enter a password: ")]

    # Goal 4: Define custom interface layouts and configurations here.
    # Adjust variables like loading captions, colors, backgrounds, or toolbar buttons
    MY_DESIGN_CONFIG = {
        "loadingCaption": "Fetching Pages, Please Wait...",
        "loadingCaptionColor": "0xFF0000",   # Example: Bright Red text hex code
        "isPrintOn": "true",                 # Enable or disable tool features
        "isDownloadOn": "false"              # Keep files locked down securely
    }

    # Run the automated pipeline
    interface = AppInterface(ui=None)
    uploaded_url = upload_file(target_pdf)
    if uploaded_url:
        print(f"File uploaded successfully to: {uploaded_url}")

        # Create book using custom definitions and design profiles
        my_book_id = create_book(
            file_src_url=uploaded_url,
            title=my_custom_title,
            description=my_custom_desc,
            link = my_custom_link,
            design_config=MY_DESIGN_CONFIG
        )

        if my_book_id:
            print(f"Book context successfully initialized. Target Book ID: {my_book_id}")

            # Wait for backend processing pools to output assets
            if poll_conversion(my_book_id):
                # Lock book visibility down behind authorization passkeys
                set_book_privacy_with_password(my_book_id, MY_PASSWORDS)