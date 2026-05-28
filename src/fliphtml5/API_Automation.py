import os
import json
import time
import base64
import hmac
import hashlib
from datetime import datetime, timezone
import requests
from dotenv import load_dotenv

# Load the variables from the .env file into the system environment
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

    authorization_header = f"{ACCESS_KEY_ID}:{signature}"

    return {
        "Date": date_str,
        "x-yzw-apiversion": "0.1.0",
        "Authorization": authorization_header
    }


# ==========================================
# 2. FILE UPLOAD API
# ==========================================
def test_upload_file(local_file_path):
    print("\n--- 1. Testing File Upload ---")
    url = "https://api.fliphtml5.com/api/common/upload-file"
    path = "/api/common/upload-file"

    headers = generate_fliphtml5_headers(path, query_params=None)

    if not os.path.exists(local_file_path):
        print(f"Error: Local test file '{local_file_path}' not found!")
        return None

    try:
        with open(local_file_path, "rb") as f:
            files = {"file": (os.path.basename(local_file_path), f, "application/pdf")}
            response = requests.post(url, headers=headers, files=files)

        print("Upload Response:", response.json())

        res_data = response.json()
        if res_data.get("code") == "OK":
            file_src = res_data["data"]["fileSrc"]
            if file_src.startswith("//"):
                file_src = "https:" + file_src
            return file_src
    except Exception as e:
        print("Upload error:", str(e))
    return None


# ==========================================
# 3. CREATE BOOK MULTI API
# ==========================================
def test_create_book(file_src_url):
    print("\n--- 2. Testing Create Book (Multi) ---")
    url = "https://api.fliphtml5.com/api/book/create-book-multi"
    path = "/api/book/create-book-multi"

    file_path_json_str = json.dumps([{"link": file_src_url}])

    params = {
        "title": "Python Test Book",
        "description": "An automated flipping book created via API",
        "filePath": file_path_json_str,
        "htmlTemplate": "Minimalist"
    }

    headers = generate_fliphtml5_headers(path, query_params=params)

    try:
        response = requests.post(url, headers=headers, data=params)
        print("Create Book Response:", response.json())

        res_data = response.json()
        if res_data.get("code") == "OK":
            return res_data["data"]["bookId"]
    except Exception as e:
        print("Create Book error:", str(e))
    return None


# ==========================================
# 4. GET BOOK CONVERSION PROGRESS API
# ==========================================
def test_poll_conversion(book_id):
    print("\n--- 3. Testing Get Conversion Progress ---")
    url = "https://api.fliphtml5.com/api/book/get-book-progress"
    path = "/api/book/get-book-progress"

    params = {"bookId": str(book_id)}
    headers = generate_fliphtml5_headers(path, query_params=params)

    for i in range(10):
        try:
            response = requests.post(url, headers=headers, params=params)
            res_data = response.json()
            print(f"Check #{i + 1}: Progress Response:", res_data)

            if res_data.get("code") == "OK":
                status = int(res_data["data"]["converStatus"])
                progress = res_data["data"]["converProgress"]

                print(f"Current Status ID: {status} | Progress: {progress}%")

                if status == 5:
                    print("🎉 Success! Book Conversion Complete!")
                    return True
                elif status in [3, 4, 6]:
                    print("❌ Book conversion failed or errored out.")
                    return False

            time.sleep(5)
        except Exception as e:
            print("Polling error:", str(e))
            break
    return False


# ==========================================
# 5. MODIFY BOOK INFORMATION API
# ==========================================
def test_modify_book(book_id):
    print("\n--- 4. Testing Modify Book Information ---")
    url = "https://api.fliphtml5.com/api/book/update-book"
    path = "/api/book/update-book"

    params = {
        "bookId": str(book_id),
        "title": "Python Updated Test Title",
        "description": "Successfully modified this description via API code!"
    }
    headers = generate_fliphtml5_headers(path, query_params=params)

    try:
        response = requests.post(url, headers=headers, params=params)
        print("Modify Response:", response.json())
    except Exception as e:
        print("Modify error:", str(e))


# ==========================================
# 6. SET BOOK ACCESS CONTROL API
# ==========================================
def test_set_privacy(book_id):
    print("\n--- 5. Testing Set Book Access Control ---")
    url = "https://api.fliphtml5.com/api/book/set-book-privacy"
    path = "/api/book/set-book-privacy"

    params = {
        "bookId": str(book_id),
        "isPublic": "2"  # 2 means Completely Private
    }
    headers = generate_fliphtml5_headers(path, query_params=params)

    try:
        response = requests.post(url, headers=headers, params=params)
        print("Privacy Response:", response.json())
    except Exception as e:
        print("Privacy error:", str(e))


# ==========================================
# MASTER WORKFLOW CONTROL
# ==========================================
if __name__ == "__main__":
    test_pdf_file = r"R:\Documents\001אתר האינטרנט ופרויקטים דיגיטליים\הכנת כתבי עת לאתר\הכנת ספרים לאתר\2159.pdf"

    # 1. Upload File
    uploaded_url = test_upload_file(test_pdf_file)

    if uploaded_url:
        print(f"File uploaded successfully to: {uploaded_url}")

        # 2. Create the Book
        my_book_id = test_create_book(uploaded_url)

        if my_book_id:
            print(f"Book created successfully! ID is: {my_book_id}")

            # 3. Poll conversion status
            conversion_success = test_poll_conversion(my_book_id)

            # 4. Modify and secure if ready
            if conversion_success:
                test_modify_book(my_book_id)
                test_set_privacy(my_book_id)
    else:
        print("Aborting remaining tests because file upload failed.")