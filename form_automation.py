import pyautogui  # For automating mouse and keyboard actions
import pandas as pd  # For reading and manipulating Excel data
import time  # For adding delays between automated actions
import requests  # For downloading images from URLs
from PIL import Image  # For processing and saving image files
from io import BytesIO  # For handling in-memory image data
import os  # For file path and directory handling
import webbrowser  # To open the Zoho form in the default browser
import urllib.parse  # For parsing complex image URLs

# === File paths ===
DATA_FILE = "data.xlsx"  # Path to the input Excel file
SCREENSHOT_DIR = "screenshots"  # Folder containing field label screenshots
DOWNLOAD_DIR = "downloads"  # Folder to save downloaded images
ZOHO_FORM_URL = "https://zfrmz.in/cJlLgptz9KIQTjmikoxi"  # Target Zoho form URL

# === Setup folders ===
# Create the download folder if it doesn't exist
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

# === Load Excel and transpose ===
# Read the Excel data without assuming a header row
raw_df = pd.read_excel(DATA_FILE, header=None)
# Use the first column as header, transpose, and reset row index
df = raw_df.set_index(0).T.reset_index(drop=True)

def download_image(url, idx):
    """
    Downloads an image from a URL and saves it locally.
    Handles special Yahoo redirect URLs by extracting the real image link.
    """
    try:
        # Extract direct image link from Yahoo search-style URL
        if "iurl=" in url:
            parsed = urllib.parse.urlparse(url)
            query_params = urllib.parse.parse_qs(parsed.fragment)  # Yahoo puts data in the fragment
            direct_urls = query_params.get("iurl")
            if direct_urls:
                url = direct_urls[0]  # Replace with actual image URL

        # Download the image
        response = requests.get(url)
        if response.status_code == 200:
            path = f"{DOWNLOAD_DIR}/image_{idx}.jpg"
            img = Image.open(BytesIO(response.content))  # Load image from memory
            img.save(path)  # Save to disk
            return path
        else:
            print(f"[HTTP Error] Status {response.status_code} for row {idx}")
    except Exception as e:
        print(f"[Image Download Error] Row {idx}: {e}")
    return None

def locate_and_click(image_file):
    """
    Locates a field on the screen using a screenshot image, and clicks on it.
    """
    full_path = os.path.join(SCREENSHOT_DIR, image_file)
    if not os.path.exists(full_path):
        print(f"[Missing Screenshot] {image_file}")
        return False
    try:
        location = pyautogui.locateCenterOnScreen(full_path, confidence=0.9)
        if location:
            pyautogui.click(location)  # Click the center of the found image
            time.sleep(0.5)
            return True
        else:
            print(f"[Not Found On Screen] {image_file}")
            return False
    except:
        print(f"[Error] Could not locate {image_file}")
        return False

def fill_field(image_file, value):
    """
    Fills a form field by locating its screenshot and typing the provided value.
    Handles special behavior for dropdown fields like 'Country'.
    """
    if value is None or pd.isna(value):
        print(f"[Skip] Value missing for {image_file}")
        return
    if locate_and_click(image_file):
        pyautogui.write(str(value), interval=0.05)  # Type with small delay
        time.sleep(0.3)

        # If it's the Country field, press keys to select dropdown match
        if "country" in image_file.lower():
            pyautogui.press('down')   # Select the first match
            time.sleep(0.2)
            pyautogui.press('enter')
            pyautogui.press('down')
            pyautogui.press('down')
            pyautogui.press('down')
            pyautogui.press('enter')  # Confirm selection
            print(f"[Dropdown Selected] {value}")

def upload_image_from_url(image_file, url, idx):
    """
    Downloads an image from the given URL and uploads it into the form.
    """
    if url is None or pd.isna(url):
        print(f"[Skip] Image URL missing for row {idx}")
        return
    img_path = download_image(url, idx)
    if img_path and locate_and_click(image_file):
        time.sleep(4)
        pyautogui.write(os.path.abspath(img_path))  # Write full path into upload dialog
        pyautogui.press('enter')  # Confirm file selection
        time.sleep(1)

def safe_get(row, column):
    """
    Safely gets the value from a column in the current row.
    """
    if column in row:
        return row[column]
    print(f"[Skip] Column '{column}' not found")
    return None

def process_form(row, idx):
    """
    Fills out a single form using values from one Excel row.
    Handles scrolling, field detection, typing, and image upload.
    """
    # Fields visible before scrolling
    fields_before_scroll = {
        "first_name.png": "First Name",
        "last_name.png": "Last Name",
        "phone.png": "Mobile",
        "email.png": "Email",
        "street_address.png": "Address"
    }

    # Fields visible after scrolling
    fields_after_scroll = {
        "city.png": "City",
        "state.png": "State",
        "postal.png": "Postal Code",
        "country.png": "Country",
        "address_2.png": "Address Line 2"
    }

    # Fill upper part of the form
    for img, col in fields_before_scroll.items():
        value = safe_get(row, col)
        fill_field(img, value)

    # Scroll to access lower fields
    print("[Scroll] Scrolling down for lower fields...")
    pyautogui.scroll(-1000)
    time.sleep(1)

    # Fill lower part of the form
    for img, col in fields_after_scroll.items():
        value = safe_get(row, col)
        fill_field(img, value)

    # Handle image upload
    upload_image_from_url("image.png", safe_get(row, "image url:"), idx)
    time.sleep(2.5)

    # Scroll again in case submit button is out of view
    pyautogui.scroll(-500)
    time.sleep(1)

    # Locate and click the Submit button
    locate_and_click("submit.png")
    time.sleep(3)

# === Start automation ===
print("Opening Zoho Form in browser...")
webbrowser.open(ZOHO_FORM_URL)  # Launch the form in the browser

print("Waiting 10 seconds for the page to load...")
time.sleep(10)  # Allow page to fully load

# Process each row in the Excel sheet
for idx, row in df.iterrows():
    print(f"\n--- Filling Form for Row {idx + 1} ---")
    process_form(row, idx)
    time.sleep(5)  # Small gap between form submissions
