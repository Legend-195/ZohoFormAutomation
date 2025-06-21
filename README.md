# Zoho Form Auto-Filler using Python and PyAutoGUI

This project automates the process of filling out a Zoho Form using data from an Excel sheet. The script uses the PyAutoGUI library to simulate user interaction—locating form fields via screenshots, typing entries, handling dropdown selections, uploading images, scrolling the page, and submitting the form.

## 📁 Project Structure

ZohoFormAutomation/

├── data.xlsx # Excel file with form data (row-wise headers)

├── downloads/ # Folder where downloaded images are saved

├── screenshots/ # Folder containing cropped screenshots of form fields

├── form_automation.py # Main Python automation script



## ✅ Features

- Reads Zoho form input data from an Excel file
- Detects form fields using screenshots
- Handles scrolling for long forms
- Supports uploading images via URL (including Yahoo redirect URLs)
- Handles dropdown menus like Country via simulated key presses
- Skips missing fields gracefully
- Logs issues related to missing images or data


## 🧠 How It Works
1. Excel Transposition
The Excel data is stored with field names in the first column and corresponding values in the second. The script transposes this using:

df = raw_df.set_index(0).T.reset_index(drop=True)

2. Screenshot-Based Field Detection
Each form field label is matched with a screenshot saved in the screenshots/ folder. These are used with pyautogui.locateCenterOnScreen() to identify where to click.

3. Typing into Form Fields
After locating each field, the script types the corresponding value using pyautogui.write().

4. Dropdown Selection (Country Field)
If the field is a dropdown like Country, after typing, the script simulates key presses (down and enter) to confirm the selection.

5. Image Upload from URL
Image URLs from the Excel file are processed using requests and saved locally before being uploaded via file path entry in the browser dialog.

6. Scrolling and Submission
The form is scrolled using pyautogui.scroll() to access lower fields and ensure the submit button is visible before clicking it.

## 📝 How to Use
Place your form data in data.xlsx with the first column as field names (e.g., "First Name", "Email") and second column as values.

Capture screenshots of all form field labels using any screen capture tool and save them in the screenshots/ folder. Use descriptive names like:

first_name.png

email.png

phone.png

submit.png

Run the script:


python form_automation.py
The script will:

Open the Zoho form

Wait for it to load

Loop through each Excel row

Fill the form and submit it

## 📸 Required Screenshots
Screenshots should be taken clearly and saved in the screenshots/ directory:

first_name.png

last_name.png

email.png

phone.png

street_address.png

address_2.png

city.png

state.png

postal.png

country.png

image.png (for image upload field)

submit.png
## 🔧 Requirements

Install the required Python packages using:

```bash
pip install pyautogui pandas pillow requests openpyxl opencv-python
