# Zoho Form Auto-Filler using Python and PyAutoGUI

This project automates the process of filling out a Zoho Form using data from an Excel sheet. The script uses the PyAutoGUI library to simulate user interaction—locating form fields via screenshots, typing entries, handling dropdown selections, uploading images, scrolling the page, and submitting the form.

## 📁 Project Structure

ZohoFormAutomation/
├── data.xlsx # Excel file with form data (row-wise headers)
├── downloads/ # Folder where downloaded images are saved
├── screenshots/ # Folder containing cropped screenshots of form fields
├── form_automation.py # Main Python automation script

markdown
Copy
Edit

## ✅ Features

- Reads Zoho form input data from an Excel file
- Detects form fields using screenshots
- Handles scrolling for long forms
- Supports uploading images via URL (including Yahoo redirect URLs)
- Handles dropdown menus like Country via simulated key presses
- Skips missing fields gracefully
- Logs issues related to missing images or data

## 🔧 Requirements

Install the required Python packages using:

```bash
pip install pyautogui pandas pillow requests openpyxl opencv-python
python form_automation.py

