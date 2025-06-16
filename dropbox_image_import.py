#!/usr/bin/env python
# coding: utf-8

# # Script: Download, Extract and Rename Images from Dropbox Links on Excel sheet

# In[ ]:


import pandas as pd
import requests
import os
import zipfile
import io
import shutil
import pytesseract
from PIL import Image

# --- PATH CONFIGURATION (update these to your local paths) ---
TESSERACT_PATH = r"C:\Path\To\Tesseract-OCR\tesseract.exe"  # Tesseract executable path
EXCEL_PATH = r"C:\Path\To\Your\Excel\SriLankan-Aug25.xlsx"  # Excel file path

# Base folder where movies will be organized
BASE_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads", "Srilankan_Downloads")

# Set Tesseract executable path
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# OCR languages
OCR_LANGS = 'eng+chi_sim+jpn+sin+tam'

# Ensure base folder exists
os.makedirs(BASE_FOLDER, exist_ok=True)

# Read the Excel file
df = pd.read_excel(EXCEL_PATH)

# Filter only valid Dropbox links
valid_rows = df[(df['Images'].str.startswith('https://www.dropbox.com', na=False)) & (df['Images'].str.lower() != 'will share asap')]

def detect_text_in_image(image_path):
    try:
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img, lang=OCR_LANGS).strip()
        return text if text else None
    except Exception as e:
        print(f"❌ Error processing OCR on {image_path}: {e}")
        return None

def organize_images_by_ocr(folder_path, title):
    posters_folder = os.path.join(folder_path, 'posters')
    stills_folder = os.path.join(folder_path, 'stills')
    os.makedirs(posters_folder, exist_ok=True)
    os.makedirs(stills_folder, exist_ok=True)

    for filename in sorted(os.listdir(folder_path)):
        file_path = os.path.join(folder_path, filename)

        # Skip folders and unsupported files
        if not os.path.isfile(file_path) or not filename.lower().endswith(('.jpg', '.jpeg', '.png', '.tif', '.tiff')):
            continue

        text = detect_text_in_image(file_path)

        if text:
            dest = os.path.join(posters_folder, filename)
            print(f"🅿️ Text detected in {filename}, moving to posters.")
        else:
            dest = os.path.join(stills_folder, filename)
            print(f"🖼️ No text in {filename}, moving to stills.")

        try:
            shutil.move(file_path, dest)
        except Exception as e:
            print(f"❌ Error moving {filename}: {e}")

for _, row in valid_rows.iterrows():
    title = row['Titles']
    raw_url = row['Images']

    # Fix URL for direct download
    if 'dl=0' in raw_url:
        url = raw_url.replace('dl=0', 'dl=1')
    elif 'dl=1' not in raw_url:
        url = raw_url + '&dl=1'
    else:
        url = raw_url

    # Safe folder name
    safe_title = "".join(c if c.isalnum() or c in (' ', '-', '_') else "_" for c in title).strip()
    extract_path = os.path.join(BASE_FOLDER, safe_title)
    os.makedirs(extract_path, exist_ok=True)

    print(f"\n⬇️ Downloading and extracting: {title}")

    try:
        response = requests.get(url)
        if response.status_code == 200:
            with zipfile.ZipFile(io.BytesIO(response.content)) as thezip:
                thezip.extractall(extract_path)

            # Rename extracted files
            files = os.listdir(extract_path)
            for i, filename in enumerate(files, start=1):
                old_path = os.path.join(extract_path, filename)
                ext = os.path.splitext(filename)[1]
                new_filename = f"{safe_title}_{i}{ext}"
                new_path = os.path.join(extract_path, new_filename)
                try:
                    os.rename(old_path, new_path)
                except Exception as e:
                    print(f"Error renaming {filename}: {e}")

            print(f"✅ Extracted and renamed at: {extract_path}")

            # Organize images based on OCR
            organize_images_by_ocr(extract_path, title)

        else:
            print(f"❌ Error downloading '{title}': status {response.status_code}")
    except zipfile.BadZipFile:
        print(f"❌ Invalid file for '{title}'")
    except Exception as e:
        print(f"❌ Error processing '{title}': {e}")

