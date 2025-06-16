# 📥 Bulk Dropbox Image Downloader and OCR-Based Sorting

## 📝 About this Project

This project automates downloading, extracting, and organizing movie images from Dropbox links listed in an Excel file. It streamlines image sourcing workflows for inflight entertainment content management.

The script downloads ZIP archives, extracts images, uses OCR to detect multilingual text for poster identification, and sorts images into folders by movie title, reducing manual work.

---

## 🛠 Technologies Used

- Python for scripting and automation  
- pandas for Excel handling  
- requests for HTTP downloads  
- zipfile for archive extraction  
- Pillow for image processing  
- pytesseract for OCR with multilingual support  
- OS and shutil for file operations  

---

## ✨ Key Features

- Bulk downloads with direct URL correction  
- Extraction and renaming to standardized filenames  
- OCR classification into posters (with text) and stills (without text)  
- Automated folder structure by movie title  
- Robust error handling for all steps  

---

## 📊 Impact

- Reduced manual sorting and renaming by over 80%  
- Scalable pipeline for multiple distributors  
- Improved image classification accuracy through OCR  
- Efficient management of inflight entertainment media  

---

## ⚠️ Notes

- Tesseract OCR must be installed locally and configured in the script  
- Demonstrates practical skills in data ingestion, file management, OCR, and automation  
