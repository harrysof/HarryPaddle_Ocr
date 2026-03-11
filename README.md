# PaddleOCR Desktop Demo

A powerful, user-friendly desktop application for Optical Character Recognition (OCR) and Table Extraction, built with Python, Tkinter, and PaddleOCR.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)
![PaddleOCR](https://img.shields.io/badge/PaddleOCR-3.3.3-orange.svg)

## Key Features

- **Multi-Format Support**: Process single images (JPG, PNG, BMP, TIFF) and multi-page scanned PDFs.
- **Multilingual OCR**: High-accuracy recognition for **French**, **Arabic**, and **English**.
- **Batch Processing**: Recursively process entire folders of images and PDFs.
- **Table Extraction**: Advanced PPStructure pipeline to detect tables and export them directly to Excel (.xlsx).
- **Interactive Preview**: Visual feedback with bounding boxes for detected text (Blue) and tables (Green).
- **Modern UI**: Clean Tkinter interface with dark-themed buttons, progress bars, and responsive threading.

## Installation

### 1. Prerequisites
- Python 3.8 or higher.
- [Poppler](https://github.com/oschwartz10612/poppler-windows/releases) (Required for PDF support on Windows). Add the `bin` folder to your System PATH.

### 2. Install Dependencies
```bash
pip install paddlepaddle==3.2.0 paddleocr==3.3.3 pdf2image pillow openpyxl
```

## How to Use

1. **Run the Script**:
   ```bash
   python paddleocr_demo.py
   ```
2. **Select Language**: Choose your target language (French/Arabic/English) from the dropdown.
3. **Load File**:
   - Use **📂 Open Image** for single pictures.
   - Use **📄 Open PDF** for scanned documents.
   - Use **📁 Batch Folder** to select an input folder and an output destination for bulk processing.
4. **Process**:
   - Click **▶ Run OCR** for standard text extraction.
   - Click **🗂 Extract Tables** to find and export tables to Excel.
5. **Save**:
   - Standard OCR results can be saved as a `.txt` file using **💾 Save Text**.
   - Table results will prompt for a `.xlsx` save location automatically.

## Project Structure

- `paddleocr_demo.py`: The main application script.
- `batch_summary.xlsx`: (Generated) Summary of OCR results during batch mode.
- `*_ocr.txt`: (Generated) Individual text outputs.
- `*_tables.xlsx`: (Generated) Extracted tables from the document.

## Important Notes

- **First Run**: PaddleOCR will download the necessary pre-trained models (~100MB per language). This only happens once.
- **Orientation**: The app uses `use_textline_orientation=False` by default for faster processing.
- **Threading**: All heavy tasks run in separate threads to keep the UI responsive.

## 📄 License
This project is open-source and available under the [MIT License](LICENSE).
