# PaddleOCR Desktop Demo

A powerful, user-friendly desktop application for Optical Character Recognition (OCR) and Table Extraction, built with Python, Tkinter, and PaddleOCR.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)
![PaddleOCR](https://img.shields.io/badge/PaddleOCR-3.3.3-orange.svg)

## Key Features

- **Multi-Format Support**: Process single images (JPG, PNG, BMP, TIFF) and multi-page scanned PDFs.
- **Smart PDF Text-Layer Detection**: On PDF load, automatically checks for an existing embedded text layer and verifies it against a quick OCR sample of page 1. If trustworthy, extracts it directly (instant, no OCR) — if missing or corrupted (e.g. broken font encoding dropping accented characters, common in some invoicing software), automatically falls back to full OCR.
- **Multilingual OCR**: High-accuracy recognition for **French**, **Arabic**, and **English**.
- **Batch Processing**: Recursively process entire folders of images and PDFs.
- **Table Extraction**: Advanced PPStructure pipeline to detect tables and export them directly to Excel (.xlsx).
- **Structured Markdown Export**: Converts detected layout regions (titles, paragraphs, lists, tables) into a clean, readable `.md` file — proper heading levels, GFM tables, and bullet lists instead of a flat text dump.
- **Interactive Preview**: Visual feedback with bounding boxes for detected text (Blue) and tables (Green).
- **Modern UI**: Clean Tkinter interface with dark-themed buttons, progress bars, and responsive threading.

## Installation

### 1. Prerequisites
- Python 3.8 or higher.
- [Poppler](https://github.com/oschwartz10612/poppler-windows/releases) (Required for PDF support on Windows). Add the `bin` folder to your System PATH.

### 2. Install Dependencies
```bash
pip install paddlepaddle==3.2.0 paddleocr==3.3.3 pdf2image pillow openpyxl pymupdf
```

## How to Use

1. **Run the Script**:
   ```bash
   python paddleocr_demo.py
   ```
2. **Select Language**: Choose your target language (French/Arabic/English) from the dropdown.
3. **Load File**:
   - Use **📂 Open Image** for single pictures.
   - Use **📄 Open PDF** for scanned or digital PDFs — the app automatically checks for an embedded text layer and verifies it against a quick OCR sample; if it's trustworthy, the text is extracted instantly with no OCR needed, and if it's missing or corrupted, OCR runs automatically instead.
   - Use **📁 Batch Folder** to select an input folder and an output destination for bulk processing.
4. **Process**:
   - Click **▶ Run OCR** for standard text extraction.
   - Click **🗂 Extract Tables** to detect layout (titles, text blocks, tables, lists) and export tables to Excel.
5. **Save**:
   - Standard OCR results can be saved as a `.txt` file using **💾 Save Text**.
   - Table results will prompt for a `.xlsx` save location automatically.
   - Click **📝 Export Markdown** to save a structured `.md` file built from the layout regions detected by Extract Tables (requires running Extract Tables first — plain OCR alone doesn't have the layout info needed).

## Project Structure

- `paddleocr_demo.py`: The main application script.
- `batch_summary.xlsx`: (Generated) Summary of OCR results during batch mode.
- `*_ocr.txt`: (Generated) Individual text outputs.
- `*_tables.xlsx`: (Generated) Extracted tables from the document.
- `*.md`: (Generated) Structured Markdown export — headings, paragraphs, lists, and GFM tables rebuilt from the detected layout.

## Important Notes

- **First Run**: PaddleOCR will download the necessary pre-trained models (~100MB per language). This only happens once.
- **Orientation**: The app uses `use_textline_orientation=False` by default for faster processing.
- **Threading**: All heavy tasks run in separate threads to keep the UI responsive.
- **Markdown headings are heuristic**: heading levels are inferred from PPStructure's `title` tag plus relative text-region size (bbox height vs. document median). This works well in practice but isn't exact — if headings are over- or under-triggering on a given document, the size-ratio thresholds in `build_markdown_document()` can be tuned.
- **Markdown tables don't merge cells**: rowspan/colspan in source tables aren't reconstructed; merged cells appear as repeated or empty cells in the output table.
- **Text-layer trust check needs calibration**: the similarity threshold between a PDF's embedded text and the OCR sample (`TEXT_LAYER_TRUST_THRESHOLD` in the code, default 0.82) is a starting guess — since OCR itself isn't perfect either, run it against a few known-good and known-corrupted PDFs and adjust if it's over- or under-triggering the OCR fallback.

## 📄 License
This project is open-source and available under the [MIT License](LICENSE).
