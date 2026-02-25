# PDFOCRd

A Windows desktop application that scans a directory of PDF files and identifies which ones have been OCR'd (contain extractable text) and which have not.

## Features

- **OCR Detection** - Checks each PDF for extractable text using PyMuPDF. Files with fewer characters than a configurable threshold are classified as non-OCR.
- **File Sorting** - Optionally moves non-OCR PDFs to a separate destination folder to organize them.
- **Duplicate Handling** - Automatically renames files when a duplicate filename already exists in the destination (e.g., `document_1.pdf`, `document_2.pdf`).
- **CSV Reporting** - Generates a timestamped CSV log of every scanned file, its OCR status, and the action taken.
- **PDF Preview** - Click any scanned file in the results list to preview it directly in the app.
- **Configurable Threshold** - Set the minimum character count required for a PDF to be considered OCR'd (default: 50).
- **Real-time Log** - Displays scan progress in the GUI as files are processed.
- **OCR in Place** - Optionally run Tesseract OCR on non-OCR PDFs directly from the app (requires Tesseract installation, see below).

## Download

A pre-built standalone Windows executable is available on the [Releases](../../releases) page — no Python installation required.

> **Note:** Even when using the standalone `.exe`, Tesseract OCR must still be installed separately if you want to use the **OCR in Place** feature (see below).

## Prerequisites

### Tesseract OCR (required for "OCR in Place" feature)

The **OCR in Place** feature uses [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) to add a text layer to scanned PDFs. Tesseract is a separate program that must be installed on your system.

**Installation steps:**

1. Download the latest Tesseract installer for Windows from:
   **https://github.com/UB-Mannheim/tesseract/wiki**

2. Run the installer. The default installation path is:
   ```
   C:\Program Files\Tesseract-OCR\tesseract.exe
   ```
   PDFOCRd expects Tesseract at this default path. Keep the default unless you know what you are doing.

3. During installation, select any additional language packs you need (English is included by default).

4. No additional configuration is needed — PDFOCRd will detect Tesseract automatically.

> **Note:** Tesseract is only required for the **OCR in Place** feature. Scanning PDFs for OCR status and sorting/moving files works without Tesseract.

## Running from Source

### Requirements

- Python 3.10+
- PyMuPDF
- customtkinter
- pytesseract
- pdf2image
- Pillow

Install Python dependencies:

```
pip install -r requirements.txt
```

### Usage

```
python pdf_ocr_check.py
```

1. Set the **Source Directory** containing your PDFs.
2. Set the **Destination Directory** for non-OCR files (if moving is enabled).
3. Adjust the **Min. Text Chars** threshold if needed.
4. Optionally enable **CSV Report** generation and choose a report directory.
5. Click **START SCAN**.
6. Click any file in the **Scan Results** list to preview it. Use **Show Preview** to toggle the preview panel.

## Building an Executable

The included `.spec` file can be used with PyInstaller to build a standalone `.exe`:

```
pyinstaller pdf_ocr_check.spec
```
