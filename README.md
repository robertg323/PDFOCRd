# PDFOCRd

A Windows desktop application that scans a directory of PDF files and identifies which ones have been OCR'd (contain extractable text) and which have not.

## Features

- **OCR Detection** - Checks each PDF for extractable text using PyMuPDF. Files with fewer characters than a configurable threshold are classified as non-OCR.
- **File Sorting** - Optionally moves non-OCR PDFs to a separate destination folder to organize them.
- **Duplicate Handling** - Automatically renames files when a duplicate filename already exists in the destination (e.g., `document_1.pdf`, `document_2.pdf`).
- **CSV Reporting** - Generates a timestamped CSV log of every scanned file, its OCR status, and the action taken.
- **Configurable Threshold** - Set the minimum character count required for a PDF to be considered OCR'd (default: 50).
- **Real-time Log** - Displays scan progress in the GUI as files are processed.

## Requirements

- Python 3
- PyMuPDF
- customtkinter

Install dependencies:

```
pip install -r requirements.txt
```

## Usage

```
python pdf_ocr_check.py
```

1. Set the **Source Directory** containing your PDFs.
2. Set the **Destination Directory** for non-OCR files (if moving is enabled).
3. Adjust the **Min. Text Chars** threshold if needed.
4. Optionally enable **CSV Report** generation and choose a report directory.
5. Click **START SCAN**.

## Building an Executable

The included `.spec` file can be used with PyInstaller to build a standalone `.exe`:

```
pyinstaller pdf_ocr_check.spec
```
