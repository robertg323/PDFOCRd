# PDFOCRd

Windows desktop GUI app (customtkinter + PyMuPDF) that scans a directory of PDFs, checks if they've been OCR'd, and optionally moves non-OCR files to a separate folder.

## Key Details

- Single-file app: `pdf_ocr_check.py`
- OCR detection works by counting extractable text characters per PDF against a configurable threshold (default: 50)
- Non-OCR files can be moved to a destination folder with automatic rename on duplicates (appends `_1`, `_2`, etc.)
- Generates timestamped CSV reports logging OCR status and actions taken
- Packaged as a standalone `.exe` via PyInstaller (`pdf_ocr_check.spec`)

## Repo

- Remote: https://github.com/robertg323/PDFOCRd.git
- Branch: master
