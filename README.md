# DataExtract 300

Lightweight browser-only tool to extract structured records from PDFs and images (JPG/PNG) and export them to Excel.

## Features
- Multi-file upload (.pdf, .jpg, .jpeg, .png)
- Text extraction from PDFs via PDF.js
- OCR fallback (Tesseract.js) for scanned PDFs and images
- Auto-extraction of: Name, Phone, LG, Ward, Position, Name of institution, Application No
- Multiple records per document
- Editable results table with auto SN
- Duplicate detection (name / phone / application no)
- Missing-field highlighting
- Export to .xlsx via SheetJS (leading zeros preserved in phone column)

## Usage
Open the GitHub Pages URL (or index.html locally) in any modern browser.

1. Select documents.
2. Tap Process.
3. Review / edit / add / delete rows.
4. Tap Export Excel.

## Tech
- Vanilla HTML/CSS/JS - no build step, no backend.
- CDN libraries: PDF.js, Tesseract.js, SheetJS.

## Deploy
GitHub Pages serves the repo root from the `main` branch.
