# AuditRAM - Coding Assignment

This repository contains a ready-to-run Python solution for the AuditRAM coding assignment: search text in uploaded files (PDF, images, Excel, Word) and produce an **overlay** that highlights matches with **red, unfilled bounding boxes**—without modifying the original input file.

Files:
- `main.py` - Main program. Supports command-line and Streamlit UI.
- `APPROACH.md` - Detailed description of how the solution works and limitations.
- `requirements.txt` - Python dependencies.
- `example_run.txt` - Quick example commands.

## Quick example (command-line)

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. For image OCR results you must install Tesseract-OCR on your system:
- Ubuntu: `sudo apt install tesseract-ocr`
- macOS (Homebrew): `brew install tesseract`
- Windows: download installer from Tesseract project and add to PATH.

3. Run the program:
```bash
python main.py --input /path/to/file.pdf --text "Invoice Number" --output /path/to/output.pdf
```

For images:
```bash
python main.py -i /path/to/image.png -t "Invoice" -o /path/to/output.png
```

For Excel:
```bash
python main.py -i /path/to/workbook.xlsx -t "Total" -o /path/to/workbook_with_borders.xlsx
```

For DOCX:
- If `docx2pdf` is installed and Word is available (Windows/macOS), the code will convert DOCX → PDF and then draw overlays.
- Otherwise a textual report (`.txt`) listing paragraphs with matches will be produced.

## Streamlit UI

Optionally run with a simple Streamlit interface:
```bash
pip install streamlit
streamlit run main.py -- --ui
```

## Notes & Limitations

Please see `APPROACH.md` for full technical notes and limitations (important: DOCX → PDF conversion requires Word & docx2pdf; image OCR requires Tesseract).

