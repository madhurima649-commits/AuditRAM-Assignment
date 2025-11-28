# Approach and Technical Notes

## 1) Supported file types and strategies
- **PDF (.pdf)**: Use `PyMuPDF` (fitz) to search text with coordinates using `page.search_for()`. That returns bounding box rectangles which we draw as red, unfilled rectangles on a copy of the PDF. This preserves the original file.

- **Images (.png, .jpg, .jpeg)**: Use `pytesseract` to perform OCR and `image_to_data` to get word-level bounding boxes. For each word that matches the search string (case-insensitive substring match), a red rectangle is drawn (Pillow). Output is a new image file.

- **Excel (.xlsx)**: Use `openpyxl` to iterate cells. When a cell contains the search string (case-insensitive substring match), we apply a red border to that cell in a copy of the workbook and save it (original workbook is not modified). Excel has no easy "overlay" concept, but the red border mimics a bounding box.

- **Word (.docx)**: There are two modes:
  - If `docx2pdf` is available and Word is installed (Windows/macOS), convert DOCX → PDF, then treat like a PDF (accurate bounding boxes).
  - Otherwise, extract text via `python-docx` and create a textual report listing the paragraphs that contain matches. Exact coordinates cannot be determined without rendering the docx to PDF or an image.

## 2) Matching rules
- Matching is **case-insensitive** and uses substring matching.
- For images, matching uses word-level OCR output; multi-word strings may match if any OCR token contains the search string (simple heuristic). You can extend to group words by proximity for exact phrase matching.

## 3) Overlay vs Modification
- The original input file is never modified. For PDFs and images we produce new files with overlays.
- For Excel, we save a copy workbook with cell borders applied (the original file remains unchanged).
- For DOCX, where precise overlays are not feasible without conversion, either a converted PDF is used (if possible) or a textual report is produced.

## 4) System dependencies
- `tesseract-ocr` (system package) for image OCR.
- `docx2pdf` requires MS Word (Windows/macOS) to perform accurate conversion.

## 5) Improvements (suggested)
- Phrase-level OCR matching: group adjacent OCR tokens into lines and search phrases.
- Better DOCX rendering via headless LibreOffice conversion on Linux (unoconv/soffice) to produce PDFs without Word.
- Create an output ZIP with original + overlay preview for submission.

## 6) Security and performance
- Large PDFs/images may increase memory usage. Process page-by-page for PDFs to reduce memory footprint.
- This implementation is meant for prototype/demo. For production, add robust error handling, logging and unit tests.
