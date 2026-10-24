# Product Requirements: Merge & convert

## Purpose
One simple web app to merge, split and convert documents. Brand name everywhere: "Merge & convert".

## Users
People who need quick document tasks without installing software: students, office staff, freelancers.

## Scope
Merge: PDF, Word, Excel, PowerPoint, Markdown, Images, TXT, CSV.
Split PDF: by range, extract selected pages, every N pages, single pages, remove pages.
Convert to PDF: JPG, Word, PowerPoint, HTML, Markdown.
Convert from PDF: JPG, Word, PowerPoint, PDF/A, Markdown.
Compress: JPG, PNG, PDF, PowerPoint and Word. The app measures the smallest clear size for each file; the user then types any size from that minimum up (KB or MB) and gets the best quality that fits. PDFs keep text and vectors; only pictures inside are re-saved.
Out of scope for now: accounts, history of past jobs, Excel to PDF and PDF to Excel.

## Experience requirements
- Same 4-step flow for every tool: Select, Upload, Settings, Download.
- Settings appear only after a file is uploaded; Download appears after the job runs.
- Preview button shows the result before download (PDF, text, Excel table, first JPG page).
- Light and dark mode with a remembered choice; responsive layout with a bottom tab bar on phones.
- Accessible: keyboard use, visible focus, labels, live announcements, reduced-motion support.
- Footer states what the app can do and that files are deleted automatically after processing.

## Technical requirements
- Frontend: HTML, CSS and plain JavaScript in separate files; pdf-lib, SheetJS, JSZip, pdf.js and PptxGenJS from CDN.
- Backend: Node and Express. Needs Chrome, Edge or Chromium for HTML, Markdown, Word and PowerPoint to PDF; LibreOffice (optional) improves Word and PowerPoint layout; Ghostscript is needed for PDF/A.
- Files are processed in a temp folder and deleted right after the response.
- Limit: 100 MB per file.

## Acceptance criteria
1. Brand reads "Merge & convert" everywhere.
2. Each merge type, split mode and convert tool listed above completes and offers a download.
3. Settings are hidden until a valid file is added.
4. Errors are clear: wrong type, empty, damaged, protected, too large, missing converter.
5. Works on mobile, tablet and desktop in light and dark mode.

## Known limits and next steps
- PDF to Word, Markdown and PowerPoint work from the PDF text; scanned pages are read with built-in OCR (Tesseract.js; language data downloads on first use, so it needs internet). PDF to PowerPoint makes one picture slide per page with the text in the notes.
- Word and PowerPoint to PDF without LibreOffice keep text, headings, lists and tables only; PowerPoint slides lose images and layout.
- PDF/A output is not validated.
- Password-protected PDFs are rejected, not unlocked.
- Word merge keeps pictures, links, styles and lists; PowerPoint merge uses the first deck's theme and drops notes, charts and animations.
- Next: HTML from URL, page thumbnails, React version, validation for PDF/A.
