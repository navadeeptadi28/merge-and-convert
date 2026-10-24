# Merge & convert

A free, self-hosted web app to **merge, split, compress and convert** PDF, Word, Excel, PowerPoint, Markdown, image, TXT and CSV files. Most tools run entirely in your browser, so your files never leave your device. A small Node.js server adds the conversions that need desktop-grade engines (LibreOffice, Microsoft Office, Ghostscript, Chrome).

Works on desktop and phone, with light and dark mode.

---

## Table of contents

1. [What is it used for?](#what-is-it-used-for)
2. [Features](#features)
3. [Quick start](#quick-start)
4. [Requirements](#requirements)
5. [How it works](#how-it-works)
6. [Project structure](#project-structure)
7. [Server API](#server-api)
8. [Configuration](#configuration)
9. [Privacy and limits](#privacy-and-limits)
10. [Troubleshooting](#troubleshooting)
11. [Customizing](#customizing)

---

## What is it used for?

Use it when you need to quickly work with documents without installing heavy software or uploading private files to a random website:

- Combine several PDFs, Word files, slide decks, spreadsheets or images into one file.
- Cut a PDF into smaller files or pull out specific pages.
- Shrink a file to an exact size (for example, "under 200 KB") for email or a government/job portal upload form.
- Convert Word, PowerPoint, HTML, Markdown or images to PDF.
- Convert a PDF to Word, PowerPoint, JPG, Markdown or PDF/A (archive format).

## Features

### Merge files
| Type | Output | Notes |
|---|---|---|
| PDF | PDF | Page ranges and rotation per file, optional page numbers |
| Word (.docx) | .docx | Keeps pictures, links, styles and lists; optional page breaks and file-name headings |
| PowerPoint (.pptx) | .pptx | Uses the first deck's theme; optional divider slides |
| Excel (.xlsx) | .xlsx | Each file as a sheet, or all rows combined into one sheet |
| Markdown | .md | Separator, table of contents, heading level offset |
| Images (JPG, PNG) | PDF | Page size, orientation, margin |
| TXT / CSV | .txt / .csv | CSV: delimiter detection, duplicate removal, column check |

Reorder files by drag and drop or the arrow buttons, sort A-Z, reverse, and preview the result before downloading.

### Split PDF
By page range, every N pages, into single pages, or by picking or removing pages.

### Compress files
Compress **JPG, PNG, WebP, PDF, Word, PowerPoint and Excel** to a size you choose.

1. Add a file. The app measures the **smallest size possible while the file stays clear**.
2. Type the size you want (KB or MB), use the slider, or tap a quick pick such as "Half" or "500 KB".
3. See an estimated clarity rating, then compress and compare before/after sizes.

A target below the minimum is not allowed, so results stay readable.

### Convert to PDF
| Tool | Engine |
|---|---|
| JPG to PDF | Browser |
| Word to PDF | Microsoft Word (Windows) or LibreOffice on the server. Same layout and fonts |
| PowerPoint to PDF | Microsoft PowerPoint (Windows) or LibreOffice on the server. One page per slide |
| HTML to PDF | Chrome / Edge / Chromium on the server |
| Markdown to PDF | Chrome / Edge / Chromium on the server, with Clean, GitHub and Academic themes |

### Convert from PDF
| Tool | Engine |
|---|---|
| PDF to JPG | Browser (choose 72, 150 or 300 DPI) |
| PDF to Word | LibreOffice on the server keeps layout and fonts. Falls back to a basic text conversion in the browser. OCR for scanned pages |
| PDF to PowerPoint | Browser, one slide per page |
| PDF to PDF/A (1b, 2b) | Ghostscript, else LibreOffice, else a best-effort file made in the browser (not validated) |
| PDF to Markdown | Browser, detects headings, lists and tables, with OCR |

### Also included
- Live **preview** of results on all devices: PDF pages, Word, PowerPoint slides and formatted Markdown.
- Tool search (`/` or `Ctrl+K`), step-by-step flow, light and dark theme.
- Mobile layout with a bottom tab bar.

---

## Quick start

### Option A: Docker (easiest, everything included)

```bash
docker build -t merge-and-convert .
docker run -p 3000:3000 merge-and-convert
```

Open <http://localhost:3000>.

The Docker image already contains LibreOffice, Ghostscript, Chromium, Poppler and OCR tools, so every feature works.

### Option B: Run with Node.js

```bash
git clone <your-repo-url>
cd merge-and-convert
npm install
npm start
```

Open <http://localhost:3000>.

To use another port:

```bash
# Mac / Linux
PORT=8080 npm start

# Windows (PowerShell)
$env:PORT=8080; npm start

# Windows (Command Prompt)
set PORT=8080 && npm start
```

### Option C: Open the page without the server

Double-click `index.html`. Merge, split, compress, JPG to PDF, PDF to JPG, PDF to PowerPoint and PDF to Markdown work in the browser alone. Tools that need the server (Word/PowerPoint to PDF, HTML/Markdown to PDF, PDF/A) will say so. If you use another port, change `http://localhost:3000` in `index.html` (the line that sets `window.API_URL`).

---

## Requirements

**Always:**
- A modern browser (Chrome, Edge, Firefox, Safari).
- An internet connection when the page loads. Helper libraries are loaded from public CDNs (see [How it works](#how-it-works)).

**For the server (Option B):**
- Node.js 18 or newer.

**Optional programs on the computer that runs the server.** A missing program shows a clear message in the app instead of crashing.

| Feature | Needs |
|---|---|
| Word / PowerPoint to PDF (exact layout) | Microsoft Word / PowerPoint (Windows) **or** [LibreOffice](https://www.libreoffice.org) |
| PDF to Word (keeps layout and fonts) | LibreOffice |
| HTML / Markdown to PDF | Google Chrome, Microsoft Edge or Chromium |
| PDF to PDF/A (validated) | [Ghostscript](https://ghostscript.com) or LibreOffice |
| PDF to JPG on the server (optional path) | Poppler (`pdftoppm`) |
| OCR with the server engine (optional) | `ocrmypdf` and Tesseract |

If a program is installed in an unusual place, set `SOFFICE_PATH` or `CHROME_PATH` before `npm start`.

---

## How it works

The app has two parts: a **frontend** that does most of the work in the browser, and a **server** for conversions a browser cannot do.

```
 Browser (index.html + js/)                           Node.js server (server/server.js)
 ┌───────────────────────────────────────────┐        ┌───────────────────────────────────┐
 │ Merge, Split, Compress, Preview           │        │ POST /api/convert/:tool           │
 │ JPG to PDF, PDF to JPG / PPT / Markdown   │ ─────► │   LibreOffice / MS Office         │
 │ (files stay on your device)               │ upload │   Chromium, Ghostscript, Poppler  │
 │                                           │ ◄───── │ GET  /api/health                  │
 │ Word/PPT/Excel/Markdown merge in own      │ result │ Temp files deleted after each job │
 │ format using JSZip + XML                  │        └───────────────────────────────────┘
 └───────────────────────────────────────────┘
```

### Typical flow for every tool
1. **Select**: choose a tool.
2. **Upload**: drag and drop or pick files. Each file is checked (type, size, empty, damaged, password protected).
3. **Settings**: choose output name and tool options.
4. **Download**: the result card shows size, a preview, and a download button.

The shared stepper, file list and result card live in `js/flow.js`, and each tool only supplies its own settings and `run()` function.

### Browser-side tools
| Task | Technique |
|---|---|
| Merge / split PDF | `pdf-lib` copies pages between documents |
| Merge Word / PowerPoint | `JSZip` opens the Office files, XML is edited with the DOM, and styles, numbering, images, relationships and slide layouts are remapped into the first file |
| Merge Excel | `SheetJS` (xlsx) |
| Compress images | The image is drawn on a canvas and re-encoded. Quality and size are searched until it fits the target, never below a clarity floor |
| Compress PDF | Only the pictures inside are re-encoded with `pdf-lib`. Text and vectors are untouched |
| Compress Word / PowerPoint / Excel | Pictures in the `media` folder are re-encoded, then the file is re-zipped |
| PDF to JPG / PowerPoint | `pdf.js` renders each page to a canvas; `PptxGenJS` builds slides |
| PDF to Word / Markdown (browser path) | `pdf.js` reads text positions, rows, headings, lists and tables are rebuilt; `Tesseract.js` runs OCR on scanned pages |
| Preview | PDF pages via `pdf.js`, Word via `docx-preview`, PowerPoint via a built-in slide renderer, Markdown via `marked` |

### How "compress to a size" works
1. The app compresses the file once at the lowest clear setting to find the **minimum possible size**.
2. When you type a target, it runs a **binary search** over quality and scale, trying a setting, measuring the real output size, and narrowing in.
3. It picks the best-looking setting whose file size is at or below your target.

### Server-side tools
`POST /api/convert/:tool` receives the file, runs the right program, returns the result, and deletes temporary files.

| Tool key | What runs |
|---|---|
| `to-word` | Microsoft Word (PowerShell, Windows) or LibreOffice |
| `to-ppt` | Microsoft PowerPoint (PowerShell, Windows) or LibreOffice |
| `to-html`, `to-md` | Chrome / Edge headless "print to PDF" |
| `from-word` | LibreOffice PDF import to DOCX (optional OCR) |
| `from-pdfa` | Ghostscript, else LibreOffice PDF/A export |
| `from-jpg`, `from-md`, `from-ppt` | Poppler / LibreOffice (optional paths) |

The page asks `GET /api/health` to learn which programs exist, then chooses the best engine, falling back to a browser version where one exists.

### External libraries (loaded from CDN)
`pdf-lib`, `JSZip`, `pdf.js`, `Tesseract.js`, `PptxGenJS`, `SheetJS`, `docx-preview`, `marked`.

---

## Project structure

```
merge-and-convert/
├── index.html            Page markup, favicon links, script includes
├── css/styles.css        Design system, light/dark theme, mobile layout
├── js/
│   ├── utils.js          Helpers, icons, file-type tiles, toasts, dialogs
│   ├── flow.js           Shared stepper, upload list, settings, result card
│   ├── helpers.js        PDF/text/Excel merge, server API call, health check
│   ├── frompdf.js        pdf.js reading, OCR, PDF to Word/JPG/PPT/Markdown
│   ├── merge.js          Merge section
│   ├── officemerge.js    Word and PowerPoint merge
│   ├── split.js          Split PDF section
│   ├── compress.js       Compress section
│   ├── convert.js        Convert to / from PDF sections, PDF/A fallback
│   ├── preview.js        Result preview (PDF, Word, PowerPoint, Markdown)
│   └── ui.js             Header, search, theme, navigation
├── server/server.js      Conversion API
├── Dockerfile            Image with every converter installed
├── favicon.svg, favicon.ico, icon-*.png, apple-touch-icon.png, site.webmanifest
├── workflow.md           How the app works
└── pdr.md                Product requirements
```

---

## Server API

| Method | Path | Purpose |
|---|---|---|
| `GET` | `/api/health` | Reports which tools the server found: `soffice`, `gs`, `chrome`, `ocr`, `word`, `powerpoint` |
| `POST` | `/api/convert/:tool` | Multipart upload: `file` plus `opts` (JSON string). Returns the converted file or `{ "error": "..." }` |

Example:

```bash
curl -F "file=@report.docx" -F 'opts={}' http://localhost:3000/api/convert/to-word -o report.pdf
```

---

## Configuration

| Setting | Where | Default |
|---|---|---|
| Port | `PORT` environment variable | `3000` |
| LibreOffice path | `SOFFICE_PATH` | auto-detected |
| Chrome / Edge path | `CHROME_PATH` | auto-detected |
| Server URL when opening `index.html` directly | `window.API_URL` in `index.html` | `http://localhost:3000` |
| Max file size | `MAX` in `js/utils.js` and the multer limit in `server/server.js` | 100 MB |

---

## Privacy and limits

- **Browser tools** (merge, split, compress, JPG to PDF, PDF to JPG / PowerPoint / Markdown, previews) process files locally. Nothing is uploaded.
- **Server tools** upload the file to the server you run, process it in a temporary folder, and delete it afterwards. If you host it yourself, the files stay on your own machine or network.
- Maximum file size is 100 MB per file.
- Password-protected PDFs cannot be processed. Remove the password first.
- Old Office formats (.doc, .ppt, .xls) are not supported. Save them as .docx, .pptx or .xlsx.
- Preview is a quick view. Fonts missing on your device may look different in the preview, but the downloaded file is unchanged.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| "Cannot reach the conversion server" | Start it with `npm start` and open the app at `http://localhost:3000` |
| "LibreOffice is not installed on this computer" | Install LibreOffice and restart the server, or use Docker |
| "Ghostscript is not installed" for PDF/A | Install Ghostscript or LibreOffice. Without both, the browser makes a best-effort PDF/A that is not validated |
| Word to PDF fonts look different | Install the fonts used by the document on the server computer, or use Microsoft Word on Windows |
| Page is blank or tools do nothing | Check your internet connection (CDN libraries) and the browser console |
| Port already in use | Start with another port: `PORT=8080 npm start` |
| Phone cannot open the app | Open the computer's address (for example `http://192.168.1.10:3000`) on the same Wi-Fi |

---

## Customizing

- **Icons:** edit the `IC` object in `js/utils.js`. Each icon is SVG drawing code on a 24×24 grid.
- **Logo and favicon:** edit `favicon.svg`, then re-export the PNG and ICO files.
- **Colors and phone layout:** `css/styles.css`. The phone rules are at the bottom, under the 640 px media query.
- **Add a new tool:** add an entry to the `TO` or `FROM` list in `js/convert.js`, a handler in `server/server.js` (`H`), and optionally a browser version in the `CLIENT` object.

---

## License

Add your license here (for example MIT).
