# Workflow

## Pages and links
The home page shows the intro, quick buttons and one card per tool. Each tool opens on its own page and has its own address (hash routes, so no server setup is needed):
`#home`, `#merge`, `#merge/docx`, `#split`, `#compress`, `#topdf`, `#topdf/word`, `#frompdf`, `#frompdf/jpg`. The router is `route()` at the end of js/ui.js. The sidebar, header, search, home cards and the browser Back button all use it.

## User workflow
Every tool follows Select -> Upload -> Settings -> Download.
1. Select: pick a merge type, open a Split PDF, or open a Convert tool card.
2. Upload: only the drop zone is shown. Drag files in or press Select files. Wrong type, empty, damaged, protected or oversized files get an error on their row.
3. Settings: appears only after a valid file is added (or text is pasted for HTML and Markdown). Set the output name and the tool options. In Split PDF only the fields for the chosen mode are shown: Page ranges, Pages per file (N), or the page grid.
4. Download: press Preview to check the result first, or the main button to run. The result card shows the file and a Download button.

## Where each job runs
- In the browser (pdf-lib, SheetJS, JSZip, pdf.js, PptxGenJS): merge PDF, Word, PowerPoint (each in its own format), Images, Markdown, TXT, CSV, Excel; Split PDF; JPG to PDF; PDF to JPG, Word, PowerPoint and Markdown; Compress to a target size. No installs needed.
- On the server: Word, PowerPoint, HTML and Markdown to PDF (needs Chrome, Edge or Chromium; uses LibreOffice instead when it is installed, for better layout), and PDF to PDF/A (needs Ghostscript).

## Developer workflow
1. Install the converters (or use Docker) and run `npm install`, then `npm start`.
2. Open http://localhost:3000. The server also serves the frontend.
3. Edit files: css/styles.css for design, js/flow.js for the shared stepper and result card, js/merge.js, split.js, convert.js for the tools, server/server.js for conversions.
4. To add a tool: add it to the TO or FROM list in js/convert.js, add a handler under the same key in `H` in server/server.js, and add its output extension to `EXT`.

## Error handling
- A missing converter returns a message naming what to install.
- The result card shows Try again and a link to Help.
- Uploaded files are removed from the server temp folder as soon as the response is sent.

## Deploy
Build the Dockerfile and run it on any container host. Set PORT if needed.
