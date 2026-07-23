# CellSnip — DataSnipper-style snipping add-in for Excel

A lightweight Office Add-in that replicates the core DataSnipper workflow: open PDF support documents in a task pane next to your workbook, drag a box over the document, and the value lands in the selected cell with a permanent, traceable link back to the exact spot in the source.

## Features

| Mode | What it does |
|---|---|
| **Text snip** | Extracts the text in the box into the selected cell (single numbers are written as real numbers, with $ / commas / parentheses handled) |
| **Sum snip** | Finds every number in the box and writes the total; the component values are recorded in the snip record and cell comment |
| **Table snip** (beta) | Splits the boxed text into rows/columns and writes a 2D block starting at the selected cell — works on digital PDFs with a text layer |
| **Table Template + Table Match** | Draw a Row Key column (e.g. cost codes on a G703 Schedule of Values) plus one box per target column once; **Table Match** then takes a column of keys already in Excel (e.g. your cost code list), finds each one on the document — per-page OCR fallback included for scanned schedules of values — and fills in the matching column values next to each key, row by row |
| **Form Template (anchored)** | Draw fields once on a sample document (e.g. a G702/G703 pay app or a recurring vendor invoice); each field is anchored to its nearby label, not a fixed coordinate, so **Extract Forms** re-finds it on every other document with the same layout, even scanned ones (falls back to a whole-page OCR pass when there's no text layer, so it works on scanned pay apps/invoices, not just digital PDFs). Fields whose anchor can't be found on a given document fall back to the original position and the output cell is flagged (amber fill + comment note) so you know to spot-check it |
| **Validation snip (✓)** | Marks the selected cell green as "agreed to source" without overwriting its value |
| **Exception snip (✗)** | Marks the cell red and prompts for an exception note |
| **Go to source** | Select any snipped cell → jumps the viewer to the source document, page, and highlighted region |
| **OCR fallback** | Toggle on for scanned documents — runs Tesseract on the boxed region (needs internet for language data on first use) |
| **Export overview** | Generates a "Snip Overview" worksheet documenting every snip (cell, type, value, document, page, note, timestamp) — your workpaper documentation tab |
| **Persistence** | Snip records live on a hidden `_CellSnip` worksheet, so they travel **with the workbook**. PDFs are cached in the browser (IndexedDB) and auto-reload next session |

Every snipped cell also gets a color code by snip type and a cell comment with the source reference (Excel desktop / recent versions).

## Files

- `manifest.xml` — the add-in manifest you sideload into Excel
- `taskpane.html` — the entire app (single file; PDF.js + Tesseract.js from CDN)
- `icon-16/32/80.png` — ribbon icons

## Deployment (one-time, ~10 minutes)

Office add-ins must be served over **HTTPS**. The easiest zero-infrastructure option is GitHub Pages, which you already use:

### 1. Host the files on GitHub Pages
1. Create a repo (e.g. `cellsnip`), add `taskpane.html` and the three icons.
2. Repo → Settings → Pages → Deploy from branch → `main` / root.
3. Your app is now at `https://YOURUSERNAME.github.io/cellsnip/taskpane.html`.

### 2. Point the manifest at it
Find/replace `https://YOUR-HOST-URL` in `manifest.xml` with `https://YOURUSERNAME.github.io/cellsnip` (no trailing slash).

### 3. Sideload into Excel

**Excel for Windows (shared-folder catalog):**
1. Create a folder, e.g. `C:\AddIns`, and share it with yourself: right-click → Properties → Sharing → Share → add your own account → note the network path (`\\YOURPC\AddIns`).
2. Put `manifest.xml` in that folder.
3. Excel → File → Options → Trust Center → Trust Center Settings → **Trusted Add-in Catalogs** → paste the network path → check "Show in Menu" → OK → restart Excel.
4. Home ribbon → Add-ins → **More Add-ins** (or Insert → My Add-ins) → **Shared Folder** tab → CellSnip.

**Excel on the web (fastest to test):** Home → Add-ins → More Add-ins → **My Add-ins → Upload My Add-in** → pick `manifest.xml`.

### Local development alternative
If you'd rather iterate locally before publishing:
```bash
npm i -g office-addin-dev-certs http-server
office-addin-dev-certs install
http-server . -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key -p 3000
```
Then set the manifest URLs to `https://localhost:3000` and sideload the same way.

## Usage

1. Open the task pane, click **+ PDF**, load one or more support documents (invoices, bank statements, loan portals, draw detail…).
2. Click the cell in Excel where the value should go.
3. Pick a snip mode, drag a box on the document. The value writes to the cell and the selection moves down one row so you can rip through a column of support quickly.
4. Later: select any snipped cell → **Go to source** highlights exactly where it came from.
5. Before archiving: **Export overview** to generate the documentation sheet.

## Known limits vs. DataSnipper

- **Document matching / auto-reconcile** (matching a whole ledger against a stack of PDFs automatically) is not implemented — snips are manual.
- Table snip (the freehand rows/columns mode) needs a digital text layer; column detection is heuristic, so verify the output.
- Form Template fields anchor to nearby label text (left of the box, or above for table-style columns); a field with no distinguishable nearby label still falls back to its fixed drawn position. On scanned documents, the anchor label itself has to be read back correctly by OCR to be found again — turn on **OCR fallback** before building *and* before running Extract Forms on scanned forms.
- Table Match/Table Template column x-positions are still fixed per template (not anchored to a header the way Form Template fields are) — it assumes every document uses the same column layout, which holds for repeated scans of the same form (e.g. monthly G703s from the same GC) but not for a mix of different vendors' schedules of values.
- The G703 auto-detect template builder (scans a digital text layer to guess column positions automatically) doesn't have an OCR path yet — on a scanned G703, build the Table Template manually instead (draw the Row Key column + each target column once; this takes a few seconds and doesn't need any text extraction to succeed).
- Both Table Match and Extract Forms OCR a full page once (cached) rather than each cell/field individually, but a large scanned document with many pages can still take a while the first time — subsequent extractions against the same loaded document reuse the cached OCR text.
- No confidence/review-queue UI beyond the per-cell amber flag on Extract Forms output — there's no batch summary of low-confidence fields yet.
- PDFs are cached per machine/browser, not embedded in the workbook. If a colleague opens the file, snip records and the overview travel with it, but they'd need the source PDFs to use "Go to source." (Embedding PDFs in the workbook is possible via base64 in the hidden sheet but bloats file size fast — happy to add it as an option.)
- Cell comments require Excel 2019+/365 (ExcelApi 1.10); on older builds the snip still works, just without the comment.

## Where things are stored

- Hidden worksheet `_CellSnip` (unhide via right-click sheet tabs if you ever want to inspect/audit it): one row per snip with cell, type, value, document, page, PDF coordinates, extracted text, timestamp.
- Browser IndexedDB database `cellsnip`: cached PDF bytes keyed by filename.
