# PDF Splitter & Renamer — Project Context for Claude CLI

## Overview

This is a **CD iQs** internal tooling project. CD iQs is a quantity surveying consultancy.
The developer is a civil/structural engineer running a QS business, building estimating
software for quantity surveyors, builders, and trade contractors.

The immediate problem being solved: importing drawing sets into **CostX** (estimating software)
requires manually renaming every PDF sheet — a process that takes hours per week. This tool
automates that by scraping metadata from drawing title blocks and pre-naming the files.

---

## What Already Existed — `pdf_viewer.js` + `index.html`

A working static web app (GitHub Pages: https://jimlakiss.github.io/PDF-Viewer/) that:

- Loads single or multiple PDFs using **PDF.js** (v3.11.174)
- Lets the user draw bounding box regions on page 1 to define where title block fields are
- Template regions ghost-propagate to all pages automatically
- Runs **multi-pass Tesseract.js OCR** (v5.0.4) on each region with field-specific preprocessing:
  - Adaptive thresholding for ID-like fields (`sheet_id`, `issue_id`)
  - Sharpening kernel for description fields
  - 0/O disambiguation logic for sheet IDs
- Extracts fields: `sheet_id`, `description`, `issue_id`, `date`, `issue_description`
  plus document-level: `prepared_by`, `project_id`
- Exports to **JSON** and **CSV**

The scraper is production-quality and works well. The region coordinates are stored as
normalised fractions (0–1) so they work across different page sizes/zoom levels.

### Key globals exposed by `pdf_viewer.js`

```javascript
pdfDoc               // PDF.js document object (or virtual multi-doc wrapper)
pdfRawBytes          // Uint8Array of the loaded PDF — NEEDS PATCH (see below)
multiPdfDocs         // Array of { doc, pageOffset, fileName, pageCount, rawBytes }
isMultiPdfMode       // boolean
documentDetails      // { prepared_by, project_id }
sheetDetailsByPage   // { [pageNum]: { sheet_id, description, issue_id, date, issue_description } }
regionTemplates      // { [fieldName]: { x, y, w, h } } — normalised coords
applyTemplatesToAllPages(logProgress) // async — runs OCR on every page
getCanonicalExportData()              // returns { document: {}, sheets: [] }
showProcessing(message)               // shows loading overlay
hideProcessing()                      // hides loading overlay
downloadBlob(blob, filename)          // triggers browser download
extractAll()                          // extracts document-level fields only
downloadJSON()                        // exports raw extraction as JSON
downloadCSV()                         // exports raw extraction as CSV
```

---

## What Was Built in This Session

### New feature: PDF Splitter & Renamer

A new standalone static page (`splitter.html` + `pdf_splitter.js`) that:

1. Reuses ALL of the existing `pdf_viewer.js` scraping machinery unchanged
2. After extraction, opens a **staging area** — a full-screen overlay panel with an
   editable table showing every page's scraped metadata and proposed filename
3. Lets the user review, edit inline, reorder via drag-and-drop, then download

### Files produced

| File | Status | Notes |
|------|--------|-------|
| `splitter.html` | ✅ New | Standalone page — same viewer UI + staging panel |
| `pdf_splitter.js` | ✅ New | All staging/split/export logic |
| `pdf_viewer_patch_notes.txt` | ✅ New | 3-line patch instructions for existing `pdf_viewer.js` |
| `pdf_viewer.js` | ⚠️ Needs patch | 3 lines to add — see patch notes |

---

## The 3-Line Patch Required in `pdf_viewer.js`

**This patch has NOT been applied yet.** The existing `pdf_viewer.js` needs these changes:

### Patch 1 — Add global (~line 33, after `let pdfDoc = null;`)
```javascript
let pdfRawBytes = null;   // raw Uint8Array — used by pdf-lib for splitting
```

### Patch 2 — Store bytes in single-file mode (~line 283, inside `reader.onload`)
```javascript
// FIND:
const data = new Uint8Array(reader.result);

// ADD AFTER:
pdfRawBytes = data;
```

### Patch 3 — Store bytes per file in multi-file mode (~line 382, inside `loadMultiplePDFs`)
```javascript
// FIND:
multiPdfDocs.push({
  doc: doc,
  pageOffset: pageOffset,
  fileName: file.name,
  pageCount: doc.numPages,
});

// REPLACE WITH:
multiPdfDocs.push({
  doc: doc,
  pageOffset: pageOffset,
  fileName: file.name,
  pageCount: doc.numPages,
  rawBytes: uint8Data,   // ADD THIS LINE ONLY
});
```

---

## `splitter.html` Structure

The HTML must provide these element IDs for `pdf_viewer.js` to bind to:

```
#file-input       — file input (multiple attribute set)
#pdf-canvas       — main canvas element
#sidebar          — thumbnail sidebar div
#page-indicator   — page counter span
#zoom-in          — zoom in button
#zoom-out         — zoom out button
#pdf-scroll       — scrollable container div
#overlay          — SVG element overlaid on canvas (for region drawing)
#region-type      — select for region type
#draw-type-swatch — span showing current region type colour
#prepared-by      — text input
#project-id       — text input
```

And these for `pdf_splitter.js`:

```
#btn-split-name       — main "Split & Name" trigger button
#staging-panel        — full-screen overlay div (display:none initially)
#sp-template-select   — filename template preset dropdown
#sp-custom-row        — div containing custom template input (hidden initially)
#sp-custom-input      — custom template text input
#sp-seq-toggle        — sequence prefix checkbox
#sp-seq-label         — span that turns blue when seq is on
#sp-seq-sep-row       — div containing separator select (hidden initially)
#sp-seq-sep           — separator select (" - " or "_")
#sp-filter-select     — filter dropdown (all/issues/ready)
#sp-select-all        — select-all checkbox in table header
#sp-tbody             — table body
#sp-stats             — stats text (N ready · N needs review…)
#sp-btn-zip           — download ZIP button
#sp-btn-csv           — export CSV button
#sp-btn-json          — export JSON button
#sp-btn-close         — close staging panel button
#sp-btn-reset         — reset row order button
```

### Canvas + Overlay structure
The SVG overlay must be positioned absolutely over the canvas:
```html
<div id="pdf-scroll">
  <div id="canvas-wrap">
    <canvas id="pdf-canvas"></canvas>
    <svg id="overlay"></svg>  <!-- position:absolute top:0 left:0 -->
  </div>
</div>
```

---

## CDN Dependencies (load order matters)

```html
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script src="https://unpkg.com/tesseract.js@5.0.4/dist/tesseract.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script src="pdf_viewer.js"></script>
<script src="pdf_splitter.js"></script>
```

**Load order is critical:** pdf.js and tesseract must be loaded before `pdf_viewer.js`.
pdf-lib and jszip must be loaded before `pdf_splitter.js`.

---

## `pdf_splitter.js` Feature Summary

### Staging area
- Builds from `getCanonicalExportData()` after `applyTemplatesToAllPages()` runs
- Each row: checkbox · page# · sheet_id · description · issue_id · filename preview · status badge
- All cells inline-editable; filename preview updates live as you type
- Status: **Ready** (all key fields present) / **Review** (partial) / **OCR failed** (all blank)

### Drag-and-drop row reordering
- `draggable="true"` on each `<tr>`
- ⠿ grip handle in leftmost column
- Blue insertion line (border-top/bottom) shows drop position
- Reorders the `stagingData` array on drop
- `↺ Reset order` button restores original page-number sequence

### Sequence prefix toggle
- Off by default
- When on: prepends `001 - `, `002 - ` etc. to every filename
- Forces CostX import order regardless of alphabetical sort
- Separator configurable: ` - ` or `_`
- Updates live across all filename previews

### Filename template
- 3 presets + custom mode
- Tokens: `{sheet_id}` `{description}` `{issue_id}` `{date}` `{issue_description}` `{project_id}` `{prepared_by}`
- Empty tokens collapse cleanly (no double-dashes)

### Download ZIP
- Uses **pdf-lib** to extract individual pages from source PDF
- For loose files (multi-file mode, 1-page sources): copies bytes unchanged — no re-encoding
- For bound PDFs: `PDFDocument.copyPages()` per page
- Filename in ZIP = staged filename (post-edit, post-reorder, with seq prefix if on)

### CSV export (manifest)
- Columns: `order, page, filename, sheet_id, description, issue_id, date, issue_description, status`
- `order` = position in staging table (post-drag), `page` = source PDF page number

### JSON export (manifest)
- Same fields as CSV
- Includes `exported_at`, `project_id`, `prepared_by` at document level

---

## Known Issue — Nothing Loading

The user reports the page is not loading. Likely causes to check:

1. **Are all 4 CDN scripts loading?** Check Network tab in DevTools for 404s or blocked requests.

2. **Is `pdf_viewer.js` throwing on load?** It references DOM elements by ID on load
   (e.g. `document.getElementById("file-input")`). If any expected ID is missing from
   `splitter.html`, it will throw and stop execution.

3. **`file://` protocol issue:** Tesseract.js workers cannot fetch language data files
   when the page is opened directly from the filesystem. Must be served via `http://`.
   Fix: `python3 -m http.server 8080` in the project folder.

4. **Script load order:** If `pdf_splitter.js` loads before `pdf_viewer.js` defines its
   globals, the `init()` function will find no elements to bind to. Check script tag order.

5. **Missing element IDs:** `pdf_viewer.js` was written for the original `index.html`.
   `splitter.html` must have exactly matching element IDs or the JS will silently fail.
   Cross-reference the ID list above against the actual HTML.

6. **The `pdfRawBytes` patch:** If `pdf_viewer.js` hasn't been patched, `pdfRawBytes`
   will be `undefined` when `pdf_splitter.js` tries to use it for PDF splitting.
   This won't stop the page loading but will cause ZIP download to fail.

---

## Repository

- GitHub: https://github.com/jimlakiss/PDF-Viewer
- Live (existing viewer): https://jimlakiss.github.io/PDF-Viewer/
- Target (new splitter): https://jimlakiss.github.io/PDF-Viewer/splitter.html
- Branch: `PDF-Plan-Data`

## Tech Stack

- Pure static HTML/JS — no build step, no framework, no backend
- Hosted on GitHub Pages
- PDF.js 3.11.174 (rendering + text extraction)
- Tesseract.js 5.0.4 (OCR with multi-pass + preprocessing)
- pdf-lib 1.17.1 (PDF page splitting)
- JSZip 3.10.1 (ZIP assembly)