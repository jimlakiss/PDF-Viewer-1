// pdf_viewer.js - COMPLETE LATEST VERSION
// Multi-page PDF viewer + region drawing + vector/OCR extraction + multi-file support

const ENABLE_VECTOR_EXTRACTION = false;
const DOCUMENT_DETAILS = ["prepared_by", "project_id"];
const REGION_TYPES = ["sheet_id", "description", "issue_id", "date", "issue_description"];

const fileInput = document.getElementById("file-input");
const uploadDropZone = document.getElementById("upload-drop-zone");
const canvas = document.getElementById("pdf-canvas");
const ctx = canvas.getContext("2d");
const sidebar = document.getElementById("sidebar");
const prevPageBtn  = document.getElementById("prev-page");
const nextPageBtn  = document.getElementById("next-page");
const pageInputEl  = document.getElementById("page-input");
const pageTotalEl  = document.getElementById("page-total");
const zoomInputEl  = document.getElementById("zoom-input");
const zoomInBtn    = document.getElementById("zoom-in");
const zoomOutBtn   = document.getElementById("zoom-out");
const fitWidthBtn  = document.getElementById("fit-width");
const pdfScroll   = document.getElementById("pdf-scroll");
const canvasOuter = document.getElementById("canvas-outer");
const overlay     = document.getElementById("overlay");
const regionTypeSelect = document.getElementById("region-type");
const drawTypeSwatch = document.getElementById("draw-type-swatch");
const preparedByInput = document.getElementById("prepared-by");
const projectIdInput = document.getElementById("project-id");

pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
if (pdfjsLib.GlobalWorkerOptions) {
  pdfjsLib.GlobalWorkerOptions.enableXfa = false;
}

let TEMPLATE_MASTER_PAGE = null;
let pdfDoc = null;
let pdfRawBytes = null;   // raw Uint8Array — used by pdf-lib for splitting
let currentPage = 1;
let scale = 1.5;
let pdfFileBaseName = "pdf_extracted_data";
let multiPdfDocs = [];
let isMultiPdfMode = false;

const documentDetails = { prepared_by: "", project_id: "" };
const sheetDetailsByPage = {};
const regionsByPage = {};
const regionTemplates = {};

let selectedRegionIds = [];
let selectedRegionId = null;
let clipboardRegions = [];
let clipboardBase = null;
let clipboardPasteSerial = 0;
let regionIdCounter = 1;
let lastPointerNorm = null;

let ocrWorkerPromise = null;
let _ocrJobChain = Promise.resolve();

let isDrawing = false;
let startX = 0;
let startY = 0;
let activeRect = null;

let isDraggingRegions = false;
let dragStartPx = { x: 0, y: 0 };
let dragHasMoved = false;
let dragClickShouldToggleOff = false;
let dragStartById = new Map();
let currentRenderTask = null;
const MAX_CANVAS_DIMENSION = 16384;
let renderedScale = scale;   // scale at which the visible canvas was last drawn
let zoomDebounceTimer = null;
let gestureAnchor = null;    // page-space point locked under the pointer for this gesture
let canvasPadX = 0;          // canvas-outer left/right padding (px) — set by scrollToCenter
let canvasPadY = 0;          // canvas-outer top/bottom padding (px)
let recenterAfterRender = false;
let currentTool = 'select'; // 'select' | 'scale-zone' | 'linear' | 'area' | 'count'
const thumbnailDataCache = new Map();

const ZOOM_SNAP = [0.10, 0.20, 0.25, 0.33, 0.50, 0.67, 0.75, 1.00, 1.25, 1.50, 2.00, 2.50, 3.00, 4.00, 5.00];

function updateControls() {
  if (pageInputEl) pageInputEl.value = pdfDoc ? currentPage : '—';
  if (pageTotalEl) pageTotalEl.textContent = pdfDoc ? pdfDoc.numPages : '—';
  if (zoomInputEl) zoomInputEl.value = pdfDoc ? `${Math.round(scale * 100)}%` : '—';
}

// ghostExclusions variable
const ghostExclusions = {}; // { pageNum: Set(['sheet_id', ...]) }
// Undo/Redo system
const undoStack = [];
const redoStack = [];
const MAX_UNDO_STACK = 50;

function saveUndoState() {
  // Save current state
  const state = {
    regionsByPage: JSON.parse(JSON.stringify(regionsByPage)),
    sheetDetailsByPage: JSON.parse(JSON.stringify(sheetDetailsByPage)),
    regionTemplates: JSON.parse(JSON.stringify(regionTemplates)),
    selectedRegionIds: [...selectedRegionIds],
    currentPage: currentPage,
  };
  
  undoStack.push(state);
  
  // Limit stack size
  if (undoStack.length > MAX_UNDO_STACK) {
    undoStack.shift();
  }
  
  // Clear redo stack when new action is performed
  redoStack.length = 0;
  
  console.log(`💾 Undo state saved (stack: ${undoStack.length})`);
}

function undo() {
  if (undoStack.length === 0) {
    console.log('❌ Nothing to undo');
    return;
  }
  
  // Save current state to redo stack
  const currentState = {
    regionsByPage: JSON.parse(JSON.stringify(regionsByPage)),
    sheetDetailsByPage: JSON.parse(JSON.stringify(sheetDetailsByPage)),
    regionTemplates: JSON.parse(JSON.stringify(regionTemplates)),
    selectedRegionIds: [...selectedRegionIds],
    currentPage: currentPage,
  };
  redoStack.push(currentState);
  
  // Restore previous state
  const prevState = undoStack.pop();
  
  Object.keys(regionsByPage).forEach(k => delete regionsByPage[k]);
  Object.assign(regionsByPage, prevState.regionsByPage);
  
  Object.keys(sheetDetailsByPage).forEach(k => delete sheetDetailsByPage[k]);
  Object.assign(sheetDetailsByPage, prevState.sheetDetailsByPage);
  
  Object.keys(regionTemplates).forEach(k => delete regionTemplates[k]);
  Object.assign(regionTemplates, prevState.regionTemplates);
  
  selectedRegionIds.length = 0;
  selectedRegionIds.push(...prevState.selectedRegionIds);
  syncLegacySelectedId();
  
  currentPage = prevState.currentPage;
  
  console.log(`↶ Undo applied (undo: ${undoStack.length}, redo: ${redoStack.length})`);
  
  renderPage(currentPage);
}

function redo() {
  if (redoStack.length === 0) {
    console.log('❌ Nothing to redo');
    return;
  }
  
  // Save current state to undo stack
  const currentState = {
    regionsByPage: JSON.parse(JSON.stringify(regionsByPage)),
    sheetDetailsByPage: JSON.parse(JSON.stringify(sheetDetailsByPage)),
    regionTemplates: JSON.parse(JSON.stringify(regionTemplates)),
    selectedRegionIds: [...selectedRegionIds],
    currentPage: currentPage,
  };
  undoStack.push(currentState);
  
  // Restore next state
  const nextState = redoStack.pop();
  
  Object.keys(regionsByPage).forEach(k => delete regionsByPage[k]);
  Object.assign(regionsByPage, nextState.regionsByPage);
  
  Object.keys(sheetDetailsByPage).forEach(k => delete sheetDetailsByPage[k]);
  Object.assign(sheetDetailsByPage, nextState.sheetDetailsByPage);
  
  Object.keys(regionTemplates).forEach(k => delete regionTemplates[k]);
  Object.assign(regionTemplates, nextState.regionTemplates);
  
  selectedRegionIds.length = 0;
  selectedRegionIds.push(...nextState.selectedRegionIds);
  syncLegacySelectedId();
  
  currentPage = nextState.currentPage;
  
  console.log(`↷ Redo applied (undo: ${undoStack.length}, redo: ${redoStack.length})`);
  
  renderPage(currentPage);
}

async function getOcrWorker() {
  if (ocrWorkerPromise) return ocrWorkerPromise;
  if (!window.Tesseract?.createWorker) {
    throw new Error("Tesseract.createWorker not available");
  }
  const workerOptions = {
    workerPath: "https://unpkg.com/tesseract.js@5.0.4/dist/worker.min.js",
    corePath: "https://unpkg.com/tesseract.js-core@5.0.0/tesseract-core-simd.wasm.js",
    langPath: "https://tessdata.projectnaptha.com/4.0.0",
  };
  ocrWorkerPromise = (async () => {
    let worker;
    try {
      worker = await Tesseract.createWorker("eng", 1, workerOptions);
      return worker;
    } catch (_) {}
    worker = await Tesseract.createWorker(workerOptions);
    await worker.loadLanguage("eng");
    await worker.initialize("eng");
    return worker;
  })();
  return ocrWorkerPromise;
}

function runOcrExclusive(fn) {
  _ocrJobChain = _ocrJobChain.then(fn, fn);
  return _ocrJobChain;
}

(function initRegionTypeSelect() {
  if (!regionTypeSelect) return;
  regionTypeSelect.innerHTML = "";
  const ogDoc = document.createElement("optgroup");
  ogDoc.label = "DOCUMENT_DETAILS";
  DOCUMENT_DETAILS.forEach((type) => {
    const opt = document.createElement("option");
    opt.value = type;
    opt.textContent = type;
    ogDoc.appendChild(opt);
  });
  const ogSheet = document.createElement("optgroup");
  ogSheet.label = "REGION_TYPES";
  REGION_TYPES.forEach((type) => {
    const opt = document.createElement("option");
    opt.value = type;
    opt.textContent = type;
    ogSheet.appendChild(opt);
  });
  regionTypeSelect.appendChild(ogDoc);
  regionTypeSelect.appendChild(ogSheet);
  function syncDrawTypeSwatch() {
    if (!drawTypeSwatch) return;
    drawTypeSwatch.setAttribute("data-swatch", regionTypeSelect.value);
  }
  regionTypeSelect.addEventListener("change", syncDrawTypeSwatch);
  syncDrawTypeSwatch();
})();

async function handleSelectedPDFs(selectedFiles, source = "picker") {
  const files = Array.from(selectedFiles || []);
  
  console.log(`🔍 PDF FILES SELECTED (${source}):`);
  console.log(`   - files array length:`, files.length);
  console.log(`   - files array:`, files.map(f => f.name));
  
  if (!files.length) {
    console.log('❌ No files selected');
    return;
  }

  const nonPdfFiles = files.filter(file => !/\.pdf$/i.test(file.name || '') && file.type !== 'application/pdf');
  if (nonPdfFiles.length) {
    alert(`Only PDF files can be uploaded. Remove: ${nonPdfFiles.map(file => file.name).join(', ')}`);
    if (fileInput) fileInput.value = '';
    return;
  }

  if (files.length > 1) {
    console.log(`📚 MULTI-FILE MODE: Loading ${files.length} PDF files...`);
    await loadMultiplePDFs(files);
    return;
  }

  console.log('📂 SINGLE-FILE MODE');
  const file = files[0];
  const fileSizeMB = file.size / (1024 * 1024);
  
  if (fileSizeMB > 20) {
    console.warn(`⚠️ Large file detected (${fileSizeMB.toFixed(1)} MB)`);
    if (!confirm(`This is a large file (${fileSizeMB.toFixed(1)} MB). Continue?`)) {
      if (fileInput) fileInput.value = '';
      return;
    }
  }

  pdfFileBaseName = (file.name || "pdf_extracted_data").replace(/\.[^.]+$/, "") || "pdf_extracted_data";

  pdfDoc = null;
  multiPdfDocs = [];
  isMultiPdfMode = false;
  currentPage = 1;
  scale = 1.5;
  selectedRegionIds = [];
  selectedRegionId = null;
  regionIdCounter = 1;

  for (const k of Object.keys(documentDetails)) documentDetails[k] = "";
  for (const k of Object.keys(sheetDetailsByPage)) delete sheetDetailsByPage[k];
  for (const k of Object.keys(regionsByPage)) delete regionsByPage[k];
  for (const k of Object.keys(regionTemplates)) delete regionTemplates[k];

  if (preparedByInput) preparedByInput.value = "";
  if (projectIdInput) projectIdInput.value = "";

  const reader = new FileReader();
  reader.onload = async () => {
    try {
      const data = new Uint8Array(reader.result);
      pdfRawBytes = data.slice(); // copy before PDF.js transfers the ArrayBuffer to its worker
      console.log('📂 Loading PDF document...');

      const loadingTask = pdfjsLib.getDocument({
        data: data,
        cMapUrl: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/cmaps/',
        cMapPacked: true,
        disableAutoFetch: true,
        disableStream: false,
        disableFontFace: false,
        useSystemFonts: false,
        standardFontDataUrl: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/standard_fonts/',
      });
      
      pdfDoc = await loadingTask.promise;
      const pageCount = pdfDoc.numPages;
      console.log(`✅ PDF loaded: ${pageCount} pages (${fileSizeMB.toFixed(1)} MB)`);
      
      if (pageCount === 0) {
        throw new Error('PDF has 0 pages');
      }
      
      try {
        const testPage = await pdfDoc.getPage(1);
        console.log(`📄 Page 1 dimensions: ${testPage.view[2]} x ${testPage.view[3]}`);
        testPage.cleanup();
      } catch (err) {
        console.error('❌ Cannot read first page:', err);
        throw new Error('PDF appears corrupted');
      }
      
      if (pageCount === 1 && file.size > 1000000) {
        console.warn('⚠️ Large file but only 1 page detected');
      }
      
      console.log(`🖼️ Starting thumbnail build for ${pageCount} pages...`);
      await buildThumbnails();
      console.log('📄 Rendering first page...');
      recenterAfterRender = true;
      await renderPage(1);
      console.log('✅ PDF ready');
    } catch (err) {
      console.error('❌ Error loading PDF:', err);
      alert(`Failed to load PDF: ${err.message}`);
      if (fileInput) fileInput.value = '';
    }
  };
  
  reader.onerror = () => {
    console.error('❌ Error reading file');
    alert('Failed to read the PDF file');
    if (fileInput) fileInput.value = '';
  };
  
  reader.readAsArrayBuffer(file);
}

fileInput?.addEventListener("change", async (e) => {
  await handleSelectedPDFs(e.target.files, "picker");
});

function setUploadDragActive(active) {
  uploadDropZone?.classList.toggle("is-drag-over", active);
}

["dragenter", "dragover"].forEach(eventName => {
  uploadDropZone?.addEventListener(eventName, e => {
    e.preventDefault();
    e.stopPropagation();
    e.dataTransfer.dropEffect = "copy";
    setUploadDragActive(true);
  });
});

["dragleave", "drop"].forEach(eventName => {
  uploadDropZone?.addEventListener(eventName, e => {
    e.preventDefault();
    e.stopPropagation();
    setUploadDragActive(false);
  });
});

uploadDropZone?.addEventListener("drop", async e => {
  const files = Array.from(e.dataTransfer?.files || []);
  await handleSelectedPDFs(files, "drop");
});

async function loadMultiplePDFs(files) {
  console.log(`📚 Loading ${files.length} separate PDF files as one combined document...`);
  
  pdfDoc = null;
  multiPdfDocs = [];
  isMultiPdfMode = true;
  currentPage = 1;
  scale = 1.5;
  selectedRegionIds = [];
  selectedRegionId = null;
  regionIdCounter = 1;

  for (const k of Object.keys(documentDetails)) documentDetails[k] = "";
  for (const k of Object.keys(sheetDetailsByPage)) delete sheetDetailsByPage[k];
  for (const k of Object.keys(regionsByPage)) delete regionsByPage[k];
  for (const k of Object.keys(regionTemplates)) delete regionTemplates[k];

  if (preparedByInput) preparedByInput.value = "";
  if (projectIdInput) projectIdInput.value = "";

  pdfFileBaseName = "combined_pdfs";
  let totalPages = 0;
  let pageOffset = 0;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    console.log(`  📄 Loading file ${i + 1}/${files.length}: ${file.name}`);

    try {
      const data = await file.arrayBuffer();
      const uint8Data = new Uint8Array(data);
      const uint8Copy = uint8Data.slice(); // copy before PDF.js transfers the ArrayBuffer to its worker

      const loadingTask = pdfjsLib.getDocument({
        data: uint8Data,
        cMapUrl: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/cmaps/',
        cMapPacked: true,
        disableAutoFetch: true,
        disableStream: false,
        disableFontFace: false,
      });

      const doc = await loadingTask.promise;
      
      multiPdfDocs.push({
        doc: doc,
        pageOffset: pageOffset,
        fileName: file.name,
        pageCount: doc.numPages,
        rawBytes: uint8Copy,
      });

      console.log(`    ✅ ${file.name}: ${doc.numPages} page(s)`);
      totalPages += doc.numPages;
      pageOffset += doc.numPages;

    } catch (err) {
      console.error(`    ❌ Failed to load ${file.name}:`, err);
      alert(`Failed to load ${file.name}: ${err.message}`);
      fileInput.value = '';
      return;
    }
  }

  console.log(`✅ All PDFs loaded: ${totalPages} total pages from ${files.length} files`);

  pdfDoc = {
    numPages: totalPages,
    getPage: async (pageNum) => {
      for (const pdf of multiPdfDocs) {
        const localPageNum = pageNum - pdf.pageOffset;
        if (localPageNum >= 1 && localPageNum <= pdf.pageCount) {
          return await pdf.doc.getPage(localPageNum);
        }
      }
      throw new Error(`Page ${pageNum} not found`);
    },
  };

  console.log(`🖼️ Building thumbnails for ${totalPages} combined pages...`);
  await buildThumbnails();
  console.log('📄 Rendering first page...');
  recenterAfterRender = true;
  await renderPage(1);
  console.log('✅ Combined PDF ready');
}

preparedByInput?.addEventListener("input", () => {
  documentDetails.prepared_by = preparedByInput.value || "";
});

projectIdInput?.addEventListener("input", () => {
  documentDetails.project_id = projectIdInput.value || "";
});

async function renderPage(pageNum) {
  if (!pdfDoc) return;

  if (currentRenderTask) {
    currentRenderTask.cancel();
    currentRenderTask = null;
  }

  currentPage = pageNum;
  selectedRegionIds = [];
  selectedRegionId = null;
  msrActiveDrawPts = [];
  msrPreviewPt = null;
  msrSelectedId = null;
  szRefState = null;

  let page;
  try {
    page = await pdfDoc.getPage(pageNum);
  } catch (err) {
    console.error(`Error loading page ${pageNum}:`, err);
    return;
  }

  let effectiveScale = scale;

  const baseViewport = page.getViewport({ scale: 1.0 });
  pageBaseDimsCache.set(pageNum, { width: baseViewport.width, height: baseViewport.height });
  const targetWidth  = baseViewport.width  * scale;
  const targetHeight = baseViewport.height * scale;

  if (targetWidth > MAX_CANVAS_DIMENSION || targetHeight > MAX_CANVAS_DIMENSION) {
    const scaleX = MAX_CANVAS_DIMENSION / baseViewport.width;
    const scaleY = MAX_CANVAS_DIMENSION / baseViewport.height;
    effectiveScale = Math.min(scaleX, scaleY, scale);
    // Cap scale so renderedScale stays in sync — mismatch breaks the CSS
    // zoom ratio and causes scroll jumps near the canvas size limit.
    scale = effectiveScale;
  }

  const viewport = page.getViewport({ scale: effectiveScale });

  // Render into a hidden offscreen canvas so the visible canvas
  // never shows a blank/loading state — swap happens atomically.
  const offscreen = document.createElement('canvas');
  offscreen.width  = viewport.width;
  offscreen.height = viewport.height;

  try {
    currentRenderTask = page.render({
      canvasContext: offscreen.getContext('2d'),
      viewport,
      intent: 'display',
      renderInteractiveForms: false,
      enableWebGL: false,
    });
    await currentRenderTask.promise;
    currentRenderTask = null;
  } catch (err) {
    offscreen.width = 0;
    offscreen.height = 0;
    if (err.name !== 'RenderingCancelledException') {
      console.error(`❌ Error rendering page ${pageNum}:`, err);
    }
    if (page?.cleanup) page.cleanup();
    return;
  }

  // Atomic swap — order matters to prevent scroll-area contraction:
  // Set canvas.width FIRST while the CSS override is still active (layout
  // unchanged), THEN remove the override so layout transitions directly from
  // CSS-scaled size → new pixel size (≈ same value, no contraction).
  // Save/restore scroll so any sub-pixel rounding difference can't clamp it.
  const savedScrollLeft = pdfScroll.scrollLeft;
  const savedScrollTop  = pdfScroll.scrollTop;
  canvas.width  = viewport.width;
  canvas.height = viewport.height;
  canvas.style.width  = '';
  canvas.style.height = '';
  pdfScroll.scrollLeft = savedScrollLeft;
  pdfScroll.scrollTop  = savedScrollTop;
  ctx.drawImage(offscreen, 0, 0);
  offscreen.width = 0;
  offscreen.height = 0;

  overlay.setAttribute('width',  viewport.width);
  overlay.setAttribute('height', viewport.height);

  renderedScale = scale;

  if (recenterAfterRender) {
    recenterAfterRender = false;
    scrollToCenter();
  }

  updateControls();

  highlightActiveThumb();
  redrawRegions();

  if (page?.cleanup) page.cleanup();
  page = null;
}

// Sets padding on canvas-outer so there is scroll room in every direction,
// then scrolls to put the canvas in the centre of the viewport.
function scrollToCenter() {
  // One full viewport-width of padding on each side gives the scroll
  // container enough range to zoom toward any pointer location.
  canvasPadX = pdfScroll.clientWidth;
  canvasPadY = pdfScroll.clientHeight;
  canvasOuter.style.padding = `${canvasPadY}px ${canvasPadX}px`;

  // Force layout so scrollWidth/Height reflect the new padding.
  void pdfScroll.scrollWidth;

  pdfScroll.scrollLeft = Math.max(0, canvasPadX + Math.round((canvas.width  - pdfScroll.clientWidth)  / 2));
  pdfScroll.scrollTop  = Math.max(0, canvasPadY + Math.round((canvas.height - pdfScroll.clientHeight) / 2));
}

zoomInBtn?.addEventListener("click", () => {
  if (!pdfDoc) return;
  const next = ZOOM_SNAP.find(s => s > scale + 0.001);
  zoomToCenter(next !== undefined ? next : ZOOM_SNAP[ZOOM_SNAP.length - 1]);
});

zoomOutBtn?.addEventListener("click", () => {
  if (!pdfDoc) return;
  const prev = [...ZOOM_SNAP].reverse().find(s => s < scale - 0.001);
  zoomToCenter(prev !== undefined ? prev : ZOOM_SNAP[0]);
});

async function zoomToCenter(newScale) {
  if (!pdfDoc) return;
  const anchorPageX = (pdfScroll.scrollLeft + pdfScroll.clientWidth  / 2 - canvasPadX) / scale;
  const anchorPageY = (pdfScroll.scrollTop  + pdfScroll.clientHeight / 2 - canvasPadY) / scale;
  scale = Math.min(Math.max(newScale, 0.01), 20.0);
  await renderPage(currentPage);
  pdfScroll.scrollLeft = anchorPageX * scale + canvasPadX - pdfScroll.clientWidth  / 2;
  pdfScroll.scrollTop  = anchorPageY * scale + canvasPadY - pdfScroll.clientHeight / 2;
}

fitWidthBtn?.addEventListener("click", async () => {
  if (!pdfDoc) return;
  const page = await pdfDoc.getPage(currentPage);
  const baseViewport = page.getViewport({ scale: 1.0 });
  page.cleanup();
  scale = pdfScroll.clientWidth / baseViewport.width;
  scale = Math.min(Math.max(scale, 0.01), 20.0);
  recenterAfterRender = true;
  renderPage(currentPage);
});

prevPageBtn?.addEventListener("click", () => {
  if (!pdfDoc || currentPage <= 1) return;
  recenterAfterRender = true;
  renderPage(currentPage - 1);
});

nextPageBtn?.addEventListener("click", () => {
  if (!pdfDoc || currentPage >= pdfDoc.numPages) return;
  recenterAfterRender = true;
  renderPage(currentPage + 1);
});

function commitPageInput() {
  if (!pdfDoc) return;
  const n = parseInt(pageInputEl.value, 10);
  if (!Number.isFinite(n) || n < 1 || n > pdfDoc.numPages) {
    updateControls();
    return;
  }
  if (n === currentPage) return;
  recenterAfterRender = true;
  renderPage(n);
}

pageInputEl?.addEventListener("keydown", (e) => {
  if (e.key === "Enter") { e.preventDefault(); pageInputEl.blur(); commitPageInput(); }
  if (e.key === "Escape") { updateControls(); pageInputEl.blur(); }
});
pageInputEl?.addEventListener("blur", commitPageInput);

function commitZoomInput() {
  if (!pdfDoc) return;
  const raw = (zoomInputEl.value || "").replace(/%/g, "").trim();
  const pct = parseFloat(raw);
  if (!Number.isFinite(pct) || pct < 1 || pct > 2000) {
    updateControls();
    return;
  }
  zoomToCenter(pct / 100);
}

zoomInputEl?.addEventListener("keydown", (e) => {
  if (e.key === "Enter") { e.preventDefault(); zoomInputEl.blur(); commitZoomInput(); }
  if (e.key === "Escape") { updateControls(); zoomInputEl.blur(); }
});
zoomInputEl?.addEventListener("blur", commitZoomInput);

async function buildThumbnails() {
  if (!pdfDoc || !sidebar) {
    console.error('❌ buildThumbnails: pdfDoc or sidebar is null!');
    return;
  }

  const numPages = pdfDoc.numPages;
  console.log(`🖼️ Building thumbnails for ${numPages} pages...`);

  sidebar.innerHTML = "";
  
  for (let i = 1; i <= numPages; i++) {
    const placeholder = document.createElement("div");
    placeholder.classList.add("thumb");
    placeholder.style.cssText = `
  background: #555;
  color: #fff;
  font-size: 11px;
`;
    placeholder.textContent = `Page ${i}`;
    placeholder.dataset.pageNum = i;
    placeholder.addEventListener("click", () => {
      console.log(`👆 Clicked thumbnail ${i}`);
      recenterAfterRender = true;
      renderPage(i);
    });
    sidebar.appendChild(placeholder);
  }
  
  const placeholderCount = sidebar.querySelectorAll('.thumb').length;
  console.log(`✅ Created ${placeholderCount} placeholder thumbnails`);
  
  if (placeholderCount !== numPages) {
    console.error(`❌ Mismatch! Expected ${numPages}, got ${placeholderCount}`);
  }
  
  highlightActiveThumb();
  
  if (numPages > 10) {
    console.log('📄 Large PDF - lazy loading thumbnails');
    for (let i = 1; i <= Math.min(3, numPages); i++) {
      await loadThumbnail(i);
    }
    return;
  }
  
  console.log('📄 Loading all thumbnails...');
  for (let i = 1; i <= numPages; i++) {
    await loadThumbnail(i);
    if (i % 3 === 0) {
      await new Promise(resolve => setTimeout(resolve, 0));
    }
  }

  console.log('✅ All thumbnails loaded');
}

async function loadThumbnail(pageNum) {
  if (thumbnailDataCache.has(pageNum)) return;
  
  try {
    const page = await pdfDoc.getPage(pageNum);
    const THUMB_SCALE = 0.12;
    const viewport = page.getViewport({ scale: THUMB_SCALE });

    const tempCanvas = document.createElement("canvas");
    const maxWidth = 180;
    const scaleFactor = Math.min(1, maxWidth / viewport.width);
    tempCanvas.width = viewport.width * scaleFactor;
    tempCanvas.height = viewport.height * scaleFactor;
    
    const tempCtx = tempCanvas.getContext("2d", { alpha: false });

    await page.render({ 
      canvasContext: tempCtx, 
      viewport: page.getViewport({ scale: tempCanvas.width / page.getViewport({ scale: 1 }).width }),
      intent: 'display',
      renderInteractiveForms: false,
    }).promise;

    const dataUrl = tempCanvas.toDataURL('image/jpeg', 0.7);
    thumbnailDataCache.set(pageNum, dataUrl);
    
    const placeholder = sidebar.querySelector(`div.thumb[data-page-num="${pageNum}"]`);
    if (placeholder) {
      const img = document.createElement('img');
      img.src = dataUrl;
      img.style.cssText = 'width: 100%; height: auto; display: block;';
      placeholder.innerHTML = '';
      placeholder.appendChild(img);
      placeholder.style.background = '#444';
    } else {
      console.warn(`Could not find placeholder for page ${pageNum}`);
    }
    
    page.cleanup();
    tempCanvas.width = 0;
    tempCanvas.height = 0;
    
  } catch (err) {
    console.error(`Error loading thumbnail ${pageNum}:`, err);
  }
}

function highlightActiveThumb() {
  const thumbs = sidebar.querySelectorAll(".thumb");
  thumbs.forEach((t, i) => {
    const isActive = (i + 1) === currentPage;
    if (isActive) {
      t.style.borderColor = '#4da3ff';
      const pageNum = i + 1;
      if (pdfDoc && pdfDoc.numPages > 10) {
        loadThumbnail(pageNum);
        if (pageNum > 1) loadThumbnail(pageNum - 1);
        if (pageNum < pdfDoc.numPages) loadThumbnail(pageNum + 1);
      }
    } else {
      t.style.borderColor = 'transparent';
    }
  });
}

function getOverlayPoint(evt) {
  const r = overlay.getBoundingClientRect();
  return { x: evt.clientX - r.left, y: evt.clientY - r.top };
}

function updateLastPointerNormFromEvent(evt) {
  if (!overlay) return;
  const r = overlay.getBoundingClientRect();
  const xPx = evt.clientX - r.left;
  const yPx = evt.clientY - r.top;
  if (xPx < 0 || yPx < 0 || xPx > r.width || yPx > r.height) return;
  const x = r.width ? xPx / r.width : 0;
  const y = r.height ? yPx / r.height : 0;
  const cx = x < 0 ? 0 : x > 1 ? 1 : x;
  const cy = y < 0 ? 0 : y > 1 ? 1 : y;
  lastPointerNorm = { x: cx, y: cy };
}

window.addEventListener("mousemove", updateLastPointerNormFromEvent, { passive: true });

function beginRegionDrag(evt, clickShouldToggleOff) {
  if (!overlay) return;
  saveUndoState(); // Save state BEFORE dragging starts
  isDraggingRegions = true;
  dragHasMoved = false;
  dragClickShouldToggleOff = !!clickShouldToggleOff;
  dragStartPx = getOverlayPoint(evt);
  dragStartById = new Map();
  const regs = regionsByPage[currentPage] || [];
  const sel = new Set(selectedRegionIds);
  regs.forEach((r) => {
    if (sel.has(r.id)) {
      dragStartById.set(r.id, { x: r.x, y: r.y, w: r.w, h: r.h });
    }
  });
  window.addEventListener("mousemove", onRegionDragMove, true);
  window.addEventListener("mouseup", onRegionDragEnd, true);
}

function onRegionDragMove(evt) {
  if (!isDraggingRegions) return;
  if (!canvas.width || !canvas.height) return;
  const p = getOverlayPoint(evt);
  const dxPx = p.x - dragStartPx.x;
  const dyPx = p.y - dragStartPx.y;
  if (!dragHasMoved && (Math.abs(dxPx) > 2 || Math.abs(dyPx) > 2)) {
    dragHasMoved = true;
  }
  let dx = dxPx / canvas.width;
  let dy = dyPx / canvas.height;
  let minDx = -Infinity, maxDx = Infinity;
  let minDy = -Infinity, maxDy = Infinity;
  dragStartById.forEach(({ x, y, w, h }) => {
    minDx = Math.max(minDx, -x);
    maxDx = Math.min(maxDx, (1 - w) - x);
    minDy = Math.max(minDy, -y);
    maxDy = Math.min(maxDy, (1 - h) - y);
  });
  dx = Math.min(Math.max(dx, minDx), maxDx);
  dy = Math.min(Math.max(dy, minDy), maxDy);
  const regs = regionsByPage[currentPage] || [];
  const byId = new Map(regs.map(r => [r.id, r]));
  dragStartById.forEach((s, id) => {
  const r = byId.get(id);
  if (!r) return;
  
  // If dragging a ghost, convert to real region FIRST with new values
  if (r.isGhost) {
    // Create a proper copy, don't modify the template
    const newX = s.x + dx;
    const newY = s.y + dy;
    
    r.x = newX;
    r.y = newY;
    delete r.isGhost;
    console.log(`👻→✓ Converting ghost "${r.type}" to real region at (${(newX*100).toFixed(1)}%, ${(newY*100).toFixed(1)}%)`);
  } else {
    r.x = s.x + dx;
    r.y = s.y + dy;
  }
});

  redrawRegions();
}

function onRegionDragEnd() {
  if (!isDraggingRegions) return;
  window.removeEventListener("mousemove", onRegionDragMove, true);
  window.removeEventListener("mouseup", onRegionDragEnd, true);
  const shouldToggleOff = dragClickShouldToggleOff && !dragHasMoved;
  if (dragHasMoved) {
    const regs = regionsByPage[currentPage] || [];
    const byId = new Map(regs.map(r => [r.id, r]));
    const changedTypes = [...new Set(selectedRegionIds.map(id => byId.get(id)?.type).filter(Boolean))];
    invalidatePageFields(currentPage, changedTypes);
  }
  isDraggingRegions = false;
  dragClickShouldToggleOff = false;
  dragStartById = new Map();
  if (shouldToggleOff) {
    clearSelection();
  }
  redrawRegions();
}

overlay?.addEventListener("mousedown", (e) => {
  if (currentTool !== 'select') return;
  if (e.target?.tagName === "rect") return;
  if (isDraggingRegions) return;
  if (isDrawing && activeRect) return;
  isDrawing = true;
  selectedRegionIds = [];
  selectedRegionId = null;
  const r = overlay.getBoundingClientRect();
  startX = e.clientX - r.left;
  startY = e.clientY - r.top;
  activeRect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
  activeRect.setAttribute("fill", "rgba(0, 123, 255, 0.2)");
  activeRect.setAttribute("stroke", "#007bff");
  activeRect.setAttribute("stroke-width", "2");
  overlay.appendChild(activeRect);
});

overlay?.addEventListener("mousemove", (e) => {
  if (currentTool !== 'select') return;
  if (isDraggingRegions) return;
  if (!isDrawing || !activeRect) return;
  const r = overlay.getBoundingClientRect();
  const x = e.clientX - r.left;
  const y = e.clientY - r.top;
  activeRect.setAttribute("x", Math.min(startX, x));
  activeRect.setAttribute("y", Math.min(startY, y));
  activeRect.setAttribute("width", Math.abs(x - startX));
  activeRect.setAttribute("height", Math.abs(y - startY));
});

overlay?.addEventListener("mouseup", () => {
  if (currentTool !== 'select') return;
  if (isDraggingRegions) return;
  if (!isDrawing || !activeRect) return;
  const x = +(activeRect.getAttribute("x") || 0);
  const y = +(activeRect.getAttribute("y") || 0);
  const w = +(activeRect.getAttribute("width") || 0);
  const h = +(activeRect.getAttribute("height") || 0);
  activeRect.remove();
  isDrawing = false;
  activeRect = null;
  if (w < 2 || h < 2) {
    return;
  }
  const region = {
    id: regionIdCounter++,
    type: regionTypeSelect?.value || REGION_TYPES[0],
    x: x / canvas.width,
    y: y / canvas.height,
    w: w / canvas.width,
    h: h / canvas.height,
  };
  saveUndoState(); // Save state BEFORE making changes

if (!regionsByPage[currentPage]) regionsByPage[currentPage] = [];
regionsByPage[currentPage].push(region);
invalidatePageFields(currentPage, [region.type]);
if (REGION_TYPES.includes(region.type)) {
  if (TEMPLATE_MASTER_PAGE === null) {
    TEMPLATE_MASTER_PAGE = currentPage;
    console.log(`📐 Master page set to: ${TEMPLATE_MASTER_PAGE}`);
    promoteRegionToTemplate(region);
    console.log(`✨ Auto-promoted "${region.type}" to template (first region of this type)`);
  } else if (currentPage === TEMPLATE_MASTER_PAGE) {
    // On master page - update template
    promoteRegionToTemplate(region);
    console.log(`✨ Updated template for "${region.type}" (drawing on master page)`);
  } else {
    // Not on master page - create page-specific override only
    console.log(`📍 Created page-specific override for "${region.type}" on page ${currentPage} (not master page)`);
  }
}
  redrawRegions();
});

function msrEnsureGroups() {
  let zonesG = overlay.querySelector('.msr-zones-g');
  let regionsG = overlay.querySelector('.msr-regions-g');
  let measG = overlay.querySelector('.msr-meas-g');
  if (!regionsG) {
    overlay.innerHTML = '';
    zonesG   = document.createElementNS("http://www.w3.org/2000/svg", 'g');
    regionsG = document.createElementNS("http://www.w3.org/2000/svg", 'g');
    measG    = document.createElementNS("http://www.w3.org/2000/svg", 'g');
    zonesG.classList.add('msr-zones-g');
    regionsG.classList.add('msr-regions-g');
    measG.classList.add('msr-meas-g');
    overlay.appendChild(zonesG);
    overlay.appendChild(regionsG);
    overlay.appendChild(measG);
  }
  return { zonesG, regionsG, measG };
}

function redrawRegions() {
  if (!overlay) return;
  const { zonesG, regionsG, measG } = msrEnsureGroups();
  zonesG.innerHTML   = '';
  regionsG.innerHTML = '';
  measG.innerHTML    = '';

  msrRenderZones(zonesG);

  // Clear previous ghost regions from current page before redrawing
  const pageRegions = regionsByPage[currentPage] || [];
  // Don't clear ghosts that are currently selected (being dragged)
const selectedSet = new Set(selectedRegionIds);
regionsByPage[currentPage] = pageRegions.filter(r => !r.isGhost || selectedSet.has(r.id));
  
  // Get fresh list after filtering ghosts
  const currentPageRegions = regionsByPage[currentPage] || [];
  
  // Add ghost regions from templates (if they don't have page-specific overrides)
  Object.values(regionTemplates).forEach((tpl) => {
  // Check if this ghost is excluded on this page
  if (ghostExclusions[currentPage] && ghostExclusions[currentPage].has(tpl.type)) {
    console.log(`🚫 Page ${currentPage}: Ghost "${tpl.type}" is excluded`);
    return;
  }
  
  // Check if we already have this ghost in selection
  const existingGhost = currentPageRegions.find(r => r.isGhost && r.type === tpl.type);
  if (existingGhost) {
    console.log(`♻️ Keeping existing ghost "${tpl.type}" (ID: ${existingGhost.id})`);
    return; // Keep the existing one, don't create a new one
  }
  
  const hasOverride = currentPageRegions.some((r) => r.type === tpl.type && !r.isGhost);
  if (hasOverride) {
    console.log(`⏭️ Page ${currentPage}: Skipping ghost for "${tpl.type}" (has real region)`);
    return;
  }

  // Create editable ghost region from template
  const ghostRegion = {
    id: regionIdCounter++,
    type: tpl.type,
    x: tpl.x,
    y: tpl.y,
    w: tpl.w,
    h: tpl.h,
    isGhost: true,
  };

  console.log(`👻 Page ${currentPage}: Adding ghost for "${tpl.type}"`);
  currentPageRegions.push(ghostRegion);
});
  
  // Draw all regions (real + ghosts)
  currentPageRegions.forEach((r) => {
    const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    rect.setAttribute("x", r.x * canvas.width);
    rect.setAttribute("y", r.y * canvas.height);
    rect.setAttribute("width", r.w * canvas.width);
    rect.setAttribute("height", r.h * canvas.height);
    rect.dataset.id = String(r.id);
    rect.dataset.type = r.type;
    
    // Visual indicator for ghost regions
    if (r.isGhost) {
      rect.setAttribute("stroke-dasharray", "6 4");
      rect.setAttribute("opacity", "0.7");
    }
    
    if (selectedRegionIds.includes(r.id)) rect.classList.add("selected");
    
    rect.addEventListener("mousedown", (e) => {
      if (currentTool !== 'select') return;
      e.stopPropagation();
      if (e.shiftKey) {
        toggleSelection(r.id);
        redrawRegions();
        return;
      }
      const preWasSelected = selectedRegionIds.includes(r.id);
      const preWasSingleSame = (selectedRegionIds.length === 1 && selectedRegionIds[0] === r.id);
      const preWasMulti = selectedRegionIds.length > 1;
      if (!preWasSelected) {
        setSingleSelection(r.id);
        redrawRegions();
        beginRegionDrag(e, false);
        return;
      }
      if (preWasMulti) {
        beginRegionDrag(e, false);
        redrawRegions();
        return;
      }
      beginRegionDrag(e, preWasSingleSame);
      redrawRegions();
    });
    regionsG.appendChild(rect);
  });

  msrRenderMeasurements(measG);
  msrUpdateScaleLabel();
}

function syncLegacySelectedId() {
  selectedRegionId = selectedRegionIds.length > 0 ? selectedRegionIds[selectedRegionIds.length - 1] : null;
}

function setSingleSelection(id) {
  selectedRegionIds = [id];
  syncLegacySelectedId();
}

function toggleSelection(id) {
  const idx = selectedRegionIds.indexOf(id);
  if (idx >= 0) {
    selectedRegionIds.splice(idx, 1);
  } else {
    selectedRegionIds.push(id);
  }
  syncLegacySelectedId();
}

function clearSelection() {
  selectedRegionIds = [];
  selectedRegionId = null;
}

function isEditableTarget(el) {
  if (!el) return false;
  const tag = (el.tagName || "").toUpperCase();
  return tag === "INPUT" || tag === "TEXTAREA" || el.isContentEditable;
}

function clamp01(v) {
  if (Number.isNaN(v)) return 0;
  if (v < 0) return 0;
  if (v > 1) return 1;
  return v;
}

function getSelectedRegionsOnCurrentPage() {
  const regions = regionsByPage[currentPage] || [];
  if (!selectedRegionIds.length) return [];
  const sel = new Set(selectedRegionIds);
  return regions.filter((r) => sel.has(r.id));
}

function copySelectionToClipboard() {
  const regs = getSelectedRegionsOnCurrentPage();
  if (!regs.length) return false;
  const minX = Math.min(...regs.map((r) => r.x));
  const minY = Math.min(...regs.map((r) => r.y));
  clipboardBase = { minX, minY };
  clipboardRegions = regs.map((r) => ({
    type: r.type,
    dx: r.x - minX,
    dy: r.y - minY,
    w: r.w,
    h: r.h,
  }));
  clipboardPasteSerial = 0;
  console.log(`📋 Copied ${clipboardRegions.length} region(s)`);
  return true;
}

function pasteClipboardToCurrentPage() {
  if (!clipboardRegions.length || !clipboardBase) return false;
  const changedTypes = [...new Set(clipboardRegions.map(r => r.type))];
  invalidatePageFields(currentPage, changedTypes);
  clipboardPasteSerial += 1;
  const nudgePx = 10 * clipboardPasteSerial;
  const nudgeX = canvas.width ? nudgePx / canvas.width : 0.01;
  const nudgeY = canvas.height ? nudgePx / canvas.height : 0.01;
  const baseX = clipboardBase.minX + nudgeX;
  const baseY = clipboardBase.minY + nudgeY;
  if (!regionsByPage[currentPage]) regionsByPage[currentPage] = [];
  const newIds = [];
  clipboardRegions.forEach((it) => {
    const id = regionIdCounter++;
    let x = baseX + it.dx;
    let y = baseY + it.dy;
    x = clamp01(x);
    y = clamp01(y);
    x = Math.min(x, 1 - it.w);
    y = Math.min(y, 1 - it.h);
    const newRegion = {
      id,
      type: it.type,
      x,
      y,
      w: it.w,
      h: it.h,
    };
    regionsByPage[currentPage].push(newRegion);
    if (REGION_TYPES.includes(newRegion.type)) {
      if (TEMPLATE_MASTER_PAGE === null) {
        TEMPLATE_MASTER_PAGE = currentPage;
        promoteRegionToTemplate(newRegion);
      } else if (currentPage === TEMPLATE_MASTER_PAGE) {
        promoteRegionToTemplate(newRegion);
      }
    }
    newIds.push(id);
  });
  selectedRegionIds = newIds;
  syncLegacySelectedId();
  redrawRegions();
  console.log(`📋 Pasted ${newIds.length} region(s) onto page ${currentPage}`);
  return true;
}

function invalidatePageFields(pageNum, types) {
  if (!sheetDetailsByPage[pageNum]) return;
  types.forEach(t => {
    if (t in sheetDetailsByPage[pageNum]) {
      delete sheetDetailsByPage[pageNum][t];
      console.log(`🔄 Invalidated cache for page ${pageNum}, field "${t}"`);
    }
  });
}

function pasteClipboardToCurrentPageAtPointer() {
  if (!clipboardRegions.length || !clipboardBase) return false;
  if (!lastPointerNorm) return pasteClipboardToCurrentPage();
  const changedTypes = [...new Set(clipboardRegions.map(r => r.type))];
  invalidatePageFields(currentPage, changedTypes);
  let baseX = lastPointerNorm.x;
  let baseY = lastPointerNorm.y;
  let groupMaxX = 0;
  let groupMaxY = 0;
  clipboardRegions.forEach((it) => {
    groupMaxX = Math.max(groupMaxX, it.dx + it.w);
    groupMaxY = Math.max(groupMaxY, it.dy + it.h);
  });
  baseX = Math.min(Math.max(baseX, 0), Math.max(0, 1 - groupMaxX));
  baseY = Math.min(Math.max(baseY, 0), Math.max(0, 1 - groupMaxY));
  if (!regionsByPage[currentPage]) regionsByPage[currentPage] = [];
  const newIds = [];
  clipboardRegions.forEach((it) => {
    const id = regionIdCounter++;
    const x = baseX + it.dx;
    const y = baseY + it.dy;
    const newRegion = {
      id,
      type: it.type,
      x,
      y,
      w: it.w,
      h: it.h,
    };
    regionsByPage[currentPage].push(newRegion);
    if (REGION_TYPES.includes(newRegion.type)) {
      if (TEMPLATE_MASTER_PAGE === null) {
        TEMPLATE_MASTER_PAGE = currentPage;
        promoteRegionToTemplate(newRegion);
      } else if (currentPage === TEMPLATE_MASTER_PAGE) {
        promoteRegionToTemplate(newRegion);
      }
    }
    newIds.push(id);
  });
  selectedRegionIds = newIds;
  syncLegacySelectedId();
  redrawRegions();
  console.log(`📋 Pasted ${newIds.length} region(s) at pointer onto page ${currentPage}`);
  return true;
}

function getMostRecentRegionOfType(pageNum, type) {
  const regions = regionsByPage[pageNum] || [];
  for (let i = regions.length - 1; i >= 0; i--) {
    if (regions[i].type === type) return regions[i];
  }
  return null;
}

function resolveRegionForPage(pageNum, type) {
  const pageRegions = regionsByPage[pageNum] || [];
  const override = [...pageRegions].reverse().find((r) => r.type === type);
  if (override) return override;
  
  // Check if this ghost is excluded on this page
  if (ghostExclusions[pageNum] && ghostExclusions[pageNum].has(type)) {
    console.log(`⏭️ Page ${pageNum}: Skipping extraction for excluded ghost "${type}"`);
    return null; // Don't extract if ghost was deleted
  }
  
  if (regionTemplates[type]) return regionTemplates[type];
  return null;
}

function resolveAllRegionsForPage(pageNum, type) {
  const pageRegions = regionsByPage[pageNum] || [];
  const overrides = pageRegions.filter((r) => r.type === type);
  
  if (overrides.length > 0) {
    // Sort regions top-to-bottom, then left-to-right
    return overrides.sort((a, b) => {
      const yDiff = a.y - b.y;
      if (Math.abs(yDiff) > 0.01) return yDiff; // Different rows
      return a.x - b.x; // Same row, sort by x
    });
  }
  
  // Check if this ghost is excluded on this page
  if (ghostExclusions[pageNum] && ghostExclusions[pageNum].has(type)) {
    console.log(`⏭️ Page ${pageNum}: Skipping extraction for excluded ghost "${type}"`);
    return []; // Don't extract if ghost was deleted
  }
  
  if (regionTemplates[type]) return [regionTemplates[type]];
  return [];
}

function promoteRegionToTemplate(region) {
  if (!region || !region.type) return;
  regionTemplates[region.type] = {
    type: region.type,
    x: region.x,
    y: region.y,
    w: region.w,
    h: region.h,
  };
  console.log(`📐 Template set for "${region.type}" at (${(region.x * 100).toFixed(1)}%, ${(region.y * 100).toFixed(1)}%) - will appear on all pages`);
}

function isIdLikeField(field) {
  return field === "sheet_id" || field === "issue_id" || field === "project_id" || field === "date";
}

function preprocessOcrCrop(canvasEl, field) {
  try {
    const c = canvasEl.getContext("2d");
    const w = canvasEl.width;
    const h = canvasEl.height;
    const img = c.getImageData(0, 0, w, h);
    const d = img.data;

    for (let i = 0; i < d.length; i += 4) {
      const g = (d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114) | 0;
      d[i] = d[i + 1] = d[i + 2] = g;
    }

    let min = 255, max = 0;
    for (let i = 0; i < d.length; i += 4) {
      const v = d[i];
      if (v < min) min = v;
      if (v > max) max = v;
    }
    
    const range = max - min;
    if (range > 10) {
      for (let i = 0; i < d.length; i += 4) {
        const stretched = ((d[i] - min) * 255 / range) | 0;
        d[i] = d[i + 1] = d[i + 2] = stretched;
      }
    }

    if (isIdLikeField(field)) {
      const blockSize = 15;
      const C = 10;
      const original = new Uint8ClampedArray(d);
      
      for (let y = 0; y < h; y++) {
        for (let x = 0; x < w; x++) {
          const idx = (y * w + x) * 4;
          let sum = 0;
          let count = 0;
          const halfBlock = Math.floor(blockSize / 2);
          
          for (let by = Math.max(0, y - halfBlock); by <= Math.min(h - 1, y + halfBlock); by++) {
            for (let bx = Math.max(0, x - halfBlock); bx <= Math.min(w - 1, x + halfBlock); bx++) {
              sum += original[(by * w + bx) * 4];
              count++;
            }
          }
          
          const localMean = sum / count;
          const threshold = localMean - C;
          const v = original[idx] > threshold ? 255 : 0;
          d[idx] = d[idx + 1] = d[idx + 2] = v;
        }
      }
    } else {
      const sharpen = [-1, -1, -1, -1, 9, -1, -1, -1, -1];
      const original = new Uint8ClampedArray(d);
      
      for (let y = 1; y < h - 1; y++) {
        for (let x = 1; x < w - 1; x++) {
          let sum = 0;
          for (let ky = -1; ky <= 1; ky++) {
            for (let kx = -1; kx <= 1; kx++) {
              const idx = ((y + ky) * w + (x + kx)) * 4;
              sum += original[idx] * sharpen[(ky + 1) * 3 + (kx + 1)];
            }
          }
          const idx = (y * w + x) * 4;
          const v = Math.max(0, Math.min(255, sum));
          d[idx] = d[idx + 1] = d[idx + 2] = v;
        }
      }
    }

    c.putImageData(img, 0, 0);
  } catch (err) {
    console.warn("Preprocessing failed:", err);
  }
}

function _normText(s) {
  return String(s || "").replace(/\s+/g, " ").trim();
}

/**
 * Disambiguates between digit 0 and letter O based on surrounding context
 * Uses heuristics: if surrounded by digits, likely a 0; if by letters, likely an O
 */
function disambiguate0AndO(text) {
  if (!text) return text;

  let result = text;

  // Replace ambiguous characters based on neighbors
  result = result.replace(/[O0]/g, (match, offset) => {
    const prev = offset > 0 ? text[offset - 1] : '';
    const next = offset < text.length - 1 ? text[offset + 1] : '';

    const prevIsDigit = /[1-9]/.test(prev);
    const nextIsDigit = /[1-9]/.test(next);
    const prevIsLetter = /[A-Za-z]/.test(prev);
    const nextIsLetter = /[A-Za-z]/.test(next);

    // Strong evidence for digit 0
    if (prevIsDigit && nextIsDigit) return '0';
    if (prevIsDigit && !nextIsLetter) return '0';
    if (nextIsDigit && !prevIsLetter) return '0';

    // Strong evidence for letter O
    if (prevIsLetter && nextIsLetter) return 'O';
    if (prevIsLetter && !nextIsDigit) return 'O';
    if (nextIsLetter && !prevIsDigit) return 'O';

    // Default: keep as-is or prefer based on position
    // At start of string: if next is letter, use O; if next is digit, use 0
    if (offset === 0) {
      if (nextIsLetter) return 'O';
      if (nextIsDigit) return '0';
    }

    // Keep original character if no clear context
    return match;
  });

  return result;
}

function cleanByField(field, raw) {
  const t = _normText(raw);
  if (!t) return "";

  if (field === "date") {
    const m = t.match(/\b(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})\b/);
    return m ? m[1] : t;
  }

  if (field === "issue_id") {
    const withoutDate = t.replace(/\b\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}\b/g, " ").trim();
    const m = withoutDate.match(/[A-Za-z0-9][A-Za-z0-9\-_.]*/);
    return m ? m[0] : (withoutDate || t);
  }

  if (field === "sheet_id") {
    // Apply 0/O disambiguation before pattern matching
    const disambiguated = disambiguate0AndO(t);
    const m = disambiguated.match(/[A-Za-z0-9][A-Za-z0-9\-_.]*/);
    const result = m ? m[0] : disambiguated;

    // Log if disambiguation made changes (for debugging)
    if (t !== disambiguated) {
      console.log(`🔍 sheet_id disambiguation: "${t}" → "${result}"`);
    }

    return result;
  }

  if (field === "project_id") {
    const m = t.match(/[A-Za-z0-9][A-Za-z0-9\-_.\/]+/);
    return m ? m[0] : t;
  }

  return t;
}

function _ocrPassConfigs(field) {
  if (field === "sheet_id" || field === "issue_id" || field === "project_id") {
    return [
      { psm: "7", whitelist: "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-_.\/()", label: "line" },
      { psm: "8", whitelist: "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-_.\/()", label: "word" },
      { psm: "6", whitelist: "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-_.\/()", label: "block" },
      { psm: "13", whitelist: "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-_.\/()", label: "raw-line" },
    ];
  }

  if (field === "date") {
    return [
      { psm: "7", whitelist: "0123456789/.-", label: "line" },
      { psm: "8", whitelist: "0123456789/.-", label: "word" },
      { psm: "6", whitelist: "0123456789/.-", label: "block" },
      { psm: "13", whitelist: "0123456789/.-", label: "raw-line" },
    ];
  }

  return [
    { psm: "6", whitelist: null, label: "block" },
    { psm: "7", whitelist: null, label: "line" },
    { psm: "4", whitelist: null, label: "single-column" },
    { psm: "11", whitelist: null, label: "sparse" },
  ];
}

async function _setOcrParams(worker, field, psm, whitelist) {
  const params = {
    user_defined_dpi: "300",
    preserve_interword_spaces: "1",
    tessedit_pageseg_mode: String(psm || "6"),
    tessedit_char_blacklist: "",
    classify_bln_numeric_mode: "0",
  };
  
  if (whitelist) {
    params.tessedit_char_whitelist = whitelist;
  } else {
    params.tessedit_char_whitelist = "";
  }
  
  await worker.setParameters(params);
}

function _scoreOcrCandidate(text, confidence) {
  const c = Number.isFinite(confidence) ? confidence : 0;
  const len = _normText(text).length;
  return (c * 1000) + Math.min(200, len);
}

async function ocrRecognizeMultiPass(worker, blob, field) {
  const passes = _ocrPassConfigs(field);
  let best = { text: "", confidence: 0, psm: null, label: null, score: -Infinity };

  for (const pass of passes) {
    await _setOcrParams(worker, field, pass.psm, pass.whitelist);
    const { data } = await worker.recognize(blob);
    const text = _normText(data?.text || "");
    const confidence = Number(data?.confidence) || 0;

    const score = _scoreOcrCandidate(text, confidence);
    if (score > best.score) {
      best = { text, confidence, psm: pass.psm, label: pass.label, score };
    }
  }

  return best;
}

async function extractVectorTextFromRegion(pageNum, region) {
  const page = await pdfDoc.getPage(pageNum);
  const viewport = page.getViewport({ scale });
  const textContent = await page.getTextContent();

  const xMin = region.x * viewport.width;
  const yMin = region.y * viewport.height;
  const xMax = xMin + region.w * viewport.width;
  const yMax = yMin + region.h * viewport.height;

  const strings = [];

  textContent.items.forEach((item) => {
    const [, , , , tx, ty] = pdfjsLib.Util.transform(viewport.transform, item.transform);
    if (tx >= xMin && tx <= xMax && ty >= yMin && ty <= yMax) {
      strings.push(item.str);
    }
  });

  return strings.join(" ").replace(/\s+/g, " ").trim();
}

async function extractOCRFromRegion(pageNum, region) {
  let page;
  try {
    page = await pdfDoc.getPage(pageNum);
  } catch (err) {
    console.error(`Error loading page ${pageNum} for OCR:`, err);
    return "";
  }

  const regionAreaPx = region.w * region.h * page.view[2] * page.view[3];
  let OCR_SCALE;
  
  if (regionAreaPx < 5000) {
    OCR_SCALE = 5.0;
  } else if (regionAreaPx < 20000) {
    OCR_SCALE = 4.0;
  } else if (regionAreaPx < 100000) {
    OCR_SCALE = 3.0;
  } else {
    OCR_SCALE = 2.5;
  }
  
  const viewport = page.getViewport({ scale: OCR_SCALE });
  const offCanvas = document.createElement("canvas");
  
  const cropW = Math.max(1, Math.round(region.w * viewport.width));
  const cropH = Math.max(1, Math.round(region.h * viewport.height));
  
  const MAX_OCR_DIMENSION = 2048;
  let finalCropW = cropW;
  let finalCropH = cropH;
  
  if (cropW > MAX_OCR_DIMENSION || cropH > MAX_OCR_DIMENSION) {
    const scaleDown = Math.min(MAX_OCR_DIMENSION / cropW, MAX_OCR_DIMENSION / cropH);
    finalCropW = Math.round(cropW * scaleDown);
    finalCropH = Math.round(cropH * scaleDown);
    console.warn(`⚠️ OCR region too large (${cropW}x${cropH}), scaling to ${finalCropW}x${finalCropH}`);
  }
  
  const MIN_OCR_WIDTH = 50;
  const MIN_OCR_HEIGHT = 20;
  
  if (finalCropW < MIN_OCR_WIDTH || finalCropH < MIN_OCR_HEIGHT) {
    console.warn(`⚠️ Region too small for OCR (${finalCropW}x${finalCropH}px)`);
    if (page && page.cleanup) page.cleanup();
    return "";
  }

  offCanvas.width = viewport.width;
  offCanvas.height = viewport.height;

  try {
    await page.render({
      canvasContext: offCanvas.getContext("2d", { alpha: false }),
      viewport,
      intent: 'print',
      renderInteractiveForms: false,
      enableWebGL: false,
    }).promise;
  } catch (err) {
    console.error('Error rendering page for OCR:', err);
    if (page && page.cleanup) page.cleanup();
    offCanvas.width = 0;
    offCanvas.height = 0;
    return "";
  }

  const crop = document.createElement("canvas");
  crop.width = finalCropW;
  crop.height = finalCropH;

  const ctx = crop.getContext("2d", { alpha: false });
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = "high";
  
  ctx.drawImage(
    offCanvas,
    region.x * offCanvas.width,
    region.y * offCanvas.height,
    cropW,
    cropH,
    0,
    0,
    finalCropW,
    finalCropH
  );

  offCanvas.width = 0;
  offCanvas.height = 0;

  const field = (region && region.type) ? region.type : "";
  preprocessOcrCrop(crop, field);

  const worker = await getOcrWorker();
  const blob = await new Promise((res) => crop.toBlob(res, "image/jpeg", 0.92));
  
  crop.width = 0;
  crop.height = 0;
  
  if (page && page.cleanup) page.cleanup();
  page = null;
  
  if (!blob) return "";

  return await runOcrExclusive(async () => {
    const best = await ocrRecognizeMultiPass(worker, blob, field);
    return cleanByField(field, best.text);
  });
}

async function applyTemplatesToAllPages(logProgress = false) {
  if (!pdfDoc) return;

  const templateCount = Object.keys(regionTemplates).length;
  console.log(`📐 Active templates: ${templateCount}`, Object.keys(regionTemplates));
  
  if (templateCount === 0) {
    console.warn("⚠️ No templates defined! Draw regions on a page first, then click Extract.");
    return;
  }

  const BATCH_SIZE = 3;
  
  for (let batchStart = 1; batchStart <= pdfDoc.numPages; batchStart += BATCH_SIZE) {
    const batchEnd = Math.min(batchStart + BATCH_SIZE - 1, pdfDoc.numPages);
    
    if (logProgress) {
      console.log(`🔍 Extracting pages ${batchStart}-${batchEnd} / ${pdfDoc.numPages}`);
    }
    
    for (let pageNum = batchStart; pageNum <= batchEnd; pageNum++) {
      if (!sheetDetailsByPage[pageNum]) sheetDetailsByPage[pageNum] = {};

      for (const field of REGION_TYPES) {
        if (Object.prototype.hasOwnProperty.call(sheetDetailsByPage[pageNum], field)) continue;

        const regions = resolveAllRegionsForPage(pageNum, field);
        if (regions.length === 0) {
          console.warn(`⚠️ Page ${pageNum}: No region for "${field}"`);
          continue;
        }

        let extractedParts = [];

        for (const region of regions) {
          let extracted = "";

          if (ENABLE_VECTOR_EXTRACTION) {
            extracted = await extractVectorTextFromRegion(pageNum, region);
          }

          if (!extracted) {
            extracted = await extractOCRFromRegion(pageNum, region);
          }

          extracted = cleanByField(field, extracted);
          if (extracted) {
            extractedParts.push(extracted);
          }
        }

        const finalExtracted = extractedParts.join(" ");
        sheetDetailsByPage[pageNum][field] = finalExtracted || "";
        
        if (logProgress && finalExtracted) {
          if (extractedParts.length > 1) {
            console.log(`  ✓ Page ${pageNum} "${field}": "${finalExtracted}" (from ${extractedParts.length} regions)`);
          } else {
            console.log(`  ✓ Page ${pageNum} "${field}": "${finalExtracted}"`);
          }
        }
      }
    }
    
    await new Promise(resolve => setTimeout(resolve, 100));
    
    if (window.gc) {
      window.gc();
    }
  }

  console.log("✅ Templates applied to all pages");
}

window.applyTemplatesToAllPages = applyTemplatesToAllPages;

async function extractAll() {
  if (!pdfDoc) return alert("No PDF loaded");

  console.log(`🚀 Extract started - extracting DOCUMENT_DETAILS only (prepared_by, project_id)`);

  for (const field of DOCUMENT_DETAILS) {
    const region = getMostRecentRegionOfType(currentPage, field);
    if (!region) {
      console.warn(`⚠️ No region drawn for document field: ${field}`);
      continue;
    }

    let extracted = await extractOCRFromRegion(currentPage, region);
    extracted = (extracted || "").trim();

    if (extracted) {
      documentDetails[field] = extracted;
      if (field === "prepared_by" && preparedByInput) preparedByInput.value = extracted;
      if (field === "project_id" && projectIdInput) projectIdInput.value = extracted;
    } else {
      console.warn(`⚠️ Document field (${field}) read empty; keeping existing value`, documentDetails[field] || "<empty>");
    }

    console.log(`📄 Document field (${field}) →`, (documentDetails[field] || "").trim() || "<empty>");
  }

  console.log("✅ Extract complete (DOCUMENT_DETAILS only)");
  console.log("ℹ️ REGION_TYPES (sheet_id, description, etc.) are NOT extracted - they use templates/overrides automatically");
}

window.extractAll = extractAll;

window.addEventListener("keydown", (e) => {
  if (isEditableTarget(document.activeElement)) return;

  const modKey = e.ctrlKey || e.metaKey;

  // UNDO: Ctrl/Cmd + Z (without Shift)
  if (modKey && !e.shiftKey && (e.key === "z" || e.key === "Z")) {
    undo();
    e.preventDefault();
    return;
  }

  // REDO: Ctrl/Cmd + Shift + Z
  if (modKey && e.shiftKey && (e.key === "z" || e.key === "Z")) {
    redo();
    e.preventDefault();
    return;
  }

  // COPY: Ctrl/Cmd + C
  if (modKey && (e.key === "c" || e.key === "C")) {
    if (selectedRegionIds.length) {
      copySelectionToClipboard();
      e.preventDefault();
    }
    return;
  }

  // CUT: Ctrl/Cmd + X
  if (modKey && (e.key === "x" || e.key === "X")) {
    if (selectedRegionIds.length) {
      const didCopy = copySelectionToClipboard();
      if (didCopy) {
        const regions = regionsByPage[currentPage] || [];
        const sel = new Set(selectedRegionIds);
        const cutTypes = [...new Set(regions.filter(r => sel.has(r.id)).map(r => r.type))];
        regionsByPage[currentPage] = regions.filter((r) => !sel.has(r.id));
        invalidatePageFields(currentPage, cutTypes);
        
        saveUndoState(); // ADDED
        
        clearSelection();
        redrawRegions();
        console.log(`✂️ Cut ${clipboardRegions.length} region(s) from page ${currentPage}`);
      }
      e.preventDefault();
    }
    return;
  }

  // PASTE: Ctrl/Cmd + V (Shift+V for paste at pointer)
  if (modKey && (e.key === "v" || e.key === "V")) {
    if (clipboardRegions.length) {
      if (e.shiftKey) {
        pasteClipboardToCurrentPageAtPointer();
      } else {
        pasteClipboardToCurrentPage();
      }
      e.preventDefault();
    }
    return;
  }

  // DELETE: Delete or Backspace key
  if (e.key === "Delete" || e.key === "Backspace") {
  // Measurement delete takes priority
  if (msrSelectedId !== null) {
    msrDeleteSelected();
    e.preventDefault();
    return;
  }
  if (!selectedRegionIds.length) return;

  saveUndoState();

  const regions = regionsByPage[currentPage] || [];
  const sel = new Set(selectedRegionIds);
  const selectedRegions = regions.filter(r => sel.has(r.id));
  
  selectedRegions.forEach(region => {
    if (region.isGhost) {
      // RULE 1: Delete ghost on this page only - add to exclusion list
      console.log(`🗑️ Hiding ghost "${region.type}" on page ${currentPage}`);
      
      if (!ghostExclusions[currentPage]) {
        ghostExclusions[currentPage] = new Set();
      }
      ghostExclusions[currentPage].add(region.type);
      
      regionsByPage[currentPage] = regionsByPage[currentPage].filter(r => r.id !== region.id);
      invalidatePageFields(currentPage, [region.type]);
      
    } else if (currentPage === TEMPLATE_MASTER_PAGE) {
      // RULE 2: Delete from master page - removes template and all instances
      console.log(`🗑️ Deleting "${region.type}" from MASTER PAGE - removing from ALL pages`);
      
      if (regionTemplates[region.type]) {
        delete regionTemplates[region.type];
      }
      
      Object.keys(regionsByPage).forEach(pageNum => {
        regionsByPage[pageNum] = regionsByPage[pageNum].filter(r => r.type !== region.type);
        invalidatePageFields(parseInt(pageNum), [region.type]);
      });
      
    } else {
      // RULE 3: Delete page-specific override - only removes from this page
      console.log(`🗑️ Deleting page-specific override "${region.type}" from page ${currentPage} only`);
      
      regionsByPage[currentPage] = regionsByPage[currentPage].filter(r => r.id !== region.id);
      invalidatePageFields(currentPage, [region.type]);
    }
  });
  
  clearSelection();
  redrawRegions();
  e.preventDefault();
}
});

pdfScroll?.addEventListener("wheel", (e) => {
  e.preventDefault();
  if (e.shiftKey) return;

  const PAN_SPEED = 3;
  const ZOOM_FACTOR = 1.1;

  if (e.ctrlKey) {
    pdfScroll.scrollLeft += e.deltaY * PAN_SPEED;
    return;
  }
  if (e.altKey) {
    pdfScroll.scrollTop += e.deltaY * PAN_SPEED;
    return;
  }

  if (!pdfDoc) return;

  const rect = pdfScroll.getBoundingClientRect();
  const mx = e.clientX - rect.left;
  const my = e.clientY - rect.top;

  // Lock the page-space anchor once at gesture start.
  // canvasPadX/Y is the fixed offset of the canvas inside canvas-outer,
  // so subtracting it gives coordinates relative to the canvas itself.
  if (!gestureAnchor) {
    gestureAnchor = {
      pageX: (pdfScroll.scrollLeft + mx - canvasPadX) / scale,
      pageY: (pdfScroll.scrollTop  + my - canvasPadY) / scale,
      mx, my,
    };
  }

  scale *= e.deltaY < 0 ? ZOOM_FACTOR : 1 / ZOOM_FACTOR;
  scale = Math.min(Math.max(scale, 0.3), 20.0);

  // Cap scale to what the browser can actually render. Without this the CSS
  // scale races far ahead of the renderable limit; when the debounce fires the
  // canvas is much smaller than expected and scrollLeft gets clamped hard.
  if (canvas.width > 0 && renderedScale > 0) {
    const baseW = canvas.width  / renderedScale;
    const baseH = canvas.height / renderedScale;
    scale = Math.min(scale, MAX_CANVAS_DIMENSION / baseW, MAX_CANVAS_DIMENSION / baseH);
  }

  const cssRatio = scale / renderedScale;
  canvas.style.width  = Math.round(canvas.width  * cssRatio) + 'px';
  canvas.style.height = Math.round(canvas.height * cssRatio) + 'px';

  // Apply scroll so the locked anchor stays under the pointer.
  // No forced reflow here — gestureAnchor keeps the formula stable even if
  // one event's scrollLeft is briefly clamped before the layout catches up.
  pdfScroll.scrollLeft = gestureAnchor.pageX * scale + canvasPadX - gestureAnchor.mx;
  pdfScroll.scrollTop  = gestureAnchor.pageY * scale + canvasPadY - gestureAnchor.my;

  updateControls();

  // Re-render at full quality once the gesture settles.
  // Save the anchor so we can re-apply scroll after the render corrects scale.
  clearTimeout(zoomDebounceTimer);
  zoomDebounceTimer = setTimeout(async () => {
    const anchor = gestureAnchor;
    gestureAnchor = null;
    await renderPage(currentPage);
    // renderPage may cap scale further (MAX_CANVAS_DIMENSION); re-apply scroll
    // now that the canvas is at its true rendered size so nothing is clamped.
    if (anchor) {
      pdfScroll.scrollLeft = anchor.pageX * scale + canvasPadX - anchor.mx;
      pdfScroll.scrollTop  = anchor.pageY * scale + canvasPadY - anchor.my;
    }
  }, 150);
}, { passive: false });

function getCanonicalExportData() {
  const doc = {
    prepared_by: (documentDetails.prepared_by || "").trim(),
    project_id: (documentDetails.project_id || "").trim(),
  };

  const sheets = [];
  const numPages = pdfDoc?.numPages || 0;

  console.log(`📊 Generating export data for ${numPages} pages...`);

  for (let p = 1; p <= numPages; p++) {
    const s = sheetDetailsByPage[p] || {};
    const sheet = {
      page: p,
      sheet_id: (s.sheet_id || "").trim(),
      description: (s.description || "").trim(),
      issue_id: (s.issue_id || "").trim(),
      date: (s.date || "").trim(),
      issue_description: (s.issue_description || "").trim(),
    };
    sheets.push(sheet);
    
    const hasData = Object.values(sheet).slice(1).some(v => v !== "");
    if (!hasData && p > 1) {
      console.warn(`⚠️ Page ${p} has no extracted data (templates may not have been applied)`);
    }
  }

  console.log(`📊 Generated ${sheets.length} sheet records`);
  return { document: doc, sheets };
}

window.exportExtractedData = async function () {
  if (typeof applyTemplatesToAllPages === "function") {
    await applyTemplatesToAllPages(true);
  }
  const data = getCanonicalExportData();
  console.log(JSON.stringify(data, null, 2));
  return data;
};

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Processing indicator functions
function showProcessing(message) {
  const existing = document.getElementById('processing-overlay');
  if (existing) existing.remove();
  
  const overlay = document.createElement('div');
  overlay.id = 'processing-overlay';
  overlay.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background: rgba(0, 0, 0, 0.7);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 10000;
    font-family: sans-serif;
  `;
  
  const messageBox = document.createElement('div');
  messageBox.style.cssText = `
    background: white;
    padding: 30px 50px;
    border-radius: 8px;
    font-size: 18px;
    font-weight: 500;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
  `;
  messageBox.textContent = message;
  
  overlay.appendChild(messageBox);
  document.body.appendChild(overlay);
  
  // Disable export buttons
  const buttons = document.querySelectorAll('button[onclick*="download"]');
  buttons.forEach(btn => btn.disabled = true);
}

function hideProcessing() {
  const overlay = document.getElementById('processing-overlay');
  if (overlay) overlay.remove();
  
  // Re-enable export buttons
  const buttons = document.querySelectorAll('button[onclick*="download"]');
  buttons.forEach(btn => btn.disabled = false);
}

window.downloadJSON = async function () {
  try {
    showProcessing('Processing JSON export...');
    console.log("📦 Starting JSON export...");
    
    if (typeof applyTemplatesToAllPages === "function") {
      await applyTemplatesToAllPages(true);
    }
    
    const data = getCanonicalExportData();
    
    console.log(`✅ Exporting ${data.sheets.length} pages`);
    
    const json = JSON.stringify(data, null, 2);
    downloadBlob(new Blob([json], { type: "application/json" }), `${pdfFileBaseName}.json`);
    console.log("⬇️ JSON exported:", `${pdfFileBaseName}.json`);
  } catch (error) {
    console.error("❌ JSON export error:", error);
    alert("Error exporting JSON: " + error.message);
  } finally {
    hideProcessing();
  }
};

window.downloadCSV = async function () {
  try {
    showProcessing('Processing CSV export...');
    console.log("📦 Starting CSV export...");
    
    if (typeof applyTemplatesToAllPages === "function") {
      await applyTemplatesToAllPages(true);
    }

    const { document, sheets } = getCanonicalExportData();
    
    console.log(`✅ Exporting ${sheets.length} pages`);

    const headers = [
      "page",
      "prepared_by",
      "project_id",
      "sheet_id",
      "description",
      "issue_id",
      "date",
      "issue_description",
    ];

    const rows = sheets.map((s) => [
      s.page,
      document.prepared_by,
      document.project_id,
      s.sheet_id,
      s.description,
      s.issue_id,
      s.date,
      s.issue_description,
    ]);

    const csv = [
      headers.join(","),
      ...rows.map((r) =>
        r.map((v) => `"${String(v ?? "").replace(/"/g, '""')}"`).join(",")
      ),
    ].join("\n");

    downloadBlob(new Blob([csv], { type: "text/csv" }), `${pdfFileBaseName}.csv`);
    console.log("⬇️ CSV exported:", `${pdfFileBaseName}.csv`);
  } catch (error) {
    console.error("❌ CSV export error:", error);
    alert("Error exporting CSV: " + error.message);
  } finally {
    hideProcessing();
  }
};

// ── Expose globals for pdf_splitter.js ──────────────────────────────────────
// top-level `let`/`const` are NOT window properties in modern browsers,
// so we use getters so splitter always reads the current value.
Object.defineProperty(window, 'pdfDoc',             { get: () => pdfDoc,             configurable: true });
Object.defineProperty(window, 'pdfRawBytes',        { get: () => pdfRawBytes,        configurable: true });
Object.defineProperty(window, 'isMultiPdfMode',     { get: () => isMultiPdfMode,     configurable: true });
Object.defineProperty(window, 'multiPdfDocs',       { get: () => multiPdfDocs,       configurable: true });
Object.defineProperty(window, 'documentDetails',    { get: () => documentDetails,    configurable: true });
Object.defineProperty(window, 'sheetDetailsByPage', { get: () => sheetDetailsByPage, configurable: true });
Object.defineProperty(window, 'regionTemplates',    { get: () => regionTemplates,    configurable: true });

// ── Measurement system ───────────────────────────────────────────────────────

const MSR_SVG_NS = "http://www.w3.org/2000/svg";

// Data stores
const measurementsByPage = {};  // { pageNum: [measurement, ...] }
const scaleZonesByPage   = {};  // { pageNum: [zone, ...] }
const pageBaseDimsCache  = new Map(); // pageNum → { width, height } in PDF pts
let   msrIdCounter       = 1;

// Drawing state
let msrActiveDrawPts = [];   // normalized points being placed
let msrPreviewPt     = null; // current mouse position (normalized)
let msrSelectedId    = null; // selected measurement id (number) or zone id (string "z<n>")

// Scale zone dialog state
let szPendingPts  = null;   // polygon verts waiting for dialog
let szRefState    = null;   // { pts: [], phase: 'drawing'|'done' }
let szActiveTab   = 'ratio';

// Element refs — measure toolbar
const msrToolEls = {
  'select':     document.getElementById('tool-select'),
  'scale-zone': document.getElementById('tool-scale-zone'),
  'linear':     document.getElementById('tool-linear'),
  'area':       document.getElementById('tool-area'),
  'count':      document.getElementById('tool-count'),
};
const activeScaleLblEl = document.getElementById('active-scale-lbl');
const clearMeasBtn     = document.getElementById('clear-measurements');

// Element refs — scale zone dialog
const szDialogEl    = document.getElementById('sz-dialog');
const szBackdropEl  = document.getElementById('sz-backdrop');
const szTabRatioEl  = document.getElementById('sz-tab-ratio');
const szTabRefEl    = document.getElementById('sz-tab-ref');
const szPanelRatio  = document.getElementById('sz-panel-ratio');
const szPanelRef    = document.getElementById('sz-panel-ref');
const szRatioSel    = document.getElementById('sz-ratio-sel');
const szRatioCust   = document.getElementById('sz-ratio-custom');
const szRefStatusEl = document.getElementById('sz-ref-status');
const szRefLenRow   = document.getElementById('sz-ref-len-row');
const szRefLenEl    = document.getElementById('sz-ref-len');
const szOkBtn       = document.getElementById('sz-ok');
const szCancelBtn   = document.getElementById('sz-cancel');

// ── Tool management ──────────────────────────────────────────────────────────

function msrSetTool(t) {
  currentTool      = t;
  msrActiveDrawPts = [];
  msrPreviewPt     = null;
  msrSelectedId    = null;
  szRefState       = null;
  Object.entries(msrToolEls).forEach(([k, btn]) => {
    if (btn) btn.classList.toggle('is-active', k === t);
  });
  if (overlay) overlay.style.cursor = (t === 'select') ? 'crosshair' : 'crosshair';
  msrRedrawOnly();
}

Object.entries(msrToolEls).forEach(([t, btn]) => btn?.addEventListener('click', () => msrSetTool(t)));

clearMeasBtn?.addEventListener('click', () => {
  if (!confirm('Clear all measurements and scale zones on all pages?')) return;
  Object.keys(measurementsByPage).forEach(k => delete measurementsByPage[k]);
  Object.keys(scaleZonesByPage).forEach(k => delete scaleZonesByPage[k]);
  msrActiveDrawPts = [];
  msrPreviewPt     = null;
  msrSelectedId    = null;
  redrawRegions();
});

// ── Geometry helpers ─────────────────────────────────────────────────────────

function msrNormToMeters(x1, y1, x2, y2, pw, ph, mpp) {
  const dx = (x2 - x1) * pw / 72 * 0.0254;
  const dy = (y2 - y1) * ph / 72 * 0.0254;
  return Math.sqrt(dx * dx + dy * dy) * mpp;
}

function msrNormToSqMeters(pts, pw, ph, mpp) {
  let area = 0;
  const n = pts.length;
  for (let i = 0; i < n; i++) {
    const j = (i + 1) % n;
    area += pts[i].x * pw * pts[j].y * ph;
    area -= pts[j].x * pw * pts[i].y * ph;
  }
  const paperM2 = Math.abs(area) / 2 / (72 * 72) * (0.0254 * 0.0254);
  return paperM2 * mpp * mpp;
}

function msrPointInPoly(x, y, verts) {
  let inside = false;
  const n = verts.length;
  for (let i = 0, j = n - 1; i < n; j = i++) {
    const xi = verts[i].x, yi = verts[i].y, xj = verts[j].x, yj = verts[j].y;
    if (((yi > y) !== (yj > y)) && x < (xj - xi) * (y - yi) / (yj - yi) + xi) inside = !inside;
  }
  return inside;
}

function msrFindZone(nx, ny, pageNum) {
  const zones = scaleZonesByPage[pageNum] || [];
  for (let i = zones.length - 1; i >= 0; i--) {
    if (msrPointInPoly(nx, ny, zones[i].vertices)) return zones[i];
  }
  return null;
}

// Pixel distance from a normalized point to the first active draw point
function msrSnapPxDist(pt) {
  if (!msrActiveDrawPts.length) return Infinity;
  const f = msrActiveDrawPts[0];
  return Math.hypot((pt.x - f.x) * canvas.width, (pt.y - f.y) * canvas.height);
}

// ── Scale zone dialog ────────────────────────────────────────────────────────

function szShowDialog(vertices) {
  szPendingPts  = vertices;
  szRefState    = null;
  szSwitchTab('ratio');
  szRatioSel.value = '100';
  szRatioCust.style.display = 'none';
  szRatioCust.value = '';
  szRefStatusEl.textContent = 'Click first point on the PDF…';
  szRefLenRow.style.display = 'none';
  szRefLenEl.value = '';
  szOkBtn.disabled = false;
  szDialogEl.hidden = false;
  szBackdropEl.hidden = false;
  szBackdropEl.style.pointerEvents = '';
}

function szHideDialog() {
  szDialogEl.hidden  = true;
  szBackdropEl.hidden = true;
  szPendingPts = null;
  szRefState   = null;
}

function szSwitchTab(tab) {
  szActiveTab = tab;
  szTabRatioEl?.classList.toggle('is-active', tab === 'ratio');
  szTabRefEl?.classList.toggle('is-active',   tab === 'ref');
  szPanelRatio?.classList.toggle('is-active', tab === 'ratio');
  szPanelRef?.classList.toggle('is-active',   tab === 'ref');
  if (tab === 'ref') {
    szRefState = { pts: [], phase: 'drawing' };
    szOkBtn.disabled = true;
    szRefStatusEl.textContent = 'Click first point on the PDF…';
    szRefLenRow.style.display = 'none';
    szBackdropEl.style.pointerEvents = 'none'; // let clicks through to canvas
  } else {
    szRefState = null;
    szOkBtn.disabled = false;
    szBackdropEl.style.pointerEvents = '';
  }
  msrRedrawOnly();
}

szTabRatioEl?.addEventListener('click', () => szSwitchTab('ratio'));
szTabRefEl?.addEventListener('click',   () => szSwitchTab('ref'));

szRatioSel?.addEventListener('change', () => {
  const custom = szRatioSel.value === 'custom';
  szRatioCust.style.display = custom ? '' : 'none';
});

szCancelBtn?.addEventListener('click', () => {
  szHideDialog();
  msrActiveDrawPts = [];
  msrRedrawOnly();
});

szOkBtn?.addEventListener('click', () => {
  if (!szPendingPts) return;

  let mpp, label;

  if (szActiveTab === 'ratio') {
    let ratio = szRatioSel.value === 'custom'
      ? parseFloat(szRatioCust.value)
      : parseFloat(szRatioSel.value);
    if (!Number.isFinite(ratio) || ratio <= 0) { alert('Enter a valid scale ratio.'); return; }
    mpp   = ratio;
    label = `1:${ratio}`;
  } else {
    if (!szRefState || szRefState.pts.length < 2) { alert('Draw a reference line on the PDF first.'); return; }
    const realLen = parseFloat(szRefLenEl.value);
    if (!Number.isFinite(realLen) || realLen <= 0) { alert('Enter a valid real-world length.'); return; }
    const dims = pageBaseDimsCache.get(currentPage);
    if (!dims) { alert('Page dimensions not cached — render the page first.'); return; }
    const [p1, p2] = szRefState.pts;
    const paperLen = msrNormToMeters(p1.x, p1.y, p2.x, p2.y, dims.width, dims.height, 1.0);
    mpp   = realLen / paperLen;
    label = `~1:${Math.round(mpp)}`;
  }

  if (!scaleZonesByPage[currentPage]) scaleZonesByPage[currentPage] = [];
  scaleZonesByPage[currentPage].push({ id: msrIdCounter++, vertices: szPendingPts, mpp, label });

  szHideDialog();
  msrActiveDrawPts = [];
  redrawRegions();
});

// ── Overlay event handlers for measurement tools ─────────────────────────────

function msrOverlayNorm(e) {
  const r = overlay.getBoundingClientRect();
  return { x: (e.clientX - r.left) / canvas.width, y: (e.clientY - r.top) / canvas.height };
}

overlay?.addEventListener('mousemove', (e) => {
  if (currentTool === 'select') return;
  msrPreviewPt = msrOverlayNorm(e);
  msrRedrawOnly();
});

overlay?.addEventListener('click', (e) => {
  if (currentTool === 'select') return;
  if (e.detail > 1) return; // handled by dblclick

  const pt = msrOverlayNorm(e);

  // Reference line point collection (dialog open, ref tab active)
  if (szDialogEl && !szDialogEl.hidden && szActiveTab === 'ref' && szRefState?.phase === 'drawing') {
    szRefState.pts.push(pt);
    if (szRefState.pts.length >= 2) {
      szRefState.phase = 'done';
      szRefStatusEl.textContent = '✓ Reference line set. Enter the real-world length below.';
      szRefLenRow.style.display = '';
      szOkBtn.disabled = false;
      szRefLenEl.focus();
      szBackdropEl.style.pointerEvents = ''; // restore modal behaviour
    } else {
      szRefStatusEl.textContent = 'Click second point on the PDF…';
    }
    msrRedrawOnly();
    return;
  }

  if (szDialogEl && !szDialogEl.hidden) return; // dialog open in non-ref mode, ignore

  switch (currentTool) {

    case 'linear':
      msrActiveDrawPts.push(pt);
      if (msrActiveDrawPts.length >= 2) {
        msrFinishLinear([...msrActiveDrawPts]);
        msrActiveDrawPts = [];
        msrPreviewPt = null;
      }
      break;

    case 'area':
    case 'scale-zone':
      if (msrActiveDrawPts.length >= 3 && msrSnapPxDist(pt) < 15) {
        if (currentTool === 'area') {
          msrFinishArea([...msrActiveDrawPts]);
        } else {
          szShowDialog([...msrActiveDrawPts]);
        }
        msrActiveDrawPts = [];
        msrPreviewPt = null;
      } else {
        msrActiveDrawPts.push(pt);
      }
      break;

    case 'count':
      msrFinishCount(pt);
      break;
  }

  msrRedrawOnly();
});

// Right-click cancels in-progress drawing
overlay?.addEventListener('contextmenu', (e) => {
  if (currentTool === 'select') return;
  e.preventDefault();
  msrActiveDrawPts = [];
  msrPreviewPt     = null;
  msrRedrawOnly();
});

// ── Finish measurements ──────────────────────────────────────────────────────

function msrFinishLinear(pts) {
  if (!measurementsByPage[currentPage]) measurementsByPage[currentPage] = [];
  measurementsByPage[currentPage].push({ id: msrIdCounter++, type: 'linear', points: pts, label: '' });
  redrawRegions();
}

function msrFinishArea(pts) {
  if (pts.length < 3) return;
  if (!measurementsByPage[currentPage]) measurementsByPage[currentPage] = [];
  measurementsByPage[currentPage].push({ id: msrIdCounter++, type: 'area', points: pts, label: '' });
  redrawRegions();
}

function msrFinishCount(pt) {
  const msrs = measurementsByPage[currentPage] || (measurementsByPage[currentPage] = []);
  // Add to the last open count group on this page, or start a new one
  const existing = [...msrs].reverse().find(m => m.type === 'count');
  if (existing) {
    existing.points.push(pt);
  } else {
    msrs.push({ id: msrIdCounter++, type: 'count', points: [pt], label: '' });
  }
  redrawRegions();
}

function msrDeleteSelected() {
  if (msrSelectedId === null) return;
  const id = msrSelectedId;
  msrSelectedId = null;

  // Zone delete
  if (typeof id === 'string' && id.startsWith('z')) {
    const zid = parseInt(id.slice(1));
    if (scaleZonesByPage[currentPage]) {
      scaleZonesByPage[currentPage] = scaleZonesByPage[currentPage].filter(z => z.id !== zid);
    }
  } else {
    // Measurement delete
    if (measurementsByPage[currentPage]) {
      measurementsByPage[currentPage] = measurementsByPage[currentPage].filter(m => m.id !== id);
    }
  }
  redrawRegions();
}

// ── Value formatters ─────────────────────────────────────────────────────────

function msrFmtLinear(m, dims) {
  if (!dims) return '? m';
  const zone = msrFindZone(m.points[0].x, m.points[0].y, currentPage);
  if (!zone) return '? m';
  const d = msrNormToMeters(m.points[0].x, m.points[0].y, m.points[1].x, m.points[1].y, dims.width, dims.height, zone.mpp);
  return `${d.toFixed(3)} m${m.label ? '  ' + m.label : ''}`;
}

function msrFmtArea(m, dims) {
  if (!dims) return '? m²';
  const cx = m.points.reduce((s, p) => s + p.x, 0) / m.points.length;
  const cy = m.points.reduce((s, p) => s + p.y, 0) / m.points.length;
  const zone = msrFindZone(cx, cy, currentPage);
  if (!zone) return '? m²';
  const a = msrNormToSqMeters(m.points, dims.width, dims.height, zone.mpp);
  return `${a.toFixed(3)} m²${m.label ? '  ' + m.label : ''}`;
}

// ── SVG renderers ────────────────────────────────────────────────────────────

function msrRenderZones(g) {
  const zones = scaleZonesByPage[currentPage] || [];
  zones.forEach(zone => {
    if (zone.vertices.length < 3) return;
    const isSel = msrSelectedId === `z${zone.id}`;
    const pts = zone.vertices.map(v => `${v.x * canvas.width},${v.y * canvas.height}`).join(' ');

    const poly = document.createElementNS(MSR_SVG_NS, 'polygon');
    poly.setAttribute('points', pts);
    poly.classList.add('msr-zone-poly');
    if (isSel) poly.classList.add('msr-sel');
    poly.addEventListener('click', (e) => {
      if (currentTool !== 'select') return;
      e.stopPropagation();
      msrSelectedId = `z${zone.id}`;
      redrawRegions();
    });
    poly.addEventListener('dblclick', (e) => {
      if (currentTool !== 'select') return;
      e.stopPropagation();
      const n = prompt('Edit zone label:', zone.label);
      if (n !== null) { zone.label = n; redrawRegions(); }
    });
    g.appendChild(poly);

    const cx = zone.vertices.reduce((s, v) => s + v.x, 0) / zone.vertices.length * canvas.width;
    const cy = zone.vertices.reduce((s, v) => s + v.y, 0) / zone.vertices.length * canvas.height;
    const lbl = document.createElementNS(MSR_SVG_NS, 'text');
    lbl.setAttribute('x', cx); lbl.setAttribute('y', cy);
    lbl.setAttribute('text-anchor', 'middle'); lbl.setAttribute('dominant-baseline', 'central');
    lbl.classList.add('msr-zone-lbl');
    lbl.textContent = zone.label;
    g.appendChild(lbl);
  });
}

function msrRenderMeasurements(g) {
  const dims = pageBaseDimsCache.get(currentPage) || null;
  const msrs = measurementsByPage[currentPage] || [];

  msrs.forEach(m => {
    const isSel = msrSelectedId === m.id;
    if (m.type === 'linear') msrRenderLinear(g, m, dims, isSel);
    else if (m.type === 'area') msrRenderArea(g, m, dims, isSel);
    else if (m.type === 'count') msrRenderCount(g, m, isSel);
  });

  msrRenderPreview(g, dims);
}

function msrRenderLinear(g, m, dims, isSel) {
  const [p1, p2] = m.points;
  const x1 = p1.x * canvas.width,  y1 = p1.y * canvas.height;
  const x2 = p2.x * canvas.width,  y2 = p2.y * canvas.height;

  const grp = document.createElementNS(MSR_SVG_NS, 'g');
  grp.classList.add('msr-linear-g');
  if (isSel) grp.classList.add('msr-sel');

  const line = document.createElementNS(MSR_SVG_NS, 'line');
  line.setAttribute('x1', x1); line.setAttribute('y1', y1);
  line.setAttribute('x2', x2); line.setAttribute('y2', y2);
  line.classList.add('msr-line');
  grp.appendChild(line);

  // Perpendicular tick marks at each end
  const perp = Math.atan2(y2 - y1, x2 - x1) + Math.PI / 2;
  const TICK = 7;
  [[x1, y1], [x2, y2]].forEach(([tx, ty]) => {
    const tick = document.createElementNS(MSR_SVG_NS, 'line');
    tick.setAttribute('x1', tx + Math.cos(perp) * TICK); tick.setAttribute('y1', ty + Math.sin(perp) * TICK);
    tick.setAttribute('x2', tx - Math.cos(perp) * TICK); tick.setAttribute('y2', ty - Math.sin(perp) * TICK);
    tick.classList.add('msr-tick');
    grp.appendChild(tick);
  });

  const lbl = document.createElementNS(MSR_SVG_NS, 'text');
  lbl.setAttribute('x', (x1 + x2) / 2); lbl.setAttribute('y', (y1 + y2) / 2 - 9);
  lbl.setAttribute('text-anchor', 'middle');
  lbl.classList.add('msr-lbl'); if (isSel) lbl.classList.add('msr-sel');
  lbl.textContent = msrFmtLinear(m, dims);
  grp.appendChild(lbl);

  // Wide invisible hit area for easier clicking
  const hit = document.createElementNS(MSR_SVG_NS, 'line');
  hit.setAttribute('x1', x1); hit.setAttribute('y1', y1);
  hit.setAttribute('x2', x2); hit.setAttribute('y2', y2);
  hit.setAttribute('stroke', 'transparent'); hit.setAttribute('stroke-width', '14');
  grp.appendChild(hit);

  grp.style.cursor = 'pointer';
  grp.addEventListener('click', (e) => {
    if (currentTool !== 'select') return;
    e.stopPropagation(); msrSelectedId = m.id; redrawRegions();
  });
  grp.addEventListener('dblclick', (e) => {
    if (currentTool !== 'select') return;
    e.stopPropagation();
    const n = prompt('Edit label:', m.label); if (n !== null) { m.label = n; redrawRegions(); }
  });
  g.appendChild(grp);
}

function msrRenderArea(g, m, dims, isSel) {
  if (m.points.length < 3) return;
  const pts = m.points.map(p => `${p.x * canvas.width},${p.y * canvas.height}`).join(' ');
  const poly = document.createElementNS(MSR_SVG_NS, 'polygon');
  poly.setAttribute('points', pts);
  poly.classList.add('msr-area-poly'); if (isSel) poly.classList.add('msr-sel');
  poly.addEventListener('click', (e) => {
    if (currentTool !== 'select') return;
    e.stopPropagation(); msrSelectedId = m.id; redrawRegions();
  });
  poly.addEventListener('dblclick', (e) => {
    if (currentTool !== 'select') return;
    e.stopPropagation();
    const n = prompt('Edit label:', m.label); if (n !== null) { m.label = n; redrawRegions(); }
  });
  g.appendChild(poly);

  const cx = m.points.reduce((s, p) => s + p.x, 0) / m.points.length * canvas.width;
  const cy = m.points.reduce((s, p) => s + p.y, 0) / m.points.length * canvas.height;
  const lbl = document.createElementNS(MSR_SVG_NS, 'text');
  lbl.setAttribute('x', cx); lbl.setAttribute('y', cy);
  lbl.setAttribute('text-anchor', 'middle'); lbl.setAttribute('dominant-baseline', 'central');
  lbl.classList.add('msr-lbl'); if (isSel) lbl.classList.add('msr-sel');
  lbl.textContent = msrFmtArea(m, dims);
  g.appendChild(lbl);
}

function msrRenderCount(g, m, isSel) {
  const R = 9;
  m.points.forEach((p, i) => {
    const cx = p.x * canvas.width, cy = p.y * canvas.height;
    const c = document.createElementNS(MSR_SVG_NS, 'circle');
    c.setAttribute('cx', cx); c.setAttribute('cy', cy); c.setAttribute('r', R);
    c.classList.add('msr-count-dot'); if (isSel) c.classList.add('msr-sel');
    c.addEventListener('click', (e) => {
      if (currentTool !== 'select') return;
      e.stopPropagation(); msrSelectedId = m.id; redrawRegions();
    });
    g.appendChild(c);
    const num = document.createElementNS(MSR_SVG_NS, 'text');
    num.setAttribute('x', cx); num.setAttribute('y', cy);
    num.classList.add('msr-count-num');
    num.textContent = i + 1;
    g.appendChild(num);
  });

  if (m.points.length > 0) {
    const last = m.points[m.points.length - 1];
    const lbl = document.createElementNS(MSR_SVG_NS, 'text');
    lbl.setAttribute('x', last.x * canvas.width + R + 5);
    lbl.setAttribute('y', last.y * canvas.height);
    lbl.setAttribute('dominant-baseline', 'central');
    lbl.classList.add('msr-lbl'); if (isSel) lbl.classList.add('msr-sel');
    lbl.textContent = `n = ${m.points.length}${m.label ? '  ' + m.label : ''}`;
    g.appendChild(lbl);
  }
}

function msrRenderPreview(g, dims) {
  const tool = currentTool;
  const pts  = msrActiveDrawPts;
  const prev = msrPreviewPt;

  // Reference line preview during dialog
  if (szDialogEl && !szDialogEl.hidden && szActiveTab === 'ref' && szRefState) {
    const rpts = [...szRefState.pts, ...(prev ? [prev] : [])];
    if (rpts.length >= 2) {
      const l = document.createElementNS(MSR_SVG_NS, 'line');
      l.setAttribute('x1', rpts[0].x * canvas.width); l.setAttribute('y1', rpts[0].y * canvas.height);
      l.setAttribute('x2', rpts[1].x * canvas.width); l.setAttribute('y2', rpts[1].y * canvas.height);
      l.classList.add('msr-preview'); l.style.stroke = '#4da3ff';
      g.appendChild(l);
    }
    return;
  }

  if (!['linear', 'area', 'scale-zone'].includes(tool) || pts.length === 0) return;

  const all = [...pts, ...(prev ? [prev] : [])];

  if (tool === 'linear' && all.length >= 2) {
    const l = document.createElementNS(MSR_SVG_NS, 'line');
    l.setAttribute('x1', all[0].x * canvas.width); l.setAttribute('y1', all[0].y * canvas.height);
    l.setAttribute('x2', all[1].x * canvas.width); l.setAttribute('y2', all[1].y * canvas.height);
    l.classList.add('msr-preview');
    g.appendChild(l);
    // Live distance label
    if (dims) {
      const zone = msrFindZone(all[0].x, all[0].y, currentPage);
      if (zone) {
        const d = msrNormToMeters(all[0].x, all[0].y, all[1].x, all[1].y, dims.width, dims.height, zone.mpp);
        const lbl = document.createElementNS(MSR_SVG_NS, 'text');
        lbl.setAttribute('x', (all[0].x + all[1].x) / 2 * canvas.width);
        lbl.setAttribute('y', (all[0].y + all[1].y) / 2 * canvas.height - 9);
        lbl.setAttribute('text-anchor', 'middle');
        lbl.classList.add('msr-lbl'); lbl.style.opacity = '0.55';
        lbl.textContent = `${d.toFixed(3)} m`;
        g.appendChild(lbl);
      }
    }
  }

  if ((tool === 'area' || tool === 'scale-zone') && all.length >= 2) {
    const polyEl = document.createElementNS(MSR_SVG_NS, 'polyline');
    polyEl.setAttribute('points', all.map(p => `${p.x * canvas.width},${p.y * canvas.height}`).join(' '));
    polyEl.classList.add('msr-preview');
    if (tool === 'scale-zone') polyEl.style.stroke = '#4da3ff';
    g.appendChild(polyEl);

    // Snap ring on first vertex when closeable
    if (pts.length >= 3) {
      const ring = document.createElementNS(MSR_SVG_NS, 'circle');
      ring.setAttribute('cx', pts[0].x * canvas.width); ring.setAttribute('cy', pts[0].y * canvas.height);
      ring.setAttribute('r', 12); ring.classList.add('msr-snap-ring');
      if (tool === 'scale-zone') ring.style.stroke = '#4da3ff';
      g.appendChild(ring);
    }

    // Live area label
    if (tool === 'area' && all.length >= 3 && dims) {
      const cx = all.reduce((s, p) => s + p.x, 0) / all.length;
      const cy = all.reduce((s, p) => s + p.y, 0) / all.length;
      const zone = msrFindZone(cx, cy, currentPage);
      if (zone) {
        const a = msrNormToSqMeters(all, dims.width, dims.height, zone.mpp);
        const lbl = document.createElementNS(MSR_SVG_NS, 'text');
        lbl.setAttribute('x', cx * canvas.width); lbl.setAttribute('y', cy * canvas.height);
        lbl.setAttribute('text-anchor', 'middle'); lbl.setAttribute('dominant-baseline', 'central');
        lbl.classList.add('msr-lbl'); lbl.style.opacity = '0.55';
        lbl.textContent = `${a.toFixed(3)} m²`;
        g.appendChild(lbl);
      }
    }
  }
}

// Lightweight refresh — only redraws the measurement layer
function msrRedrawOnly() {
  const measG = overlay?.querySelector('.msr-meas-g');
  if (!measG) { redrawRegions(); return; }
  measG.innerHTML = '';
  msrRenderMeasurements(measG);
}

function msrUpdateScaleLabel() {
  if (!activeScaleLblEl) return;
  const zones = scaleZonesByPage[currentPage] || [];
  activeScaleLblEl.textContent = zones.length === 0 ? 'No scale set'
    : zones.length === 1 ? `Scale: ${zones[0].label}`
    : `${zones.length} scale zones`;
}

// ── Sidebar resize handle ────────────────────────────────────────────────────
(function initSidebarResizer() {
  const resizer = document.getElementById('sidebar-resizer');
  const sidebarEl = document.getElementById('sidebar');
  if (!resizer || !sidebarEl) return;

  const MIN_WIDTH = 80;
  const MAX_WIDTH = 340;

  let isResizing = false;
  let startX = 0;
  let startWidth = 0;

  resizer.addEventListener('mousedown', (e) => {
    isResizing = true;
    startX = e.clientX;
    startWidth = sidebarEl.offsetWidth;
    resizer.classList.add('is-resizing');
    document.body.style.cursor = 'ew-resize';
    document.body.style.userSelect = 'none';
    e.preventDefault();
  });

  window.addEventListener('mousemove', (e) => {
    if (!isResizing) return;
    const newWidth = Math.min(Math.max(startWidth + (e.clientX - startX), MIN_WIDTH), MAX_WIDTH);
    sidebarEl.style.width = newWidth + 'px';
    sidebarEl.style.flex = `0 0 ${newWidth}px`;
  });

  window.addEventListener('mouseup', () => {
    if (!isResizing) return;
    isResizing = false;
    resizer.classList.remove('is-resizing');
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
  });
})();
