// pdf_splitter.js
// Staging area · PDF splitting · ZIP / CSV / JSON export
// Drag-and-drop row reordering · Sequence prefix
// Depends on: pdf_viewer.js globals, pdf-lib, JSZip

(function () {
  'use strict';

  // ─────────────────────────────────────────────────────────────────────
  // STATE
  // ─────────────────────────────────────────────────────────────────────

  let stagingData      = [];    // display order — may differ from PDF page order after drag
  let originalOrder    = [];    // page numbers in original order — for reset
  let currentFilter    = 'all';
  let filenameTemplate = '{sheet_id} - {description} - {issue_id}';
  let useSeqPrefix     = true;
  let seqSeparator     = ' - '; // between sequence number and filename

  const PRESETS = [
    { label: 'Sheet · Description · Revision',           value: '{sheet_id} - {description} - {issue_id}' },
    { label: 'Project · Sheet · Description · Revision', value: '{project_id} - {sheet_id} - {description} - {issue_id}' },
    { label: 'Sheet · Description',                      value: '{sheet_id} - {description}' },
    { label: 'Custom…',                                  value: '__custom__' },
  ];

  // ─────────────────────────────────────────────────────────────────────
  // DRAG-AND-DROP STATE
  // ─────────────────────────────────────────────────────────────────────

  let dragSrcIndex  = -1;
  let dragOverIndex = -1;
  let dragDropAbove = true;

  // ─────────────────────────────────────────────────────────────────────
  // FILENAME HELPERS
  // ─────────────────────────────────────────────────────────────────────

  function todayYYMMDD() {
    const d  = new Date();
    const yy = String(d.getFullYear()).slice(2);
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return yy + mm + dd;
  }

  function sanitize(str) {
    return String(str || '')
      .replace(/[/\\:*?"<>|]/g, '')
      .replace(/\s{2,}/g, ' ')
      .trim();
  }

  function generateFilename(sheet, template, seqIndex) {
    const doc = window.documentDetails || {};
    const map = {
      sheet_id:          sheet.sheet_id          || '',
      description:       sheet.description       || '',
      issue_id:          sheet.issue_id          || '',
      date:              sheet.date              || '',
      issue_description: sheet.issue_description || '',
      project_id:        doc.project_id          || '',
      prepared_by:       doc.prepared_by         || '',
    };

    let result = template.replace(/\{(\w+)\}/g, (_, k) => map[k] !== undefined ? map[k] : '');

    // Collapse consecutive separators left by empty tokens  e.g. " -  - " → " - "
    result = result
      .replace(/([ \t]*-[ \t]*){2,}/g, ' - ')
      .replace(/^[ \t\-_]+|[ \t\-_]+$/g, '')
      .trim();

    const base = (sanitize(result) || `page-${sheet.page}`) + '.pdf';

    if (useSeqPrefix && seqIndex !== undefined) {
      const pad = String(seqIndex + 1).padStart(3, '0');
      return pad + seqSeparator + base;
    }
    return base;
  }

  function getStatus(sheet) {
    const allBlank = !sheet.sheet_id && !sheet.description && !sheet.issue_id;
    if (allBlank) return 'failed';
    if (!sheet.sheet_id || !sheet.description) return 'review';
    return 'ready';
  }

  function refreshAllFilenames() {
    stagingData.forEach((s, i) => {
      if (!s.filenameOverride) s.filename = generateFilename(s, filenameTemplate, i);
      s.status = getStatus(s);
    });
    renderTable();
  }

  // ─────────────────────────────────────────────────────────────────────
  // BUILD STAGING DATA
  // ─────────────────────────────────────────────────────────────────────

  function buildStagingData() {
    const fn         = window.getCanonicalExportData;
    const exportData = fn ? fn() : { sheets: [] };

    stagingData = exportData.sheets.map((sheet, i) => ({
      page:              sheet.page,
      sheet_id:          sheet.sheet_id          || '',
      description:       sheet.description       || '',
      issue_id:          sheet.issue_id          || '',
      date:              sheet.date              || '',
      issue_description: sheet.issue_description || '',
      filename:          generateFilename(sheet, filenameTemplate, i),
      filenameOverride:  false,
      status:            getStatus(sheet),
      included:          true,
    }));

    originalOrder = stagingData.map(s => s.page); // snapshot for reset
  }

  function resetOrder() {
    stagingData.sort((a, b) => originalOrder.indexOf(a.page) - originalOrder.indexOf(b.page));
    refreshAllFilenames();
  }

  // ─────────────────────────────────────────────────────────────────────
  // RENDER TABLE
  // ─────────────────────────────────────────────────────────────────────

  function esc(str) {
    return String(str || '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function renderTable() {
    const tbody = document.getElementById('sp-tbody');
    if (!tbody) return;

    const rows = currentFilter === 'all'    ? stagingData
               : currentFilter === 'issues' ? stagingData.filter(s => s.status !== 'ready')
               :                              stagingData.filter(s => s.status === 'ready');

    tbody.innerHTML = '';

    rows.forEach(sheet => {
      const trueIndex = stagingData.indexOf(sheet);

      const tr = document.createElement('tr');
      tr.dataset.page  = sheet.page;
      tr.dataset.index = trueIndex;
      tr.draggable     = true;
      if (sheet.status === 'review') tr.classList.add('sp-row-warn');
      if (sheet.status === 'failed') tr.classList.add('sp-row-fail');

      const bCls   = { ready: 'sp-ok', review: 'sp-warn', failed: 'sp-fail' }[sheet.status];
      const bLabel = { ready: 'Ready', review: 'Review',  failed: 'OCR failed' }[sheet.status];

      tr.innerHTML = `
        <td style="width:28px;padding:0 2px;text-align:center">
          <span class="sp-drag-handle" title="Drag to reorder">⠿</span>
        </td>
        <td class="sp-td-check"><input type="checkbox" class="sp-cb" data-page="${sheet.page}" ${sheet.included ? 'checked' : ''}></td>
        <td class="sp-td-pg">${sheet.page}</td>
        <td><input class="sp-cell" data-page="${sheet.page}" data-field="sheet_id"    value="${esc(sheet.sheet_id)}"    placeholder="?"></td>
        <td><input class="sp-cell" data-page="${sheet.page}" data-field="description" value="${esc(sheet.description)}" placeholder="?"></td>
        <td class="sp-td-rev"><input class="sp-cell" data-page="${sheet.page}" data-field="issue_id" value="${esc(sheet.issue_id)}" placeholder="?"></td>
        <td><input class="sp-cell sp-filename" id="sp-fp-${sheet.page}" value="${esc(sheet.filename)}" placeholder="filename.pdf" style="font-family:var(--mono);font-size:11px;width:100%;min-width:200px;"></td>
        <td class="sp-td-status"><span class="sp-badge ${bCls}">${bLabel}</span></td>
      `;

      // ── Drag events ────────────────────────────────────────────────
      tr.addEventListener('dragstart', e => {
        dragSrcIndex = trueIndex;
        tr.classList.add('sp-dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', String(trueIndex)); // required for Firefox
      });

      tr.addEventListener('dragover', e => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';

        const rect = tr.getBoundingClientRect();
        dragDropAbove = e.clientY < rect.top + rect.height / 2;
        dragOverIndex = trueIndex;

        tbody.querySelectorAll('tr').forEach(r => r.classList.remove('sp-drop-above', 'sp-drop-below'));
        tr.classList.add(dragDropAbove ? 'sp-drop-above' : 'sp-drop-below');
      });

      tr.addEventListener('dragleave', () => {
        tr.classList.remove('sp-drop-above', 'sp-drop-below');
      });

      tr.addEventListener('drop', e => {
        e.preventDefault();
        tbody.querySelectorAll('tr').forEach(r => r.classList.remove('sp-drop-above', 'sp-drop-below'));

        if (dragSrcIndex === -1 || dragSrcIndex === dragOverIndex) return;

        // Remove dragged item from array
        const [moved] = stagingData.splice(dragSrcIndex, 1);

        // After splice, find drop target's new index
        const adjustedTarget = dragOverIndex > dragSrcIndex ? dragOverIndex - 1 : dragOverIndex;
        const insertAt = dragDropAbove ? adjustedTarget : adjustedTarget + 1;
        stagingData.splice(Math.max(0, Math.min(insertAt, stagingData.length)), 0, moved);

        dragSrcIndex = dragOverIndex = -1;
        refreshAllFilenames(); // re-renders table + updates seq numbers
      });

      tr.addEventListener('dragend', () => {
        tr.classList.remove('sp-dragging');
        tbody.querySelectorAll('tr').forEach(r => r.classList.remove('sp-drop-above', 'sp-drop-below'));
        dragSrcIndex = dragOverIndex = -1;
      });

      // ── Inline edit events ─────────────────────────────────────────
      tr.querySelectorAll('.sp-cb').forEach(cb => {
        cb.addEventListener('change', () => {
          const s = stagingData.find(x => x.page === sheet.page);
          if (s) s.included = cb.checked;
          refreshStats();
        });
      });

      tr.querySelectorAll('.sp-cell:not(.sp-filename)').forEach(input => {
        input.addEventListener('input', () => {
          const s = stagingData.find(x => x.page === sheet.page);
          if (!s) return;
          s[input.dataset.field] = input.value;
          s.status = getStatus(s);
          if (!s.filenameOverride) {
            s.filename = generateFilename(s, filenameTemplate, stagingData.indexOf(s));
            const fp = document.getElementById(`sp-fp-${sheet.page}`);
            if (fp) fp.value = s.filename;
          }
          refreshStats();
        });
      });

      tr.querySelectorAll('.sp-filename').forEach(input => {
        input.addEventListener('input', () => {
          const s = stagingData.find(x => x.page === sheet.page);
          if (!s) return;
          s.filename         = input.value;
          s.filenameOverride = true;
          refreshStats();
        });
      });

      tbody.appendChild(tr);
    });

    refreshStats();
  }

  function refreshStats() {
    const inc    = stagingData.filter(s => s.included);
    const ready  = inc.filter(s => s.status === 'ready').length;
    const review = inc.filter(s => s.status === 'review').length;
    const failed = inc.filter(s => s.status === 'failed').length;

    const el = document.getElementById('sp-stats');
    if (el) el.textContent = `${ready} ready · ${review} needs review · ${failed} OCR failed · ${inc.length} selected`;

    const btn = document.getElementById('sp-btn-zip');
    if (btn) btn.textContent = `Download ZIP (${inc.length} files)`;
  }

  // ─────────────────────────────────────────────────────────────────────
  // SEQUENCE PREFIX CONTROLS
  // ─────────────────────────────────────────────────────────────────────

  function onSeqToggle(checked) {
    useSeqPrefix = checked;
    const sepRow   = document.getElementById('sp-seq-sep-row');
    const seqLabel = document.getElementById('sp-seq-label');
    if (sepRow)   sepRow.style.display = checked ? 'flex' : 'none';
    if (seqLabel) seqLabel.className   = checked ? 'sp-seq-on' : '';
    refreshAllFilenames();
  }

  function onSeqSepChange(val) {
    seqSeparator = val;
    if (useSeqPrefix) refreshAllFilenames();
  }

  // ─────────────────────────────────────────────────────────────────────
  // TEMPLATE CONTROLS
  // ─────────────────────────────────────────────────────────────────────

  function onPresetChange(val) {
    const customRow = document.getElementById('sp-custom-row');
    if (val === '__custom__') {
      if (customRow) customRow.style.display = 'flex';
    } else {
      filenameTemplate = val;
      if (customRow) customRow.style.display = 'none';
      refreshAllFilenames();
    }
  }

  function onCustomInput(val) {
    filenameTemplate = val || '{sheet_id} - {description} - {issue_id}';
    refreshAllFilenames();
  }

  function toggleSelectAll(checked) {
    stagingData.forEach(s => s.included = checked);
    document.querySelectorAll('.sp-cb').forEach(cb => cb.checked = checked);
    refreshStats();
  }

  // ─────────────────────────────────────────────────────────────────────
  // STAGING PANEL OPEN / CLOSE
  // ─────────────────────────────────────────────────────────────────────

  function openPanel() {
    const panel = document.getElementById('staging-panel');
    if (panel) panel.style.display = 'flex';
    renderTable();
  }

  function closePanel() {
    const panel = document.getElementById('staging-panel');
    if (panel) panel.style.display = 'none';
  }

  // ─────────────────────────────────────────────────────────────────────
  // PDF SPLIT + ZIP DOWNLOAD
  // ─────────────────────────────────────────────────────────────────────

  async function downloadZip() {
    const included = stagingData.filter(s => s.included);
    if (!included.length) { alert('No sheets selected.'); return; }

    const showP = window.showProcessing || (m => console.log(m));
    const hideP = window.hideProcessing || (() => {});
    const dlB   = window.downloadBlob;

    showP('Preparing ZIP…');

    try {
      const zip       = new JSZip();
      const isMulti   = window.isMultiPdfMode;
      const multiDocs = window.multiPdfDocs || [];
      const rawBytes  = window.pdfRawBytes;

      for (let i = 0; i < included.length; i++) {
        const sheet = included[i];
        showP(`Splitting sheet ${i + 1} / ${included.length}: ${sheet.filename}`);

        let pageBytes;

        if (isMulti) {
          let found = false;
          for (const pdf of multiDocs) {
            const localPage = sheet.page - pdf.pageOffset;
            if (localPage >= 1 && localPage <= pdf.pageCount) {
              if (pdf.pageCount === 1) {
                pageBytes = pdf.rawBytes; // single-page source — copy unchanged, no re-encoding
              } else {
                const { PDFDocument } = PDFLib;
                const srcDoc = await PDFDocument.load(pdf.rawBytes);
                const newDoc = await PDFDocument.create();
                const [pg]   = await newDoc.copyPages(srcDoc, [localPage - 1]);
                newDoc.addPage(pg);
                pageBytes = await newDoc.save();
              }
              found = true;
              break;
            }
          }
          if (!found) { console.warn(`⚠ Page ${sheet.page} not found in source.`); continue; }
        } else {
          const { PDFDocument } = PDFLib;
          const srcDoc = await PDFDocument.load(rawBytes);
          const newDoc = await PDFDocument.create();
          const [pg]   = await newDoc.copyPages(srcDoc, [sheet.page - 1]);
          newDoc.addPage(pg);
          pageBytes = await newDoc.save();
        }

        zip.file(sheet.filename, pageBytes);
        if (i % 5 === 4) await new Promise(r => setTimeout(r, 0)); // stay responsive
      }

      showP('Compressing ZIP…');

      const blob = await zip.generateAsync({
        type: 'blob',
        compression: 'DEFLATE',
        compressionOptions: { level: 1 }, // PDFs are already compressed
      });

      const doc  = window.documentDetails || {};
      const name = todayYYMMDD() + '_' + sanitize(doc.project_id || 'drawings') + '.zip';

      if (dlB) { dlB(blob, name); } else {
        const url = URL.createObjectURL(blob);
        const a   = Object.assign(document.createElement('a'), { href: url, download: name });
        document.body.appendChild(a); a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }

    } catch (err) {
      console.error('ZIP error:', err);
      alert('Error creating ZIP: ' + err.message);
    } finally {
      hideP();
    }
  }

  // ─────────────────────────────────────────────────────────────────────
  // CSV EXPORT  — manifest with order, page, filename columns
  // ─────────────────────────────────────────────────────────────────────

  function exportCSV() {
    const included = stagingData.filter(s => s.included);
    if (!included.length) { alert('No sheets selected.'); return; }

    const headers = ['order', 'page', 'filename', 'sheet_id', 'description', 'issue_id', 'date', 'issue_description', 'status'];
    const q       = v => `"${String(v ?? '').replace(/"/g, '""')}"`;

    const csv = [
      headers.join(','),
      ...included.map((s, i) =>
        [i + 1, s.page, s.filename, s.sheet_id, s.description, s.issue_id, s.date, s.issue_description, s.status].map(q).join(',')
      ),
    ].join('\n');

    const doc  = window.documentDetails || {};
    const name = todayYYMMDD() + '_' + sanitize(doc.project_id || 'drawings') + '.csv';
    if (window.downloadBlob) window.downloadBlob(new Blob([csv], { type: 'text/csv' }), name);
  }

  // ─────────────────────────────────────────────────────────────────────
  // JSON EXPORT  — manifest with order, page, filename
  // ─────────────────────────────────────────────────────────────────────

  function exportJSON() {
    const included = stagingData.filter(s => s.included);
    if (!included.length) { alert('No sheets selected.'); return; }

    const doc  = window.documentDetails || {};
    const data = {
      exported_at: new Date().toISOString(),
      project_id:  doc.project_id  || '',
      prepared_by: doc.prepared_by || '',
      sheets: included.map((s, i) => ({
        order:             i + 1,     // position in export (post-drag)
        page:              s.page,    // source PDF page number
        filename:          s.filename,
        sheet_id:          s.sheet_id,
        description:       s.description,
        issue_id:          s.issue_id,
        date:              s.date,
        issue_description: s.issue_description,
        status:            s.status,
      })),
    };

    const name = todayYYMMDD() + '_' + sanitize(doc.project_id || 'drawings') + '.json';
    if (window.downloadBlob) window.downloadBlob(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }), name);
  }

  // ─────────────────────────────────────────────────────────────────────
  // ENTRY POINT — wired to the ✦ Split & Name button
  // ─────────────────────────────────────────────────────────────────────

  async function openStagingFromExtract() {
    if (!window.pdfDoc) { alert('No PDF loaded.'); return; }

    const showP = window.showProcessing || (() => {});
    const hideP = window.hideProcessing || (() => {});

    showP('Extracting all pages — this may take a moment…');

    try {
      if (window.applyTemplatesToAllPages) await window.applyTemplatesToAllPages(true);
      buildStagingData();
      openPanel();
    } catch (err) {
      console.error('Staging error:', err);
      alert('Extraction error: ' + err.message);
    } finally {
      hideP();
    }
  }

  // ─────────────────────────────────────────────────────────────────────
  // INIT
  // ─────────────────────────────────────────────────────────────────────

  function init() {
    // Template preset dropdown
    const presetSel = document.getElementById('sp-template-select');
    if (presetSel) {
      PRESETS.forEach(p => {
        const opt = Object.assign(document.createElement('option'), { value: p.value, textContent: p.label });
        presetSel.appendChild(opt);
      });
      presetSel.addEventListener('change', () => onPresetChange(presetSel.value));
    }

    // Custom template input
    const customInput = document.getElementById('sp-custom-input');
    if (customInput) {
      customInput.value = filenameTemplate;
      customInput.addEventListener('input', () => onCustomInput(customInput.value));
    }

    // Sequence prefix toggle
    document.getElementById('sp-seq-toggle')?.addEventListener('change', e => onSeqToggle(e.target.checked));

    // Sequence separator
    document.getElementById('sp-seq-sep')?.addEventListener('change', e => onSeqSepChange(e.target.value));

    // Filter dropdown
    document.getElementById('sp-filter-select')?.addEventListener('change', e => {
      currentFilter = e.target.value; renderTable();
    });

    // Select-all
    document.getElementById('sp-select-all')?.addEventListener('change', e => toggleSelectAll(e.target.checked));

    // Action buttons
    document.getElementById('sp-btn-zip')?.addEventListener('click', downloadZip);
    document.getElementById('sp-btn-csv')?.addEventListener('click', exportCSV);
    document.getElementById('sp-btn-json')?.addEventListener('click', exportJSON);
    document.getElementById('sp-btn-close')?.addEventListener('click', closePanel);
    document.getElementById('sp-btn-reset')?.addEventListener('click', resetOrder);
    document.getElementById('btn-split-name')?.addEventListener('click', openStagingFromExtract);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

  // Public API
  window.openStagingFromExtract = openStagingFromExtract;
  window.closeStaging           = closePanel;
  window.exportStagingCSV       = exportCSV;
  window.exportStagingJSON      = exportJSON;

})();
