/**
 * Loan Recovery PDF Converter — script.js
 * ================================================
 * Libraries expected (loaded in index.html):
 *   XLSX (SheetJS)  — Excel parsing
 *   jsPDF           — PDF generation
 *   jspdf-autotable — Table plugin for jsPDF
 * ================================================
 */

/* ═══════════════════════════════════════════════
   1.  COLUMN MAPPING CONFIGURATION
   ═══════════════════════════════════════════════ */

/**
 * Each entry defines one TARGET field.
 * aliases[] = possible Excel column header names (case-insensitive, partial match).
 */
const FIELD_DEFS = [
  {
    key: 'loanNo',
    label: 'Loan No',
    aliases: ['loan no', 'loan number', 'loan num', 'loan_no', 'loanno', 'loan account', 'account no', 'account number'],
  },
  {
    key: 'customerName',
    label: 'Customer Name',
    aliases: ['customer name', 'customer', 'cust name', 'name', 'borrower', 'applicant name'],
  },
  {
    key: 'emi',
    label: 'EMI',
    aliases: ['emi', 'emi amount', 'monthly emi', 'instalment', 'installment'],
  },
  {
    key: 'bkt',
    label: 'BKT',
    aliases: ['bkt', 'bucket', 'opening bkt', 'opening bucket', 'dpd bucket', 'bucket no'],
  },
  {
    key: 'address',
    label: 'Address',
    aliases: ['permanent address', 'address', 'customer address', 'residential address', 'perm address', 'addr'],
  },
  {
    key: 'mobile',
    label: 'Mobile',
    aliases: ['contact number', 'mobile', 'phone', 'contact no', 'mobile no', 'customer mobile', 'mob no', 'contact', 'phone no'],
  },
  {
    key: 'ref1Name',
    label: 'Ref1 Name',
    aliases: ['reference 1', 'ref1 name', 'ref 1', 'reference1', 'reference1 name', 'ref1', 'guarantor 1', 'guarantor name'],
  },
  {
    key: 'ref1Mobile',
    label: 'Ref1 Mobile',
    aliases: ['reference1 mobile', 'reference 1 mobile', 'ref1 mobile', 'ref 1 mobile', 'ref1mobile', 'guarantor 1 mobile', 'guarantor mobile'],
  },
  {
    key: 'agencyName',
    label: 'Agency Name',
    aliases: ['agency', 'agency name', 'dma', 'dma name', 'agency/dma', 'agency/dma name'],
  },
  {
    key: 'pos',
    label: 'POS',
    aliases: ['pos', 'principal outstanding'],
  },
  {
    key: 'openingOD',
    label: 'Opening OD',
    aliases: ['opening od', 'opning od', 'opning o', 'open od'],
  },
  {
    key: 'totalOD',
    label: 'Total OD',
    aliases: ['total od', 'od amount'],
  },
  {
    key: 'piDues',
    label: 'PI Dues',
    aliases: ['pi dues', 'pidues', 'pi due'],
  },
  {
    key: 'make',
    label: 'Make',
    aliases: ['make', 'vehicle make', 'model', 'asset make', 'make name'],
  },
  {
    key: 'regNo',
    label: 'Reg No',
    aliases: ['reg no', 'registration no', 'registration number', 'vehicle no', 'reg_no', 'reg. no.', 'vehicle number'],
  },
  {
    key: 'ref2Name',
    label: 'Ref2 Name',
    aliases: ['reference 2', 'ref2 name', 'ref 2', 'reference2', 'reference2 name', 'ref2', 'guarantor 2', 'guarantor 2 name'],
  },
  {
    key: 'ref2Mobile',
    label: 'Ref2 Mobile',
    aliases: ['reference2 mobile', 'reference 2 mobile', 'ref2 mobile', 'ref 2 mobile', 'ref2mobile', 'guarantor 2 mobile'],
  },
  {
    key: 'dealerName',
    label: 'DLR Name',
    aliases: ['dealer', 'dealer name', 'showroom', 'showroom name', 'dealer/showroom', 'dlr', 'dlr name'],
  },
];

/* ═══════════════════════════════════════════════
   2.  STATE
   ═══════════════════════════════════════════════ */
let parsedRows = [];   // raw data rows from Excel
let mappedKeys = {};   // { fieldKey: excelColumnName }

/* ═══════════════════════════════════════════════
   3.  DOM REFERENCES
   ═══════════════════════════════════════════════ */
const dropZone       = document.getElementById('dropZone');
const fileInput      = document.getElementById('fileInput');
const fileInfo       = document.getElementById('fileInfo');
const fileNameEl     = document.getElementById('fileName');
const recordCountEl  = document.getElementById('recordCount');
const clearFileBtn   = document.getElementById('clearFile');
const mappingSection = document.getElementById('mapping-section');
const mappingGrid    = document.getElementById('mappingGrid');
const searchSection  = document.getElementById('search-section');
const searchInput    = document.getElementById('searchInput');
const searchBtn      = document.getElementById('searchBtn');
const searchError    = document.getElementById('searchError');
const searchResult   = document.getElementById('searchResult');
const resultGrid     = document.getElementById('resultGrid');
const downloadSingleBtn = document.getElementById('downloadSingleBtn');
const generateSection= document.getElementById('generate-section');
const agencySelect   = document.getElementById('agencySelect');
const statRecords    = document.getElementById('statRecords');
const statPages      = document.getElementById('statPages');
const generateBtn    = document.getElementById('generateBtn');
const genBtnText     = document.getElementById('genBtnText');
const spinner        = document.getElementById('spinner');
const progressWrap   = document.getElementById('progressWrap');
const progressFill   = document.getElementById('progressFill');
const progressLabel  = document.getElementById('progressLabel');

/* ═══════════════════════════════════════════════
   4.  DRAG & DROP / FILE INPUT
   ═══════════════════════════════════════════════ */

dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('dragover');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});
dropZone.addEventListener('click', e => {
  // Only trigger if the click isn't on the label/button
  if (e.target.tagName !== 'LABEL') fileInput.click();
});
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
});
clearFileBtn.addEventListener('click', resetAll);

/* ═══════════════════════════════════════════════
   5.  FILE HANDLER
   ═══════════════════════════════════════════════ */

function handleFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx','xls','csv'].includes(ext)) {
    alert('Please upload an Excel file (.xlsx, .xls) or CSV.');
    return;
  }

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb   = XLSX.read(data, { type: 'array', cellDates: true });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: '' });

      if (!json.length) { alert('The file appears to be empty.'); return; }

      parsedRows = json;
      mappedKeys = autoMapColumns(json[0]);

      // Show file info
      fileNameEl.textContent = file.name;
      recordCountEl.textContent = `${json.length.toLocaleString()} records`;
      fileInfo.classList.remove('hidden');

      // Show sections
      renderMappingGrid(mappedKeys);
      
      // Populate Agency Dropdown
      const agencies = [...new Set(parsedRows.map(r => get(r, 'agencyName') || 'Unknown Agency'))].filter(Boolean).sort();
      agencySelect.innerHTML = '<option value="ALL">-- Generate All (Download as ZIP) --</option>';
      agencies.forEach(ag => {
        const opt = document.createElement('option');
        opt.value = ag;
        opt.textContent = ag;
        agencySelect.appendChild(opt);
      });

      mappingSection.classList.remove('hidden');
      searchSection.classList.remove('hidden');
      generateSection.classList.remove('hidden');

      // Stats
      const updateStats = () => {
        const sel = agencySelect.value;
        const filtered = sel === 'ALL' ? parsedRows : parsedRows.filter(r => (get(r, 'agencyName') || 'Unknown Agency') === sel);
        statRecords.textContent = filtered.length.toLocaleString();
        statPages.textContent   = '~' + Math.ceil(filtered.length / 35);
      };
      updateStats();
      agencySelect.onchange = updateStats;

    } catch (err) {
      console.error(err);
      alert('Failed to read the file. Please make sure it is a valid Excel/CSV file.');
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ═══════════════════════════════════════════════
   6.  INTELLIGENT COLUMN MAPPING
   ═══════════════════════════════════════════════ */

/**
 * Takes the first row object (keys = Excel column headers) and
 * returns { fieldKey: matchedExcelColumnName | null }.
 */
function autoMapColumns(sampleRow) {
  const excelCols = Object.keys(sampleRow);
  const result    = {};

  for (const fieldDef of FIELD_DEFS) {
    result[fieldDef.key] = findBestMatch(fieldDef.aliases, excelCols);
  }
  return result;
}

/**
 * Finds the best matching Excel column for a list of aliases.
 * Strategy:
 *   1. Exact match (case-insensitive)
 *   2. Alias is substring of column name
 *   3. Column name is substring of alias
 */
function findBestMatch(aliases, excelCols) {
  const normalize = s => s.toLowerCase().replace(/[\s_\-\.]+/g, ' ').trim();

  for (const alias of aliases) {
    const a = normalize(alias);
    // Exact
    const exact = excelCols.find(c => normalize(c) === a);
    if (exact) return exact;
  }
  for (const alias of aliases) {
    const a = normalize(alias);
    // Alias in column
    const partial = excelCols.find(c => normalize(c).includes(a));
    if (partial) return partial;
  }
  for (const alias of aliases) {
    const a = normalize(alias);
    // Column in alias
    const reverse = excelCols.find(c => a.includes(normalize(c)) && normalize(c).length > 3);
    if (reverse) return reverse;
  }
  return null;
}

/* ═══════════════════════════════════════════════
   7.  MAPPING PREVIEW GRID
   ═══════════════════════════════════════════════ */

function renderMappingGrid(map) {
  mappingGrid.innerHTML = '';
  for (const fieldDef of FIELD_DEFS) {
    const matched = map[fieldDef.key];
    const div     = document.createElement('div');
    div.className = `map-item ${matched ? 'found' : 'missing'}`;
    div.innerHTML = `
      <div class="map-field">${fieldDef.label}</div>
      <div class="map-source">
        ${matched
          ? '✓ → ' + matched
          : '✕ Missing — will be blank'}
      </div>`;
    mappingGrid.appendChild(div);
  }
}

/* ═══════════════════════════════════════════════
   8.  EXTRACT VALUE FROM ROW
   ═══════════════════════════════════════════════ */

function get(row, key) {
  const col = mappedKeys[key];
  if (!col) return '';
  const val = row[col];
  if (val === undefined || val === null) return '';
  return String(val).trim();
}

/* ═══════════════════════════════════════════════
   9.  PDF GENERATION
   ═══════════════════════════════════════════════ */

generateBtn.addEventListener('click', () => {
  if (!parsedRows.length) return;
  generatePDF();
});

async function generatePDF() {
  // UI: loading state
  genBtnText.classList.add('hidden');
  spinner.classList.remove('hidden');
  progressWrap.classList.remove('hidden');
  generateBtn.disabled = true;
  setProgress(0, 'Preparing data…');

  await sleep(30);

  try {
    const { jsPDF } = window.jspdf;
    const selectedAgency = agencySelect.value;

    // ── Group by Agency ──
    const groups = {};
    for (const r of parsedRows) {
      let rawAgency = get(r, 'agencyName') || 'Unknown Agency';
      if (selectedAgency !== 'ALL' && rawAgency !== selectedAgency) continue;

      let safeAgency = rawAgency.replace(/[^a-z0-9_\-\s]/gi, '_').trim();
      if (!groups[safeAgency]) groups[safeAgency] = [];
      groups[safeAgency].push(r);
    }

    const agencyNames = Object.keys(groups);
    if (agencyNames.length === 0) {
      alert('No records found for the selected agency.');
      genBtnText.classList.remove('hidden');
      spinner.classList.add('hidden');
      progressWrap.classList.add('hidden');
      generateBtn.disabled = false;
      return;
    }

    const useZip = selectedAgency === 'ALL';
    let zip;
    if (useZip) zip = new JSZip();

    const timestamp = new Date().toISOString().slice(0,10);

    for (let gIndex = 0; gIndex < agencyNames.length; gIndex++) {
      const agencyName = agencyNames[gIndex];
      const rows = groups[agencyName];
      const totalRows = rows.length;

      setProgress(
        10 + (gIndex / agencyNames.length) * 80,
        `Processing ${agencyName} (${gIndex + 1}/${agencyNames.length})…`
      );
      await sleep(30);

      const doc = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: 'a4',
      });

      const MARGIN   = 3;
      const PAGE_W   = 297;
      const PAGE_H   = 210;
      const USABLE_W = PAGE_W - MARGIN * 2;

      /* ── Column definitions ── */
      const COLS = [
        { w: 22, hdr: 'Loan No',        get: r => get(r, 'loanNo') },
        { w: 25, hdr: 'Customer Name',  get: r => get(r, 'customerName') },
        { w: 6,  hdr: 'BKT',            get: r => get(r, 'bkt') },
        { w: 11, hdr: 'POS',            get: r => get(r, 'pos') },
        {
          w: 13,
          hdr: 'OD Details',
          get: r => {
            const opn = get(r, 'openingOD');
            const tot = get(r, 'totalOD');
            const parts = [];
            if (opn) parts.push('Opn: ' + opn);
            if (tot) parts.push('Tot: ' + tot);
            return parts.join('\n');
          }
        },
        { w: 8,  hdr: 'EMI',            get: r => get(r, 'emi') },
        { w: 10, hdr: 'PI Dues',        get: r => get(r, 'piDues') },
        { w: 44, hdr: 'Permanent Address', get: r => get(r, 'address') },
        {
          w: 44,
          hdr: 'Contact / Reference',
          get: r => {
            const mob   = get(r, 'mobile');
            const ref1n = get(r, 'ref1Name');
            const ref1m = get(r, 'ref1Mobile');
            const ref2n = get(r, 'ref2Name');
            const ref2m = get(r, 'ref2Mobile');
            const parts = [];
            if (mob)   parts.push('Mob: ' + mob);
            if (ref1n) parts.push('Ref1: ' + ref1n);
            if (ref1m) parts.push('RefMob: ' + ref1m);
            if (ref2n) parts.push('Ref2: ' + ref2n);
            if (ref2m) parts.push('Ref2Mob: ' + ref2m);
            return parts.join('\n');
          }
        },
        { w: 13, hdr: 'Make', get: r => get(r, 'make') },
        { w: 16, hdr: 'REG NO', get: r => get(r, 'regNo') },
        { w: 16, hdr: 'DLR Name', get: r => get(r, 'dealerName') },
      ];

      const totalDefined = COLS.reduce((s, c) => s + c.w, 0);
      const scale = USABLE_W / totalDefined;
      COLS.forEach(c => c.w = c.w * scale);

      const tableBody = rows.map(r => COLS.map(c => c.get(r)));

      const HEADER_ROW_H = 26;
      const DATA_ROW_H   = 9;
      const FONT_SIZE    = 8.5;
      const HDR_FONT_SZ  = 9;

      const COLOR_GOLD   = [232, 162, 23];
      const COLOR_NAVY   = [13,  27,  42];
      const COLOR_STRIPE = [240, 244, 250];
      const COLOR_BORDER = [180, 190, 205];

      doc.autoTable({
        head: [COLS.map(c => c.hdr)],
        body: tableBody,
        startY: MARGIN + 8,
        margin: { top: MARGIN + 8, right: MARGIN, bottom: MARGIN + 4, left: MARGIN },
        tableWidth: USABLE_W,
        columnStyles: (() => {
          const cs = {};
          COLS.forEach((c, i) => {
            cs[i] = { cellWidth: c.w, fontSize: FONT_SIZE, cellPadding: { top: 0.8, right: 0.8, bottom: 0.8, left: 0.8 } };
          });
          return cs;
        })(),
        headStyles: {
          fillColor: COLOR_GOLD, textColor: COLOR_NAVY, fontStyle: 'bold', fontSize: HDR_FONT_SZ,
          cellPadding: { top: 1, right: 1, bottom: 1, left: 1 },
          minCellHeight: HEADER_ROW_H, valign: 'bottom', overflow: 'hidden',
        },
        bodyStyles: {
          fontSize: FONT_SIZE, textColor: COLOR_NAVY,
          cellPadding: { top: 0.8, right: 0.8, bottom: 0.8, left: 0.8 },
          minCellHeight: DATA_ROW_H, valign: 'top', lineColor: COLOR_BORDER, lineWidth: 0.15,
        },
        alternateRowStyles: { fillColor: COLOR_STRIPE },
        styles: { font: 'helvetica', overflow: 'linebreak', lineColor: COLOR_BORDER, lineWidth: 0.15 },
        didDrawCell(data) {
          if (data.section !== 'head') return;
          const { x, y, width, height } = data.cell;
          const txt = data.cell.raw;
          doc.setFillColor(...COLOR_GOLD);
          doc.rect(x, y, width, height, 'F');
          doc.setDrawColor(...COLOR_BORDER);
          doc.setLineWidth(0.15);
          doc.rect(x, y, width, height, 'S');
          doc.setFont('helvetica', 'bold');
          doc.setFontSize(HDR_FONT_SZ);
          doc.setTextColor(...COLOR_NAVY);
          const cx = x + width / 2;
          const cy = y + height - 2;
          doc.saveGraphicsState();
          doc.text(txt, cx, cy, { angle: 90, align: 'left', baseline: 'middle' });
          doc.restoreGraphicsState();
        },
        didDrawPage(data) {
          const pageCount = doc.internal.getNumberOfPages();
          const pageNum   = data.pageNumber;
          
          doc.setFillColor(...COLOR_NAVY);
          doc.rect(0, 0, PAGE_W, MARGIN + 6, 'F');
          
          doc.setFont('helvetica', 'bold');
          doc.setFontSize(12);
          doc.setTextColor(255, 255, 255);
          doc.text(`Agency / DMA Name : ${agencyName}`, MARGIN + 1, MARGIN + 3.5);

          doc.setFillColor(...COLOR_NAVY);
          doc.rect(0, PAGE_H - MARGIN - 3, PAGE_W, MARGIN + 3, 'F');
          
          doc.setFont('helvetica', 'normal');
          doc.setFontSize(7);
          doc.setTextColor(200, 200, 200);
          doc.text(`Total Records: ${totalRows.toLocaleString()}  •  Generated: ${new Date().toLocaleDateString('en-IN')}  •  created by pratik verma`, MARGIN, PAGE_H - MARGIN - 0.8);
          doc.text(`Page ${pageNum} of ${pageCount}`, PAGE_W - MARGIN, PAGE_H - MARGIN - 0.8, { align: 'right' });
        },
        showHead: 'everyPage',
        rowPageBreak: 'avoid',
      });

      const filename = `${agencyName}_Recovery_Sheet_${timestamp}.pdf`;
      if (useZip) {
        zip.file(filename, doc.output('blob'));
      } else {
        doc.save(filename);
      }
    }

    if (useZip) {
      setProgress(95, 'Saving ZIP file…');
      await sleep(30);
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      saveAs(zipBlob, `Recovery_Sheets_${timestamp}.zip`);
    }

    setProgress(100, 'Done! File(s) downloaded.');
    await sleep(800);

  } catch (err) {
    console.error('PDF generation error:', err);
    alert('PDF generation failed:\n' + err.message);
  } finally {
    genBtnText.classList.remove('hidden');
    spinner.classList.add('hidden');
    generateBtn.disabled = false;
    setTimeout(() => {
      progressWrap.classList.add('hidden');
      setProgress(0, '');
    }, 2000);
  }
}

/* ═══════════════════════════════════════════════
   10. SEARCH AND SINGLE VIEW
   ═══════════════════════════════════════════════ */

let currentFoundRow = null;

searchBtn.addEventListener('click', () => {
  const query = searchInput.value.trim().toLowerCase();
  if (!query) return;

  searchError.classList.add('hidden');
  searchResult.classList.add('hidden');
  currentFoundRow = null;

  const match = parsedRows.find(r => {
    const ln = get(r, 'loanNo').toLowerCase();
    return ln === query;
  });

  if (match) {
    currentFoundRow = match;
    const fieldsToShow = ['loanNo', 'customerName', 'agencyName', 'pos', 'openingOD', 'totalOD', 'piDues', 'emi', 'bkt', 'mobile', 'address', 'make', 'regNo', 'dealerName'];
    resultGrid.innerHTML = '';
    for (const k of fieldsToShow) {
      const def = FIELD_DEFS.find(d => d.key === k);
      const val = get(match, k);
      if (val) {
        resultGrid.innerHTML += `<div class="r-item"><div class="r-lbl">${def.label}</div><div class="r-val">${val}</div></div>`;
      }
    }
    searchResult.classList.remove('hidden');
  } else {
    searchError.classList.remove('hidden');
  }
});

downloadSingleBtn.addEventListener('click', () => {
  if (currentFoundRow) {
    generateSinglePDF(currentFoundRow);
  }
});

function generateSinglePDF(row) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const MARGIN = 3;
  const PAGE_W = 297;
  const PAGE_H = 210;
  const USABLE_W = PAGE_W - MARGIN * 2;
  const agencyName = get(row, 'agencyName') || 'Single_Record';

  const COLS = [
    { w: 22, hdr: 'Loan No',        get: r => get(r, 'loanNo') },
    { w: 25, hdr: 'Customer Name',  get: r => get(r, 'customerName') },
    { w: 6,  hdr: 'BKT',            get: r => get(r, 'bkt') },
    { w: 11, hdr: 'POS',            get: r => get(r, 'pos') },
    {
      w: 13,
      hdr: 'OD Details',
      get: r => {
        const opn = get(r, 'openingOD');
        const tot = get(r, 'totalOD');
        const parts = [];
        if (opn) parts.push('Opn: ' + opn);
        if (tot) parts.push('Tot: ' + tot);
        return parts.join('\n');
      }
    },
    { w: 8,  hdr: 'EMI',            get: r => get(r, 'emi') },
    { w: 10, hdr: 'PI Dues',        get: r => get(r, 'piDues') },
    { w: 44, hdr: 'Permanent Address', get: r => get(r, 'address') },
    {
      w: 44,
      hdr: 'Contact / Reference',
      get: r => {
        const mob   = get(r, 'mobile');
        const ref1n = get(r, 'ref1Name');
        const ref1m = get(r, 'ref1Mobile');
        const ref2n = get(r, 'ref2Name');
        const ref2m = get(r, 'ref2Mobile');
        const parts = [];
        if (mob)   parts.push('Mob: ' + mob);
        if (ref1n) parts.push('Ref1: ' + ref1n);
        if (ref1m) parts.push('RefMob: ' + ref1m);
        if (ref2n) parts.push('Ref2: ' + ref2n);
        if (ref2m) parts.push('Ref2Mob: ' + ref2m);
        return parts.join('\n');
      }
    },
    { w: 13, hdr: 'Make', get: r => get(r, 'make') },
    { w: 16, hdr: 'REG NO', get: r => get(r, 'regNo') },
    { w: 16, hdr: 'DLR Name', get: r => get(r, 'dealerName') },
  ];

  const totalDefined = COLS.reduce((s, c) => s + c.w, 0);
  const scale = USABLE_W / totalDefined;
  COLS.forEach(c => c.w = c.w * scale);

  const tableBody = [COLS.map(c => c.get(row))];
  const HEADER_ROW_H = 26;
  const DATA_ROW_H   = 9;
  const FONT_SIZE    = 8.5;
  const HDR_FONT_SZ  = 9;

  const COLOR_GOLD   = [232, 162, 23];
  const COLOR_NAVY   = [13,  27,  42];
  const COLOR_STRIPE = [240, 244, 250];
  const COLOR_BORDER = [180, 190, 205];

  doc.autoTable({
    head: [COLS.map(c => c.hdr)],
    body: tableBody,
    startY: MARGIN + 8,
    margin: { top: MARGIN + 8, right: MARGIN, bottom: MARGIN + 4, left: MARGIN },
    tableWidth: USABLE_W,
    columnStyles: (() => {
      const cs = {};
      COLS.forEach((c, i) => {
        cs[i] = { cellWidth: c.w, fontSize: FONT_SIZE, cellPadding: { top: 0.8, right: 0.8, bottom: 0.8, left: 0.8 } };
      });
      return cs;
    })(),
    headStyles: {
      fillColor: COLOR_GOLD, textColor: COLOR_NAVY, fontStyle: 'bold', fontSize: HDR_FONT_SZ,
      cellPadding: { top: 1, right: 1, bottom: 1, left: 1 },
      minCellHeight: HEADER_ROW_H, valign: 'bottom', overflow: 'hidden',
    },
    bodyStyles: {
      fontSize: FONT_SIZE, textColor: COLOR_NAVY,
      cellPadding: { top: 0.8, right: 0.8, bottom: 0.8, left: 0.8 },
      minCellHeight: DATA_ROW_H, valign: 'top', lineColor: COLOR_BORDER, lineWidth: 0.15,
    },
    alternateRowStyles: { fillColor: COLOR_STRIPE },
    styles: { font: 'helvetica', overflow: 'linebreak', lineColor: COLOR_BORDER, lineWidth: 0.15 },
    didDrawCell(data) {
      if (data.section !== 'head') return;
      const { x, y, width, height } = data.cell;
      const txt = data.cell.raw;
      doc.setFillColor(...COLOR_GOLD);
      doc.rect(x, y, width, height, 'F');
      doc.setDrawColor(...COLOR_BORDER);
      doc.setLineWidth(0.15);
      doc.rect(x, y, width, height, 'S');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(HDR_FONT_SZ);
      doc.setTextColor(...COLOR_NAVY);
      const cx = x + width / 2;
      const cy = y + height - 2;
      doc.saveGraphicsState();
      doc.text(txt, cx, cy, { angle: 90, align: 'left', baseline: 'middle' });
      doc.restoreGraphicsState();
    },
    didDrawPage(data) {
      doc.setFillColor(...COLOR_NAVY);
      doc.rect(0, 0, PAGE_W, MARGIN + 6, 'F');
      
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(12);
      doc.setTextColor(255, 255, 255);
      doc.text(`Agency / DMA Name : ${agencyName}`, MARGIN + 1, MARGIN + 3.5);

      doc.setFillColor(...COLOR_NAVY);
      doc.rect(0, PAGE_H - MARGIN - 3, PAGE_W, MARGIN + 3, 'F');
      
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7);
      doc.setTextColor(200, 200, 200);
      doc.text(`Generated: ${new Date().toLocaleDateString('en-IN')}  •  created by pratik verma`, MARGIN, PAGE_H - MARGIN - 0.8);
    },
  });

  const cName = get(row, 'customerName').replace(/[^a-z0-9]/gi, '_') || 'Customer';
  const lNo = get(row, 'loanNo').replace(/[^a-z0-9]/gi, '_') || 'Loan';
  doc.save(`${cName}_${lNo}_Recovery.pdf`);
}

/* ═══════════════════════════════════════════════
   11. HELPERS
   ═══════════════════════════════════════════════ */

function setProgress(pct, label) {
  progressFill.style.width  = pct + '%';
  progressLabel.textContent = label;
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function resetAll() {
  parsedRows = [];
  mappedKeys = {};
  fileInput.value = '';
  fileInfo.classList.add('hidden');
  mappingSection.classList.add('hidden');
  searchSection.classList.add('hidden');
  generateSection.classList.add('hidden');
  mappingGrid.innerHTML = '';
  agencySelect.innerHTML = '<option value="ALL">-- Generate All (Download as ZIP) --</option>';
  
  searchInput.value = '';
  searchError.classList.add('hidden');
  searchResult.classList.add('hidden');
  resultGrid.innerHTML = '';
  currentFoundRow = null;

  statRecords.textContent = '0';
  statPages.textContent   = '~0';
}

/* ═══════════════════════════════════════════════
   12. PWA & WEB SHARE TARGET
   ═══════════════════════════════════════════════ */

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js').then(reg => {
      console.log('SW registered:', reg);
    }).catch(err => console.log('SW registration failed:', err));
  });
}

// Check if app was launched via File Handling API (Open With dialog)
if ('launchQueue' in window) {
  window.launchQueue.setConsumer(async (launchParams) => {
    if (launchParams.files && launchParams.files.length > 0) {
      for (const fileHandle of launchParams.files) {
        const file = await fileHandle.getFile();
        handleFile(file);
        break; // Process only the first file
      }
    }
  });
}

// Check if app was launched via Share Target
window.addEventListener('DOMContentLoaded', async () => {
  const urlParams = new URLSearchParams(window.location.search);
  if (urlParams.has('shared')) {
    try {
      const cache = await caches.open('shared-file');
      const response = await cache.match('/shared-file');
      if (response) {
        const blob = await response.blob();
        const filename = response.headers.get('X-Filename') || 'shared_file.xlsx';
        const file = new File([blob], filename, { type: blob.type });
        
        // Remove the query param from URL
        window.history.replaceState({}, document.title, window.location.pathname);
        
        // Process the shared file
        handleFile(file);
        
        // Clear cache
        await cache.delete('/shared-file');
      }
    } catch (err) {
      console.error('Failed to load shared file:', err);
    }
  }
});
