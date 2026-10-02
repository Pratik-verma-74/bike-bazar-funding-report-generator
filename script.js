/**
 * Loan Recovery PDF Converter — script.js
 * ================================================
 * Generates custom, tailored recovery sheet PDFs
 * based strictly on the uploaded Excel file's format.
 *
 * Supported Modes:
 *   1. Format 1: Bike Bazar Standard Format (22 columns)
 *   2. Format 2: DMA / Agency Recovery Sheet (30 columns)
 *   3. Dynamic:  Any other custom Excel file (all columns rendered)
 * ================================================
 */

/* ═══════════════════════════════════════════════
   1. STATE
   ═══════════════════════════════════════════════ */
let parsedRows = [];        // Clean data rows
let currentHeaders = [];    // Raw headers from uploaded Excel
let detectedFormat = null;  // 'FORMAT_1' | 'FORMAT_2' | 'DYNAMIC'
let currentFoundRow = null; // Currently matched search row

/* ═══════════════════════════════════════════════
   2. DOM REFERENCES
   ═══════════════════════════════════════════════ */
const dropZone          = document.getElementById('dropZone');
const fileInput         = document.getElementById('fileInput');
const fileInfo          = document.getElementById('fileInfo');
const fileNameEl        = document.getElementById('fileName');
const recordCountEl     = document.getElementById('recordCount');
const clearFileBtn      = document.getElementById('clearFile');
const mappingSection    = document.getElementById('mapping-section');
const formatBadge       = document.getElementById('formatBadge');
const mappingGrid       = document.getElementById('mappingGrid');
const searchSection     = document.getElementById('search-section');
const searchInput       = document.getElementById('searchInput');
const searchBtn         = document.getElementById('searchBtn');
const searchError       = document.getElementById('searchError');
const searchResult      = document.getElementById('searchResult');
const resultGrid        = document.getElementById('resultGrid');
const downloadSingleBtn = document.getElementById('downloadSingleBtn');
const generateSection   = document.getElementById('generate-section');
const agencySelect      = document.getElementById('agencySelect');
const statRecords       = document.getElementById('statRecords');
const statPages         = document.getElementById('statPages');
const statCols          = document.getElementById('statCols');
const generateBtn       = document.getElementById('generateBtn');
const genBtnText        = document.getElementById('genBtnText');
const spinner           = document.getElementById('spinner');
const progressWrap      = document.getElementById('progressWrap');
const progressFill      = document.getElementById('progressFill');
const progressLabel     = document.getElementById('progressLabel');

/* ═══════════════════════════════════════════════
   3. DRAG & DROP / FILE INPUT
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
  if (e.target.tagName !== 'LABEL') fileInput.click();
});
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
});
clearFileBtn.addEventListener('click', resetAll);

/* ═══════════════════════════════════════════════
   4. FILE HANDLER
   ═══════════════════════════════════════════════ */

function handleFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls', 'csv'].includes(ext)) {
    alert('Please upload an Excel file (.xlsx, .xls) or CSV.');
    return;
  }

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb   = XLSX.read(data, { type: 'array', cellDates: true });
      const ws   = wb.Sheets[wb.SheetNames[0]];

      // Raw 2D grid to reliably identify the header row
      const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (!rawRows || !rawRows.length) {
        alert('The file appears to be empty.');
        return;
      }

      // Auto-detect header row index (scanning first 5 rows)
      let headerRowIndex = 0;
      for (let i = 0; i < Math.min(5, rawRows.length); i++) {
        const rowStr = rawRows[i].map(c => String(c).toLowerCase()).join(' ');
        if (rowStr.includes('loan') || rowStr.includes('customer') || rowStr.includes('emi') || rowStr.includes('bkt') || rowStr.includes('pos') || rowStr.includes('agency')) {
          headerRowIndex = i;
          break;
        }
      }

      const headers = rawRows[headerRowIndex].map(h => String(h || '').trim());
      const cleanHeaders = [];
      const headerCounts = {};
      headers.forEach((h, i) => {
        let name = h || `Column_${i + 1}`;
        if (headerCounts[name]) {
          headerCounts[name]++;
          name = `${name}_${headerCounts[name]}`;
        } else {
          headerCounts[name] = 1;
        }
        cleanHeaders.push(name);
      });

      const rows = [];
      for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
        const rawRow = rawRows[i];
        if (!rawRow || !rawRow.length) continue;

        const rowObj = {};
        let hasAnyVal = false;
        cleanHeaders.forEach((hdr, colIdx) => {
          let val = rawRow[colIdx];
          if (val instanceof Date) {
            val = val.toLocaleDateString('en-IN');
          } else if (val !== undefined && val !== null) {
            val = String(val).trim();
          } else {
            val = '';
          }
          if (val) hasAnyVal = true;
          rowObj[hdr] = val;
        });

        // Filter out completely blank lines
        if (hasAnyVal) {
          rows.push(rowObj);
        }
      }

      if (!rows.length) {
        alert('No data rows found in the uploaded file.');
        return;
      }

      parsedRows = rows;
      currentHeaders = cleanHeaders;
      detectedFormat = identifyFormat(cleanHeaders);

      // Show file info
      fileNameEl.textContent = file.name;
      recordCountEl.textContent = `${rows.length.toLocaleString()} records`;
      fileInfo.classList.remove('hidden');

      // Render format preview
      renderFormatPreview(detectedFormat, cleanHeaders);

      // Populate Agency Dropdown & Stats
      refreshAgenciesAndStats();
      agencySelect.onchange = updateStats;

      // Reveal UI sections
      mappingSection.classList.remove('hidden');
      searchSection.classList.remove('hidden');
      generateSection.classList.remove('hidden');

    } catch (err) {
      console.error(err);
      alert('Failed to read the file. Please make sure it is a valid Excel/CSV file.');
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ═══════════════════════════════════════════════
   5. FORMAT DETECTION ENGINE
   ═══════════════════════════════════════════════ */

function normalize(s) {
  return String(s || '').toLowerCase().replace(/[\s_\-\.\/]+/g, ' ').trim();
}

function identifyFormat(headers) {
  const norm = headers.map(h => normalize(h));
  const has = keyword => norm.some(h => h.includes(keyword));

  // Format 1 identifiers: 'opening bkt', 'total pos', 'vehicle no', 'alternate mobile'
  const isF1 = has('total pos') || has('opening bkt') || has('alternate mobile');

  // Format 2 identifiers: 'ref 1', 'ref1', 'alt/data mob', 'dma/dma name', 'agency/dma', 'opnig od', 'totl od'
  const isF2 = has('ref 1') || has('ref1') || has('alt/data') || has('agency/dma') || has('opnig od') || has('totl od');

  if (isF1 && !has('ref 1') && !has('ref1')) return 'FORMAT_1';
  if (isF2) return 'FORMAT_2';
  return 'DYNAMIC';
}

/**
 * Robust case-insensitive lookup helper from row object
 */
function val(row, ...keys) {
  if (!row) return '';
  for (const k of keys) {
    if (row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '') {
      return String(row[k]).trim();
    }
    const normK = normalize(k);
    for (const [col, v] of Object.entries(row)) {
      if (normalize(col) === normK && v !== undefined && v !== null && String(v).trim() !== '') {
        return String(v).trim();
      }
    }
  }
  return '';
}

/**
 * Cleans phone number strings (strips Excel .0 decimals, scientific notation)
 */
function cleanPhone(raw) {
  if (!raw && raw !== 0) return '';
  let s = String(raw).trim();
  if (/^\d+\.?\d*e\+\d+$/i.test(s)) {
    try {
      s = BigInt(Math.round(Number(s))).toString();
    } catch (e) {}
  }
  if (/^\d+\.0+$/.test(s)) {
    s = s.split('.')[0];
  }
  return s;
}

/**
 * Cleanly formats primary and all alternate mobile numbers together
 */
function formatContactNumbers(...phoneSources) {
  const parts = [];
  const seen = new Set();

  phoneSources.forEach((source, sourceIdx) => {
    if (!source) return;
    const numbers = String(source).split(/[\/;,]+/).map(s => cleanPhone(s)).filter(Boolean);
    numbers.forEach((num, subIdx) => {
      const pureDigits = num.replace(/\D/g, '');
      if (pureDigits.length >= 5 && !seen.has(pureDigits)) {
        seen.add(pureDigits);
        let prefix = 'Mob';
        if (sourceIdx === 0 && subIdx === 0) prefix = 'Mob';
        else if (sourceIdx === 0 && subIdx > 0) prefix = `Mob${subIdx + 1}`;
        else if (sourceIdx === 1) prefix = subIdx === 0 ? 'Alt' : `Alt${subIdx + 1}`;
        else prefix = `Alt${sourceIdx + subIdx}`;

        parts.push(`${prefix}: ${num}`);
      } else if (pureDigits.length < 5 && num && !seen.has(num)) {
        seen.add(num);
        parts.push(`Mob: ${num}`);
      }
    });
  });

  return parts.length ? parts.join('\n') : '';
}

/* ═══════════════════════════════════════════════
   6. FORMAT-SPECIFIC TABLE CONFIGURATION
   Tailored to produce the exact PDF matching
   whichever Excel is uploaded.
   ═══════════════════════════════════════════════ */

function getFormatConfig(format, usableWidth) {
  if (format === 'FORMAT_1') {
    // ── FORMAT 1: Bike Bazar Standard (22 Excel columns accommodated in 12 PDF columns) ──
    const COLS = [
      { w: 22, hdr: 'Loan Number', get: r => val(r, 'Loan Number', 'Loan No', 'Loan_No') },
      { w: 24, hdr: 'Customer Name', get: r => val(r, 'Customer Name', 'Customer') },
      { w: 7,  hdr: 'BKT', get: r => val(r, 'Opening Bkt - 7+', 'Opening Bkt', 'BKT') },
      { w: 12, hdr: 'Total POS', get: r => val(r, 'Total POS', 'POS') },
      {
        w: 14,
        hdr: 'OD Details',
        get: r => {
          const opn = val(r, 'Opening OD');
          const tot = val(r, 'Total OD');
          const parts = [];
          if (opn) parts.push('Opn: ' + opn);
          if (tot) parts.push('Tot: ' + tot);
          return parts.join('\n');
        }
      },
      {
        w: 12,
        hdr: 'PI Details',
        get: r => {
          const opn = val(r, 'Opening PI');
          const tot = val(r, 'Total PI');
          const parts = [];
          if (opn) parts.push('Opn: ' + opn);
          if (tot) parts.push('Tot: ' + tot);
          return parts.join('\n');
        }
      },
      { w: 9, hdr: 'EMI', get: r => val(r, 'EMI') },
      {
        w: 42,
        hdr: 'Current Address',
        get: r => {
          const addr = val(r, 'Current Address', 'Address');
          const pin = val(r, 'Pin Code', 'Pincode');
          if (pin && !addr.includes(pin)) return addr ? `${addr}\nPIN: ${pin}` : `PIN: ${pin}`;
          return addr;
        }
      },
      {
        w: 36,
        hdr: 'Contact & Alt Mob',
        get: r => {
          const mob = val(r, 'Contact Number', 'Contact No', 'Contact No.', 'Customer Mobile', 'Mobile No', 'Mobile No.', 'Mobile Number', 'Mobile', 'Contact', 'Phone Number', 'Phone');
          const alt = val(r, 'Alternate Mobile no', 'Alternate Mobile no.', 'Alternate Mobile number', 'Alternate Mobile', 'Alternate Mob', 'Alternate Mob no', 'Alt Mobile no', 'Alt Mobile no.', 'Alt Mobile', 'Alt Mob', 'Alternate Contact', 'Alt Contact', 'Alternate Number', 'Alt Number', 'Alt No', 'ALT/DATA MOB-1', 'ALT MOB-1');
          const alt2 = val(r, 'Alternate Mobile no 2', 'Alternate Mobile 2', 'Alt Mobile 2', 'ALT/DATA MOB-2', 'ALT MOB-2');
          return formatContactNumbers(mob, alt, alt2);
        }
      },
      {
        w: 16,
        hdr: 'Make / Year',
        get: r => {
          const mk = val(r, 'Vehicle', 'Make');
          const yr = val(r, 'Mfg Yr', 'Year');
          return yr ? `${mk}\nYr: ${yr}` : mk;
        }
      },
      {
        w: 24,
        hdr: 'Vehicle / Assets',
        get: r => {
          const reg = val(r, 'Vehicle No', 'Reg No');
          const ch  = val(r, 'Chassis No.', 'Chassis No');
          const eng = val(r, 'Engine No.', 'Engine No');
          const parts = [];
          if (reg) parts.push(reg);
          if (ch)  parts.push('Ch: ' + ch);
          if (eng) parts.push('Eng: ' + eng);
          return parts.join('\n');
        }
      },
      {
        w: 24,
        hdr: 'Dealer / TCM',
        get: r => {
          const dlr  = val(r, 'Dealer Name', 'Dealer');
          const cost = val(r, 'Dealer Cost', 'Dealer Code');
          const tcm  = val(r, 'TCM NAME', 'TCM');
          const parts = [];
          if (dlr)  parts.push(dlr);
          if (cost) parts.push('Cost: ' + cost);
          if (tcm)  parts.push('TCM: ' + tcm);
          return parts.join('\n');
        }
      },
    ];

    const totalDefined = COLS.reduce((s, c) => s + c.w, 0);
    const scale = usableWidth / totalDefined;
    COLS.forEach(c => c.w = c.w * scale);

    return {
      name: 'Format 1 (Bike Bazar Standard)',
      agencyHeader: 'Agency Name',
      getAgency: r => val(r, 'AGENCY NAME', 'Agency Name', 'Agency'),
      cols: COLS,
      fieldCount: 22
    };

  } else if (format === 'FORMAT_2') {
    // ── FORMAT 2: DMA / Agency Sheet (30 Excel columns accommodated in 13 PDF columns) ──
    const COLS = [
      { w: 20, hdr: 'Loan No', get: r => val(r, 'LOAN NO', 'Loan No') },
      { w: 22, hdr: 'Customer Name', get: r => val(r, 'CUSTOMER', 'Customer Name') },
      { w: 6,  hdr: 'BKT', get: r => val(r, 'BKT') },
      { w: 11, hdr: 'POS', get: r => val(r, 'POS') },
      {
        w: 13,
        hdr: 'OD Details',
        get: r => {
          const opn = val(r, 'OPNIG OD', 'Opening OD');
          const tot = val(r, 'TOTL OD', 'Total OD');
          const parts = [];
          if (opn) parts.push('Opn: ' + opn);
          if (tot) parts.push('Tot: ' + tot);
          return parts.join('\n');
        }
      },
      { w: 9, hdr: 'PI Dues', get: r => val(r, 'PI DUES', 'PI Due') },
      { w: 8, hdr: 'EMI', get: r => val(r, 'EMI AMOUNT', 'EMI') },
      {
        w: 38,
        hdr: 'Permanent Address',
        get: r => {
          const addr = val(r, 'PERMNT ADD', 'Permanent Address', 'Address');
          const pin = val(r, 'PIN CODE', 'Pin Code');
          if (pin && !addr.includes(pin)) return addr ? `${addr}\nPIN: ${pin}` : `PIN: ${pin}`;
          return addr;
        }
      },
      {
        w: 34,
        hdr: 'Contact & Alt Mob',
        get: r => {
          const mob  = val(r, 'CONTACT', 'Contact Number', 'Contact No', 'Customer Mobile', 'Mobile No', 'Mobile', 'Phone');
          const alt1 = val(r, 'ALT/DATA MOB-1', 'ALT MOB-1', 'ALT/DATA MOB 1', 'Alternate Mobile no', 'Alternate Mobile', 'Alt Mobile');
          const alt2 = val(r, 'ALT/DATA MOB-2', 'ALT MOB-2', 'ALT/DATA MOB 2', 'Alternate Mobile 2', 'Alt Mobile 2');
          return formatContactNumbers(mob, alt1, alt2);
        }
      },
      {
        w: 36,
        hdr: 'References',
        get: r => {
          const r1n = val(r, 'DATA/REF-1', 'REF-1', 'Reference 1');
          const r1m = cleanPhone(val(r, 'Reference-1 mobile no', 'REF-1 MOB', 'Reference 1 Mobile'));
          const r2n = val(r, 'DATA/REF-2', 'REF-2', 'Reference 2');
          const r2m = cleanPhone(val(r, 'Reference-2 mobile no', 'REF-2 MOB', 'Reference 2 Mobile'));
          const parts = [];
          if (r1n || r1m) parts.push(`Ref1: ${r1n}${r1m ? ' (' + r1m + ')' : ''}`);
          if (r2n || r2m) parts.push(`Ref2: ${r2n}${r2m ? ' (' + r2m + ')' : ''}`);
          return parts.join('\n');
        }
      },
      {
        w: 24,
        hdr: 'Vehicle Details',
        get: r => {
          const mk  = val(r, 'MAKE', 'Make');
          const reg = val(r, 'REG NO', 'Reg No');
          const ch  = val(r, 'CHASSIS NO', 'Chassis No');
          const parts = [];
          if (mk)  parts.push('Make: ' + mk);
          if (reg) parts.push('Reg: ' + reg);
          if (ch)  parts.push('Ch: ' + ch);
          return parts.join('\n');
        }
      },
      {
        w: 20,
        hdr: 'Dealer / TCM',
        get: r => {
          const dlr = val(r, 'DLR NAME', 'Dealer Name');
          const tcm = val(r, 'TCM', 'TCM Name');
          const parts = [];
          if (dlr) parts.push('DLR: ' + dlr);
          if (tcm) parts.push('TCM: ' + tcm);
          return parts.join('\n');
        }
      },
      {
        w: 18,
        hdr: 'Product / Status',
        get: r => {
          const prod = val(r, 'PRODUCT');
          const repo = val(r, 'Repossess Repo (Y/N)?', 'Repo');
          const sbot = val(r, 'SBOT AMOUNT', 'SBOT');
          const parts = [];
          if (prod) parts.push('Prod: ' + prod);
          if (repo) parts.push('Repo: ' + repo);
          if (sbot && sbot !== '0') parts.push('SBOT: ' + sbot);
          return parts.join('\n');
        }
      }
    ];

    const totalDefined = COLS.reduce((s, c) => s + c.w, 0);
    const scale = usableWidth / totalDefined;
    COLS.forEach(c => c.w = c.w * scale);

    return {
      name: 'Format 2 (DMA / Agency Sheet)',
      agencyHeader: 'Agency / DMA Name',
      getAgency: r => val(r, 'AGENCY/DMA NAME', 'DMA/DMA NAME', 'AGENCY NAME', 'Agency Code'),
      cols: COLS,
      fieldCount: 30
    };

  } else {
    // ── DYNAMIC: Any other custom Excel file ──
    const cols = currentHeaders.map(hdr => ({
      w: Math.max(16, Math.min(45, hdr.length * 2.5 + 8)),
      hdr: hdr,
      get: r => String(r[hdr] || '').trim()
    }));

    const totalDefined = cols.reduce((s, c) => s + c.w, 0);
    const scale = usableWidth / totalDefined;
    cols.forEach(c => c.w = c.w * scale);

    return {
      name: 'Custom Excel Format',
      agencyHeader: 'Agency / Department',
      getAgency: r => val(r, 'AGENCY NAME', 'AGENCY/DMA NAME', 'Agency', 'DMA', 'Branch') || 'All Records',
      cols: cols,
      fieldCount: currentHeaders.length
    };
  }
}

/* ═══════════════════════════════════════════════
   7. STEP 2: FORMAT & COLUMNS PREVIEW IN UI
   ═══════════════════════════════════════════════ */

function renderFormatPreview(format, headers) {
  mappingGrid.innerHTML = '';

  if (format === 'FORMAT_1') {
    formatBadge.className = 'format-badge fmt-1';
    formatBadge.innerHTML = `
      <div style="font-size: 1.1rem">✨</div>
      <div>
        <strong>Format 1 Detected: Bike Bazar Standard Recovery Sheet</strong><br>
        <span style="font-size: 0.76rem; opacity: 0.9">All 22 columns from your Excel are mapped directly to Format 1 PDF layout. 0 columns missing.</span>
      </div>
    `;
  } else if (format === 'FORMAT_2') {
    formatBadge.className = 'format-badge fmt-2';
    formatBadge.innerHTML = `
      <div style="font-size: 1.1rem">✨</div>
      <div>
        <strong>Format 2 Detected: DMA / Agency Recovery Sheet</strong><br>
        <span style="font-size: 0.76rem; opacity: 0.9">All 30 columns from your Excel (including References, Products, Alt Mobiles) are mapped directly to Format 2 PDF layout. 0 columns missing.</span>
      </div>
    `;
  } else {
    formatBadge.className = 'format-badge fmt-custom';
    formatBadge.innerHTML = `
      <div style="font-size: 1.1rem">✨</div>
      <div>
        <strong>Custom Excel File Detected</strong><br>
        <span style="font-size: 0.76rem; opacity: 0.9">All ${headers.length} columns from your Excel will be dynamically rendered in the PDF.</span>
      </div>
    `;
  }

  // Display all detected columns as active verified chips
  headers.forEach((h, index) => {
    const div = document.createElement('div');
    div.className = 'map-item found';
    div.innerHTML = `
      <div class="map-header">
        <span class="map-field">${escapeHtml(h)}</span>
        <span class="req-tag" style="background: rgba(74, 222, 128, 0.15); color: #4ade80;">Active</span>
      </div>
      <div class="map-source" style="color: var(--green);">✓ Included in PDF output</div>
    `;
    mappingGrid.appendChild(div);
  });

  statCols.textContent = headers.length;
}

/* ═══════════════════════════════════════════════
   8. PDF GENERATION (BATCH / ZIP)
   ═══════════════════════════════════════════════ */

generateBtn.addEventListener('click', () => {
  if (!parsedRows.length) return;
  generatePDF();
});

async function generatePDF() {
  genBtnText.classList.add('hidden');
  spinner.classList.remove('hidden');
  progressWrap.classList.remove('hidden');
  generateBtn.disabled = true;
  setProgress(0, 'Preparing data…');

  await sleep(30);

  try {
    const { jsPDF } = window.jspdf;
    const selectedAgency = agencySelect.value;
    const MARGIN   = 3;
    const PAGE_W   = 297;
    const PAGE_H   = 210;
    const USABLE_W = PAGE_W - MARGIN * 2;

    const config = getFormatConfig(detectedFormat, USABLE_W);

    // Group by Agency
    const groups = {};
    for (const r of parsedRows) {
      let rawAgency = config.getAgency(r) || 'Unassigned Agency';
      if (selectedAgency !== 'ALL' && rawAgency !== selectedAgency) continue;

      let safeAgency = rawAgency.replace(/[^a-z0-9_\-\s]/gi, '_').trim() || 'Agency';
      if (!groups[safeAgency]) groups[safeAgency] = [];
      groups[safeAgency].push(r);
    }

    const agencyNames = Object.keys(groups);
    if (agencyNames.length === 0) {
      alert('No records found for the selected agency.');
      return;
    }

    const useZip = selectedAgency === 'ALL';
    let zip;
    if (useZip) zip = new JSZip();

    const timestamp = new Date().toISOString().slice(0, 10);

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

      const COLS = config.cols;
      const tableBody = rows.map(r => COLS.map(c => c.get(r)));

      const HEADER_ROW_H = 26;
      const DATA_ROW_H   = 9;
      const FONT_SIZE    = 8.2;
      const HDR_FONT_SZ  = 8.8;

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
          doc.text(`${config.agencyHeader} : ${agencyName}`, MARGIN + 1, MARGIN + 3.5);

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
   9. SEARCH AND SINGLE VIEW
   ═══════════════════════════════════════════════ */

searchBtn.addEventListener('click', doSearch);
searchInput.addEventListener('keydown', e => {
  if (e.key === 'Enter') doSearch();
});

function doSearch() {
  const query = searchInput.value.trim().toLowerCase();
  if (!query) return;

  searchError.classList.add('hidden');
  searchResult.classList.add('hidden');
  currentFoundRow = null;

  // Search across all fields in the row
  const match = parsedRows.find(r => {
    return Object.values(r).some(val => String(val).toLowerCase().includes(query));
  });

  if (match) {
    currentFoundRow = match;
    resultGrid.innerHTML = '';
    // Show all non-empty key-value pairs from the uploaded Excel row
    for (const [k, v] of Object.entries(match)) {
      if (v) {
        const isPhone = /mobile|contact|phone|mob/i.test(k);
        const displayVal = isPhone ? cleanPhone(v) : v;
        resultGrid.innerHTML += `
          <div class="r-item">
            <div class="r-lbl">${escapeHtml(k)}</div>
            <div class="r-val">${escapeHtml(displayVal)}</div>
          </div>`;
      }
    }
    searchResult.classList.remove('hidden');
  } else {
    searchError.textContent = `No record found matching "${searchInput.value.trim()}".`;
    searchError.classList.remove('hidden');
  }
}

downloadSingleBtn.addEventListener('click', () => {
  if (currentFoundRow) {
    generateSinglePDF(currentFoundRow);
  }
});

function generateSinglePDF(row) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const MARGIN   = 3;
  const PAGE_W   = 297;
  const PAGE_H   = 210;
  const USABLE_W = PAGE_W - MARGIN * 2;

  const config = getFormatConfig(detectedFormat, USABLE_W);
  const agencyName = config.getAgency(row) || 'Single_Record';

  const COLS = config.cols;
  const tableBody = [COLS.map(c => c.get(row))];
  const HEADER_ROW_H = 26;
  const DATA_ROW_H   = 9;
  const FONT_SIZE    = 8.2;
  const HDR_FONT_SZ  = 8.8;

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
      doc.text(`${config.agencyHeader} : ${agencyName}`, MARGIN + 1, MARGIN + 3.5);

      doc.setFillColor(...COLOR_NAVY);
      doc.rect(0, PAGE_H - MARGIN - 3, PAGE_W, MARGIN + 3, 'F');

      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7);
      doc.setTextColor(200, 200, 200);
      doc.text(`Generated: ${new Date().toLocaleDateString('en-IN')}  •  created by pratik verma`, MARGIN, PAGE_H - MARGIN - 0.8);
    },
  });

  const cName = (val(row, 'Customer Name', 'CUSTOMER', 'Customer') || 'Customer').replace(/[^a-z0-9]/gi, '_');
  const lNo  = (val(row, 'Loan Number', 'LOAN NO', 'Loan No') || 'Record').replace(/[^a-z0-9]/gi, '_');
  doc.save(`${cName}_${lNo}_Recovery.pdf`);
}

/* ═══════════════════════════════════════════════
   10. HELPERS
   ═══════════════════════════════════════════════ */

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function refreshAgenciesAndStats() {
  const config = getFormatConfig(detectedFormat, 291);
  const agencies = [...new Set(parsedRows.map(r => config.getAgency(r) || 'Unassigned Agency'))].filter(Boolean).sort();
  const currentSel = agencySelect.value;
  agencySelect.innerHTML = '<option value="ALL">-- Generate All (Download as ZIP) --</option>';
  agencies.forEach(ag => {
    const opt = document.createElement('option');
    opt.value = ag;
    opt.textContent = ag;
    agencySelect.appendChild(opt);
  });
  if (agencies.includes(currentSel)) {
    agencySelect.value = currentSel;
  } else {
    agencySelect.value = 'ALL';
  }
  updateStats();
}

function updateStats() {
  const config = getFormatConfig(detectedFormat, 291);
  const sel = agencySelect.value;
  const filtered = sel === 'ALL' ? parsedRows : parsedRows.filter(r => (config.getAgency(r) || 'Unassigned Agency') === sel);
  statRecords.textContent = filtered.length.toLocaleString();
  statPages.textContent   = '~' + Math.ceil(filtered.length / 35);
  statCols.textContent    = currentHeaders.length;
}

function setProgress(pct, label) {
  progressFill.style.width  = pct + '%';
  progressLabel.textContent = label;
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function resetAll() {
  parsedRows = [];
  currentHeaders = [];
  detectedFormat = null;
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
  statCols.textContent    = '0';
}

/* ═══════════════════════════════════════════════
   11. PWA & WEB SHARE TARGET / FILE HANDLING
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
        break;
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

        window.history.replaceState({}, document.title, window.location.pathname);
        handleFile(file);
        await cache.delete('/shared-file');
      }
    } catch (err) {
      console.error('Failed to load shared file:', err);
    }
  }
});
