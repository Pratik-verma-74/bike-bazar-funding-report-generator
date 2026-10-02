/**
 * Loan Recovery PDF Converter — script.js
 * ================================================
 * Libraries expected (loaded in index.html):
 *   XLSX (SheetJS)  — Excel parsing
 *   jsPDF           — PDF generation
 *   jspdf-autotable — Table plugin for jsPDF
 *   JSZip           — ZIP packaging
 *   FileSaver       — File download helper
 * ================================================
 */

/* ═══════════════════════════════════════════════
   1. COLUMN MAPPING CONFIGURATION
   Supports Format 1 (Standard Bike Bazar) and
   Format 2 (DMA / Agency Sheet) seamlessly.
   ═══════════════════════════════════════════════ */

const FIELD_DEFS = [
  {
    key: 'loanNo',
    label: 'Loan No',
    required: true,
    aliases: ['loan no', 'loan number', 'loan num', 'loan_no', 'loanno', 'loan account', 'account no', 'account number', 'lan', 'agreement no', 'loan no.'],
  },
  {
    key: 'customerName',
    label: 'Customer Name',
    required: true,
    aliases: ['customer name', 'customer', 'cust name', 'borrower', 'borrower name', 'applicant name', 'client name', 'name'],
  },
  {
    key: 'emi',
    label: 'EMI',
    required: true,
    aliases: ['emi', 'emi amount', 'monthly emi', 'instalment', 'installment', 'installment amount'],
  },
  {
    key: 'bkt',
    label: 'BKT',
    required: true,
    aliases: ['bkt', 'opening bkt - 7+', 'opening bkt', 'bucket', 'opening bucket', 'dpd bucket', 'bucket no', 'curr bkt'],
  },
  {
    key: 'pos',
    label: 'POS',
    required: true,
    aliases: ['total pos', 'pos', 'principal outstanding', 'total pos amount', 'pos amount', 'pos amt'],
  },
  {
    key: 'openingOD',
    label: 'Opening OD',
    required: true,
    aliases: ['opening od', 'opnig od', 'opning od', 'open od', 'opn od', 'opning o', 'opening overdue'],
  },
  {
    key: 'totalOD',
    label: 'Total OD',
    required: true,
    aliases: ['total od', 'totl od', 'tot od', 'total overdue', 'od amount', 'od amt'],
  },
  {
    key: 'piDues',
    label: 'PI Dues',
    required: false,
    aliases: ['pi dues', 'total pi', 'opening pi', 'pidues', 'pi due', 'pi amt', 'pi amount', 'pi charge', 'pi charges'],
  },
  {
    key: 'address',
    label: 'Address',
    required: true,
    aliases: ['permanent address', 'permnt add', 'current address', 'perm address', 'customer address', 'residential address', 'residence address', 'address', 'addr'],
  },
  {
    key: 'pinCode',
    label: 'Pin Code',
    required: false,
    aliases: ['pin code', 'pincode', 'pin', 'postal code', 'zip code'],
  },
  {
    key: 'mobile',
    label: 'Mobile',
    required: true,
    aliases: ['contact number', 'contact no', 'contact', 'customer mobile', 'mobile no', 'mobile number', 'mobile', 'phone no', 'phone', 'mob no', 'primary contact'],
  },
  {
    key: 'altMobile',
    label: 'Alt Mobile',
    required: false,
    aliases: ['alternate mobile no', 'alternate mobile number', 'alt/data mob-1', 'alternate mobile', 'alt mobile', 'alt contact', 'alternate num', 'alt mob 1'],
  },
  {
    key: 'altMobile2',
    label: 'Alt Mobile 2',
    required: false,
    aliases: ['alt/data mob-2', 'alt mobile 2', 'alternate mobile 2', 'alt mob 2'],
  },
  {
    key: 'ref1Name',
    label: 'Ref1 Name',
    required: false,
    aliases: ['data/ref-1', 'data / ref-1', 'reference-1', 'reference 1', 'ref1 name', 'ref 1 name', 'ref 1', 'reference1', 'ref1', 'guarantor 1'],
  },
  {
    key: 'ref1Mobile',
    label: 'Ref1 Mobile',
    required: false,
    aliases: ['reference-1 mobile no', 'reference 1 mobile no', 'reference 1 mobile', 'ref1 mobile', 'ref 1 mobile', 'ref-1 mobile', 'ref1mobile', 'guarantor 1 mobile'],
  },
  {
    key: 'ref2Name',
    label: 'Ref2 Name',
    required: false,
    aliases: ['data/ref-2', 'data / ref-2', 'reference-2', 'reference 2', 'ref2 name', 'ref 2 name', 'ref 2', 'reference2', 'ref2', 'guarantor 2'],
  },
  {
    key: 'ref2Mobile',
    label: 'Ref2 Mobile',
    required: false,
    aliases: ['reference-2 mobile no', 'reference 2 mobile no', 'reference 2 mobile', 'ref2 mobile', 'ref 2 mobile', 'ref-2 mobile', 'ref2mobile', 'guarantor 2 mobile'],
  },
  {
    key: 'regNo',
    label: 'Reg No',
    required: true,
    aliases: ['reg no', 'registration no', 'registration number', 'vehicle no', 'vehicle number', 'vehicle no.', 'reg_no', 'reg. no.', 'rc no', 'veh no'],
  },
  {
    key: 'make',
    label: 'Make',
    required: true,
    aliases: ['make', 'vehicle make', 'asset make', 'make name', 'vehicle model', 'model', 'vehicle'],
  },
  {
    key: 'chassisNo',
    label: 'Chassis No',
    required: false,
    aliases: ['chassis no', 'chassis no.', 'chassis number', 'chassis', 'vin'],
  },
  {
    key: 'engineNo',
    label: 'Engine No',
    required: false,
    aliases: ['engine no', 'engine no.', 'engine number', 'engine'],
  },
  {
    key: 'mfgYr',
    label: 'Mfg Year',
    required: false,
    aliases: ['mfg yr', 'mfg year', 'manufacturing year', 'model year', 'year'],
  },
  {
    key: 'dealerName',
    label: 'DLR Name',
    required: true,
    aliases: ['dealer name', 'dealer', 'dlr name', 'showroom', 'showroom name', 'dealer/showroom', 'dlr'],
  },
  {
    key: 'agencyName',
    label: 'Agency Name',
    required: true,
    aliases: ['agency/dma name', 'agency/dma', 'agency name', 'dma name', 'dma/dma name', 'agency', 'dma'],
  },
  {
    key: 'tcmName',
    label: 'TCM Name',
    required: false,
    aliases: ['tcm name', 'tcm', 'telecaller', 'telecaller name'],
  },
];

/* ═══════════════════════════════════════════════
   2. STATE
   ═══════════════════════════════════════════════ */
let parsedRows = [];        // Raw data rows from Excel
let mappedKeys = {};        // { fieldKey: excelColumnName }
let currentExcelCols = [];  // Array of column header strings from current file
let currentFoundRow = null; // Currently matched search row

/* ═══════════════════════════════════════════════
   3. DOM REFERENCES
   ═══════════════════════════════════════════════ */
const dropZone          = document.getElementById('dropZone');
const fileInput         = document.getElementById('fileInput');
const fileInfo          = document.getElementById('fileInfo');
const fileNameEl        = document.getElementById('fileName');
const recordCountEl     = document.getElementById('recordCount');
const clearFileBtn      = document.getElementById('clearFile');
const mappingSection    = document.getElementById('mapping-section');
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
const generateBtn       = document.getElementById('generateBtn');
const genBtnText        = document.getElementById('genBtnText');
const spinner           = document.getElementById('spinner');
const progressWrap      = document.getElementById('progressWrap');
const progressFill      = document.getElementById('progressFill');
const progressLabel     = document.getElementById('progressLabel');

/* ═══════════════════════════════════════════════
   4. DRAG & DROP / FILE INPUT
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
   5. FILE HANDLER
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

      // Detect header row index (scanning first 5 rows for standard keywords)
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
      currentExcelCols = cleanHeaders;
      mappedKeys = autoMapColumns(cleanHeaders);

      // Show file info
      fileNameEl.textContent = file.name;
      recordCountEl.textContent = `${rows.length.toLocaleString()} records`;
      fileInfo.classList.remove('hidden');

      // Render Mapping Grid
      renderMappingGrid(mappedKeys, currentExcelCols);

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
   6. INTELLIGENT COLUMN MAPPING ENGINE
   - Multi-pass precision matching
   - Disambiguates Make vs Reg No vs Chassis No
   - Reserves matched columns so they aren't stolen
   ═══════════════════════════════════════════════ */

function normalize(s) {
  return String(s || '').toLowerCase().replace(/[\s_\-\.\/]+/g, ' ').trim();
}

function autoMapColumns(excelCols) {
  const result = {};
  const usedCols = new Set();

  // Pass 1: Exact matches first
  for (const def of FIELD_DEFS) {
    for (const alias of def.aliases) {
      const a = normalize(alias);
      const exact = excelCols.find(c => !usedCols.has(c) && normalize(c) === a);
      if (exact) {
        result[def.key] = exact;
        usedCols.add(exact);
        break;
      }
    }
  }

  // Pass 2: Partial matches with boundary and collision guards
  for (const def of FIELD_DEFS) {
    if (result[def.key]) continue;
    let bestCol = null;
    let bestScore = 0;

    for (const alias of def.aliases) {
      const a = normalize(alias);
      for (const col of excelCols) {
        if (usedCols.has(col)) continue;
        const normCol = normalize(col);

        // Guard: Prevent 'make' from grabbing 'vehicle no' or 'vehicle number'
        if (def.key === 'make' && (normCol.includes('no') || normCol.includes('num'))) continue;
        // Guard: Prevent 'customerName' from grabbing other names (agency name, dealer name, etc.)
        if (def.key === 'customerName' && (normCol.includes('agency') || normCol.includes('dealer') || normCol.includes('dlr') || normCol.includes('tcm') || normCol.includes('ref'))) continue;
        // Guard: 'bkt' must relate to bucket
        if (def.key === 'bkt' && !normCol.includes('bkt') && !normCol.includes('bucket')) continue;
        // Guard: 'address' shouldn't match IP or email address
        if (def.key === 'address' && (normCol.includes('ip') || normCol.includes('email') || normCol.includes('mail'))) continue;

        if (normCol === a) {
          if (1000 > bestScore) { bestScore = 1000; bestCol = col; }
        } else if (normCol.startsWith(a) || normCol.endsWith(a)) {
          if (500 > bestScore) { bestScore = 500; bestCol = col; }
        } else if (normCol.includes(a)) {
          const score = 100 + a.length;
          if (score > bestScore) { bestScore = score; bestCol = col; }
        } else if (a.includes(normCol) && normCol.length >= 4) {
          const score = 50 + normCol.length;
          if (score > bestScore) { bestScore = score; bestCol = col; }
        }
      }
    }
    if (bestCol) {
      result[def.key] = bestCol;
      usedCols.add(bestCol);
    }
  }

  return result;
}

/* ═══════════════════════════════════════════════
   7. MAPPING PREVIEW GRID & USER OVERRIDE
   ═══════════════════════════════════════════════ */

function renderMappingGrid(map, excelCols) {
  mappingGrid.innerHTML = '';

  // Detect format banner
  const hasFormat1 = map['bkt'] && /7\+|opening/i.test(map['bkt']);
  const hasFormat2 = (map['ref1Name'] && /ref/i.test(map['ref1Name'])) || (map['loanNo'] && /LOAN NO/i.test(map['loanNo']));

  const badgeDiv = document.getElementById('formatBadge');
  if (badgeDiv) {
    if (hasFormat1) {
      badgeDiv.className = 'format-badge fmt-1';
      badgeDiv.innerHTML = '⚡ Detected Format: <strong>Format 1 (Bike Bazar Standard)</strong> — All columns mapped successfully!';
    } else if (hasFormat2) {
      badgeDiv.className = 'format-badge fmt-2';
      badgeDiv.innerHTML = '⚡ Detected Format: <strong>Format 2 (DMA / Agency Sheet)</strong> — All columns mapped successfully!';
    } else {
      badgeDiv.className = 'format-badge fmt-custom';
      badgeDiv.innerHTML = '⚡ Detected Format: <strong>Custom Excel Sheet</strong>';
    }
  }

  for (const fieldDef of FIELD_DEFS) {
    const matched = map[fieldDef.key];
    const isRequired = fieldDef.required;
    const div = document.createElement('div');

    let statusClass = 'missing';
    let statusText = '';

    if (matched) {
      statusClass = 'found';
      statusText = `✓ → ${matched}`;
    } else if (!isRequired) {
      statusClass = 'optional';
      statusText = `ℹ Optional (Blank in this format)`;
    } else {
      statusClass = 'missing';
      statusText = `✕ Missing — will be blank`;
    }

    div.className = `map-item ${statusClass}`;
    div.innerHTML = `
      <div class="map-header">
        <span class="map-field">${escapeHtml(fieldDef.label)}</span>
        ${isRequired ? '<span class="req-tag">Required</span>' : '<span class="opt-tag">Optional</span>'}
      </div>
      <div class="map-source" id="source-${fieldDef.key}">${escapeHtml(statusText)}</div>
      <select class="map-select" data-key="${fieldDef.key}">
        <option value="">-- Manual Override / None --</option>
        ${excelCols.map(c => `<option value="${escapeHtml(c)}" ${c === matched ? 'selected' : ''}>${escapeHtml(c)}</option>`).join('')}
      </select>
    `;

    // Dropdown listener for manual reassignment
    const select = div.querySelector('select');
    select.addEventListener('change', e => {
      const chosen = e.target.value;
      mappedKeys[fieldDef.key] = chosen || null;
      const srcEl = div.querySelector('.map-source');
      if (chosen) {
        div.className = 'map-item found';
        srcEl.textContent = `✓ → ${chosen}`;
      } else if (!fieldDef.required) {
        div.className = 'map-item optional';
        srcEl.textContent = `ℹ Optional (Blank in this format)`;
      } else {
        div.className = 'map-item missing';
        srcEl.textContent = `✕ Missing — will be blank`;
      }
      refreshAgenciesAndStats();
    });

    mappingGrid.appendChild(div);
  }
}

/* ═══════════════════════════════════════════════
   8. EXTRACT VALUE FROM ROW
   ═══════════════════════════════════════════════ */

function get(row, key) {
  const col = mappedKeys[key];
  if (!col) return '';
  const val = row[col];
  if (val === undefined || val === null) return '';
  return String(val).trim();
}

/* ═══════════════════════════════════════════════
   9. COMPREHENSIVE TABLE COLUMNS (12 COLUMNS)
   Covers every data point from Format 1 & 2
   ═══════════════════════════════════════════════ */

function getTableColumns(usableWidth) {
  const COLS = [
    { w: 22, hdr: 'Loan No', get: r => get(r, 'loanNo') },
    { w: 25, hdr: 'Customer Name', get: r => get(r, 'customerName') },
    { w: 6,  hdr: 'BKT', get: r => get(r, 'bkt') },
    { w: 11, hdr: 'POS', get: r => get(r, 'pos') },
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
    { w: 8,  hdr: 'EMI', get: r => get(r, 'emi') },
    { w: 10, hdr: 'PI Dues', get: r => get(r, 'piDues') },
    {
      w: 44,
      hdr: 'Permanent Address',
      get: r => {
        const addr = get(r, 'address');
        const pin = get(r, 'pinCode');
        if (pin && !addr.includes(pin)) {
          return addr ? `${addr}\nPIN: ${pin}` : `PIN: ${pin}`;
        }
        return addr;
      }
    },
    {
      w: 44,
      hdr: 'Contact / Reference',
      get: r => {
        const mob   = get(r, 'mobile');
        const alt   = get(r, 'altMobile');
        const alt2  = get(r, 'altMobile2');
        const ref1n = get(r, 'ref1Name');
        const ref1m = get(r, 'ref1Mobile');
        const ref2n = get(r, 'ref2Name');
        const ref2m = get(r, 'ref2Mobile');
        const parts = [];
        if (mob) parts.push('Mob: ' + mob);
        if (alt) parts.push('Alt: ' + alt);
        if (alt2) parts.push('Alt2: ' + alt2);
        if (ref1n || ref1m) {
          const r1 = [ref1n, ref1m ? `(${ref1m})` : ''].filter(Boolean).join(' ');
          parts.push('Ref1: ' + r1);
        }
        if (ref2n || ref2m) {
          const r2 = [ref2n, ref2m ? `(${ref2m})` : ''].filter(Boolean).join(' ');
          parts.push('Ref2: ' + r2);
        }
        return parts.join('\n');
      }
    },
    {
      w: 13,
      hdr: 'Make / Model',
      get: r => {
        const mk = get(r, 'make');
        const yr = get(r, 'mfgYr');
        return yr ? `${mk}\nYr: ${yr}` : mk;
      }
    },
    {
      w: 16,
      hdr: 'REG / Chassis',
      get: r => {
        const reg = get(r, 'regNo');
        const ch = get(r, 'chassisNo');
        const eng = get(r, 'engineNo');
        const parts = [];
        if (reg) parts.push(reg);
        if (ch) parts.push('Ch: ' + ch);
        if (eng) parts.push('Eng: ' + eng);
        return parts.join('\n');
      }
    },
    {
      w: 16,
      hdr: 'DLR / TCM',
      get: r => {
        const dlr = get(r, 'dealerName');
        const tcm = get(r, 'tcmName');
        const parts = [];
        if (dlr) parts.push(dlr);
        if (tcm) parts.push('TCM: ' + tcm);
        return parts.join('\n');
      }
    },
  ];

  const totalDefined = COLS.reduce((s, c) => s + c.w, 0);
  const scale = usableWidth / totalDefined;
  COLS.forEach(c => c.w = c.w * scale);
  return COLS;
}

/* ═══════════════════════════════════════════════
   10. PDF GENERATION (BATCH / ZIP)
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

    // Group by Agency (preserves unassigned rows under 'Unassigned Agency')
    const groups = {};
    for (const r of parsedRows) {
      let rawAgency = get(r, 'agencyName') || 'Unassigned Agency';
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

      const MARGIN   = 3;
      const PAGE_W   = 297;
      const PAGE_H   = 210;
      const USABLE_W = PAGE_W - MARGIN * 2;

      const COLS = getTableColumns(USABLE_W);
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
   11. SEARCH AND SINGLE VIEW
   Searches by Loan No, Customer Name, Mobile, or Reg No
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

  const match = parsedRows.find(r => {
    const ln = get(r, 'loanNo').toLowerCase();
    const cn = get(r, 'customerName').toLowerCase();
    const mb = get(r, 'mobile').toLowerCase();
    const rn = get(r, 'regNo').toLowerCase();
    return ln === query || ln.includes(query) || cn.includes(query) || mb.includes(query) || rn.includes(query);
  });

  if (match) {
    currentFoundRow = match;
    const fieldsToShow = [
      'loanNo', 'customerName', 'agencyName', 'pos', 'openingOD', 'totalOD',
      'piDues', 'emi', 'bkt', 'mobile', 'altMobile', 'altMobile2',
      'address', 'pinCode', 'ref1Name', 'ref1Mobile', 'ref2Name', 'ref2Mobile',
      'make', 'mfgYr', 'regNo', 'chassisNo', 'engineNo', 'dealerName', 'tcmName'
    ];
    resultGrid.innerHTML = '';
    for (const k of fieldsToShow) {
      const def = FIELD_DEFS.find(d => d.key === k);
      const val = get(match, k);
      if (val && def) {
        resultGrid.innerHTML += `
          <div class="r-item">
            <div class="r-lbl">${escapeHtml(def.label)}</div>
            <div class="r-val">${escapeHtml(val)}</div>
          </div>`;
      }
    }
    searchResult.classList.remove('hidden');
  } else {
    searchError.textContent = `No record found matching "${searchInput.value.trim()}". Try Loan No, Customer Name, Mobile, or Reg No.`;
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
  const agencyName = get(row, 'agencyName') || 'Single_Record';

  const COLS = getTableColumns(USABLE_W);
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
   12. HELPERS
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
  const agencies = [...new Set(parsedRows.map(r => get(r, 'agencyName') || 'Unassigned Agency'))].filter(Boolean).sort();
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
  const sel = agencySelect.value;
  const filtered = sel === 'ALL' ? parsedRows : parsedRows.filter(r => (get(r, 'agencyName') || 'Unassigned Agency') === sel);
  statRecords.textContent = filtered.length.toLocaleString();
  statPages.textContent   = '~' + Math.ceil(filtered.length / 35);
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
  mappedKeys = {};
  currentExcelCols = [];
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
   13. PWA & WEB SHARE TARGET / FILE HANDLING
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
