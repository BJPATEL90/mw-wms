/**
 * Modular split from Code.gs
 * File: 05-operations-current-inventory.gs
 */

function saveAdjustmentLog(rows) {
  try {
    if (!rows || !rows.length) return { success: false, message: 'No adjustment rows' };

    const sh = SS.getSheetByName('ADJUSTMENT_LOG') || SS.insertSheet('ADJUSTMENT_LOG');
    const headers = ['Date', 'Time', 'Bin / Shelf', 'SKU Code', 'SKU Name', 'Batch No',
                     'MFG Date', 'EXP Date', 'Stock Type', 'Qty Moved', 'Action',
                     'Reference', 'Saved At', 'By'];

    // Write header only if sheet is brand new / empty
    if (sh.getLastRow() < 1) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setTabColor('#7c3aed');
    }

    const now = new Date().toISOString();

    const tz_ = Session.getScriptTimeZone();
    const data = rows.map(r => {
      const dt = r.time ? new Date(r.time) : new Date();
      return [
        Utilities.formatDate(dt, tz_, 'dd-MM-yyyy'),   // ⑬ proper date format
        Utilities.formatDate(dt, tz_, 'HH:mm:ss'),
        String(r.bin        || '').trim(),
        String(r.skuCode    || '').trim(),
        String(r.skuName    || '').trim(),
        String(r.batch      || '').trim(),
        "'" + String(r.mfgDate    || '').trim(),   // PASS AS-IS + force text format
        "'" + String(r.expDate    || '').trim(),   // PASS AS-IS + force text format
        String(r.stockType  || 'GOOD').trim(),
        Number(r.qty        || 0),   // negative = REMOVE, positive = ADD
        String(r.action     || '').trim(),
        String(r.ref        || '').trim(),
        now,
        String(r.by         || '').trim()
      ];
    }).filter(r => r[3]); // must have SKU Code

    if (!data.length) return { success: true, message: 'No valid rows to save', count: 0 };

    // APPEND ONLY — never wipe existing rows
    const startRow = sh.getLastRow() + 1;
    // Set MFG Date (col 7) and EXP Date (col 8) to text format BEFORE writing
    sh.getRange(startRow, 7, data.length, 2).setNumberFormat('@');
    sh.getRange(startRow, 1, data.length, headers.length).setValues(data);

    // Colour-code rows: REMOVE = soft red, ADD = soft green
    data.forEach((row, i) => {
      const color = (row[10] === 'REMOVE') ? '#fff5f5' : '#f0fdf4';
      sh.getRange(startRow + i, 1, 1, headers.length).setBackground(color);
    });

    return { success: true, message: 'Adjustment log saved (' + data.length + ' rows appended)', count: data.length };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


/* ════════════════════════════════════════════════════════
   NEW v5: getBatchDates
   Looks up MFG/EXP dates from INVENTORY_DUMP for a given
   SKU + Batch.  Used by Add-Stock form to auto-fill
   dates when the batch already exists elsewhere.
   ════════════════════════════════════════════════════════ */
function getBatchDates(skuCode, batch) {
  try {
    skuCode = String(skuCode || '').trim();
    batch   = String(batch   || '').trim();
    if (!skuCode || !batch) return { found: false };

    const sh = SS.getSheetByName('INVENTORY_DUMP');
    if (!sh || sh.getLastRow() < 2) return { found: false };

    const data = sh.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === skuCode &&
          String(data[i][2] || '').trim() === batch) {
        return {
          found:   true,
          mfgDate: String(data[i][3] || '').trim(),
          expDate: String(data[i][4] || '').trim(),
          skuName: String(data[i][1] || '').trim()
        };
      }
    }
    return { found: false };
  } catch(e) {
    return { found: false, message: e.message };
  }
}

/* ════════════════════════════════════════════════════════
   ⑦ FIX v6: getInventoryByBin — now also returns shelf GP block qty
   ════════════════════════════════════════════════════════ */
function getInventoryByBin(bin) {
  try {
    bin = normalizeBinIdForMaster_(bin);
    if (!bin) return { success: false, message: 'Bin is required' };

    const sh = SS.getSheetByName('INVENTORY_DUMP');
    if (!sh || sh.getLastRow() < 2) return { success: true, rows: [], sameBatchElsewhere: [], gpBlockByBin: 0 };

    const data = sh.getDataRange().getDisplayValues();
    const binRows = [];
    const allRows = [];

    for (let i = 1; i < data.length; i++) {
      const r      = data[i];
      const rowBin = normalizeBinIdForMaster_(r[6]);
      const qty    = Number(r[9] || 0);
      if (!rowBin || qty <= 0) continue;
      const rowObj = {
        skuCode:   String(r[0]  || '').trim(),
        skuName:   String(r[1]  || '').trim(),
        batch:     String(r[2]  || '').trim(),
        mfgDate:   String(r[3]  || '').trim(),   // PASS AS-IS from INVENTORY_DUMP
        expDate:   String(r[4]  || '').trim(),   // PASS AS-IS from INVENTORY_DUMP
        bin:       rowBin,
        stockType: String(r[7]  || 'GOOD').trim(),
        qty:       qty,
        ean:       String(r[10] || '').trim()
      };
      allRows.push(rowObj);
      if (rowBin === bin) binRows.push(rowObj);
    }

    // Build same-batch-on-other-shelves suggestions
    const sameBatchElsewhere = [];
    binRows.forEach(br => {
      if (!br.batch) return;
      allRows
        .filter(r => r.bin !== bin && r.skuCode === br.skuCode && r.batch === br.batch)
        .forEach(o => {
          const dup = sameBatchElsewhere.find(x => x.fromBin === o.bin && x.skuCode === o.skuCode && x.batch === o.batch);
          if (!dup) sameBatchElsewhere.push({
            fromBin: o.bin, skuCode: o.skuCode, skuName: o.skuName,
            batch: o.batch, mfgDate: o.mfgDate, expDate: o.expDate,
            stockType: o.stockType, qty: o.qty
          });
        });
    });

    // ⑦ v6: Get GP block qty for this specific shelf
    const gpBlockByBin = getGPBlockByBin_(bin);

    return { success: true, rows: binRows, sameBatchElsewhere, gpBlockByBin };
  } catch(e) {
    return { success: false, message: e.message, rows: [], sameBatchElsewhere: [], gpBlockByBin: 0 };
  }
}

/* ════════════════════════════════════════════════════════
   ⑦ NEW v6: getGPBlockByBin — returns GP blocked qty for a specific shelf
   ════════════════════════════════════════════════════════ */
function getGPBlockByBin(bin) {
  try {
    return { success: true, gpBlock: getGPBlockByBin_(bin) };
  } catch(e) {
    return { success: false, gpBlock: 0 };
  }
}

function getGPBlockByBin_(bin) {
  try {
    bin = normalizeBinIdForMaster_(bin);
    const sh = SS.getSheetByName('GATEPASS');
    if (!sh || sh.getLastRow() < 2) return 0;
    const data = sh.getDataRange().getDisplayValues();
    // GATEPASS cols: 0:Gatepass No, 1:Date, 2:SKU Code, 3:SKU Name, 4:Bin/Shelf, 5:Qty
    let total = 0;
    for (let i = 1; i < data.length; i++) {
      if (normalizeBinIdForMaster_(data[i][4]) === bin) {
        total += Number(data[i][5] || 0);
      }
    }
    return total;
  } catch(e) { return 0; }
}

/* ════════════════════════════════════════════════════════
   ⑤⑥ NEW v6: syncCurrentInventory
   Rebuilds the CURRENT_INVENTORY sheet from INVENTORY_DUMP.
   Aggregates at Bin + SKU + Batch level (no duplicates).
   Also pulls GP block qty per shelf from GATEPASS sheet.
   Can be called:
     a) From the HTML "Sync Inventory" button (payload = full inventory rows)
     b) Automatically after any stock movement
   ════════════════════════════════════════════════════════ */
function syncCurrentInventory(payload) {
  try {
    const sh = SS.getSheetByName('CURRENT_INVENTORY') || SS.insertSheet('CURRENT_INVENTORY');
    const headers = ['Bin / Shelf', 'SKU Code', 'SKU Name', 'Batch No', 'MFG Date',
                     'EXP Date', 'MRP', 'Stock Type', 'SKU Status', 'System Qty',
                     'GP Block Qty', 'Last Updated'];

    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#0891b2').setFontColor('#fff').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setTabColor('#0891b2');

    const srcSh = SS.getSheetByName('INVENTORY_DUMP');
    if (!srcSh || srcSh.getLastRow() < 2) {
      return { success: true, message: 'No data', count: 0 };
    }

    const srcData = srcSh.getDataRange().getDisplayValues();
    const map = {};

    for (let i = 1; i < srcData.length; i++) {
      const r = srcData[i];
      const skuCode   = String(r[0] || '').trim();
      const skuName   = String(r[1] || '').trim();
      const batch     = String(r[2] || '').trim();
      const mfgDate   = String(r[3] || '').trim();   // PASS AS-IS from INVENTORY_DUMP
      const expDate   = String(r[4] || '').trim();   // PASS AS-IS from INVENTORY_DUMP
      const mrp       = String(r[5] || '').trim();
      const bin       = normalizeBinIdForMaster_(r[6]);
      const stockType = String(r[7] || 'GOOD').trim().toUpperCase();
      const skuStatus = String(r[8] || 'ACTIVE').trim().toUpperCase();
      const qty       = Number(r[9] || 0);

      if (!bin || !skuCode || qty <= 0) continue;

      const key = `${bin}|${skuCode}|${batch}|${stockType}`;

      if (!map[key]) {
        map[key] = { 
          bin, skuCode, skuName, batch, 
          mfgDate,      // Already correct format from dump
          expDate, 
          mrp, stockType, skuStatus, qty: 0 
        };
      }
      map[key].qty += qty;
    }

    const gpMap = getGPBlockMap_();

    const now = new Date().toLocaleString('en-IN');
    const rows = Object.values(map).map(r => [
      r.bin, r.skuCode, r.skuName, r.batch,
      "'" + r.mfgDate,    // Prepend apostrophe to force text format
      "'" + r.expDate,    // Prepend apostrophe to force text format
      r.mrp,
      r.stockType, r.skuStatus,
      r.qty,
      gpMap[r.bin] || 0,
      now
    ]);

    rows.sort((a, b) => String(a[0]).localeCompare(String(b[0]), {numeric:true}) || String(a[1]).localeCompare(String(b[1])));

    if (rows.length) {
      // Set format to text BEFORE writing data to prevent auto-conversion
      sh.getRange(2, 5, rows.length, 2).setNumberFormat('@');
      sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    sh.autoResizeColumns(1, headers.length);
    return { success: true, message: 'Current Inventory synced', count: rows.length };

  } catch(e) {
    console.error(e);
    return { success: false, message: e.message };
  }
}

/** Helper Function - GP Block per Bin */
function getGPBlockMap_() {
  const map = {};
  const gpSh = SS.getSheetByName('GATEPASS');
  if (!gpSh || gpSh.getLastRow() < 2) return map;

  const data = gpSh.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    const bin = normalizeBinIdForMaster_(data[i][4]);   // Bin/Shelf column
    const qty = Number(data[i][5] || 0);           // Qty column
    if (bin && qty > 0) {
      map[bin] = (map[bin] || 0) + qty;
    }
  }
  return map;
}

/* ════════════════════════════════════════════════════════
   ⑧ NEW v6: saveAdjustmentLogIdempotent
   Duplicate prevention: checks if an entry with the same
   idempotency key already exists before appending.
   The HTML sends a unique key per button-click.
   ════════════════════════════════════════════════════════ */
function saveAdjustmentLogIdempotent(rows, idempotencyKey) {
  try {
    if (!rows || !rows.length) return { success: false, message: 'No rows' };
    if (!idempotencyKey) return saveAdjustmentLog(rows); // fallback

    const sh = SS.getSheetByName('ADJUSTMENT_LOG');
    if (sh && sh.getLastRow() >= 2) {
      // Check if this key was already saved (Reference column = col 12, 0-indexed 11)
      const refs = sh.getRange(2, 12, sh.getLastRow() - 1, 1).getValues().flat().map(String);
      if (refs.some(r => r.includes('key:' + idempotencyKey))) {
        return { success: true, message: 'Already saved (duplicate prevented)', count: 0, duplicate: true };
      }
    }

    // Tag the ref field with the idempotency key
    rows = rows.map(r => ({ ...r, ref: (r.ref || '') + ' [key:' + idempotencyKey + ']' }));
    return saveAdjustmentLog(rows);
  } catch(e) {
    return { success: false, message: e.message };
  }
}


/* ════════════════════════════════════════════════════════
   ⑭⑯ NEW v7: normaliseDate_
   Converts various date formats to DD-MM-YYYY string:
   • Excel serial number (e.g. 44927 → 01-01-2023)
   • ISO string / JS Date string
   • Already-formatted strings (passthrough)
   Returns empty string for blank/null values.
   ════════════════════════════════════════════════════════ */
function normaliseDate_(val) {
  if (val === null || val === undefined || val === '') return '';
  const s = String(val).trim();
  if (!s) return '';

  // Pure number → Excel serial date
  if (/^\d{5}(\.\d+)?$/.test(s)) {
    const excelSerial = Math.floor(Number(s));
    // Excel epoch is 1 Jan 1900; JS epoch is 1 Jan 1970
    // Excel wrongly counts 1900 as a leap year, hence -2 offset
    const jsDate = new Date((excelSerial - 25569) * 86400 * 1000);
    try {
      return Utilities.formatDate(jsDate, Session.getScriptTimeZone(), 'dd-MM-yyyy');
    } catch(e) { return s; }
  }

  // Try parsing as a known date string
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    try {
      return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'dd-MM-yyyy');
    } catch(e) { return s; }
  }

  // Already a string like "01-Jan-2024" or "dd-MM-yyyy" — return as-is
  return s;
}

/* ════════════════════════════════════════════════════════
   ⑩ NEW v7: clearSheetData
   Deletes all DATA rows from a named sheet but keeps header.
   Works for: SKU_MASTER, BIN_MASTER_WMS, INVENTORY_DUMP,
              GATEPASS, CYCLE_COUNT_LOG, ADJUSTMENT_LOG,
              CURRENT_INVENTORY, GRN_LOG, and any other sheet.
   ════════════════════════════════════════════════════════ */
