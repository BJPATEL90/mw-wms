/**
 * Modular split from Code.gs
 * File: 04-wms-save-load.gs
 */

function syncWMSBinStatusToMaster_(binData) {
  const sh = SS.getSheetByName('BIN_MASTER');
  if (!sh) return;
  const d = sh.getDataRange().getDisplayValues();
  const statusMap = {};
  binData.forEach(r => { if (r[0]) statusMap[r[0]] = r[4]; });
  for (let i = 1; i < d.length; i++) {
    const bin = String(d[i][0] || '').trim();
    if (bin && statusMap[bin]) sh.getRange(i + 1, 3).setValue(statusMap[bin].toUpperCase());
  }
}

/**
 * Save Inventory Dump from HTML upload.
 * rows: [{skuCode, skuName, batch, mfgDate, expDate, mrp, bin, stockType, skuStatus, qty, ean}]
 * mode: 'replace' (default) — clears sheet first | 'append' — adds to existing
 */
function saveInventoryDump(rows, mode) {
  _invalidateCache_('binMaster'); // bin statuses change after dump upload
  try {
    if (!rows || !rows.length) return { success: false, message: 'No inventory rows received' };
    const sh = SS.getSheetByName('INVENTORY_DUMP') || SS.insertSheet('INVENTORY_DUMP');
    const headers = ['SKU Code', 'SKU Name', 'Batch No', 'MFG Date', 'EXP Date', 'MRP',
                     'Bin / Shelf', 'Stock Type', 'SKU Status', 'Qty', 'EAN',
                     'Uploaded At', 'Uploaded By'];

    if (!mode || mode === 'replace') {
      sh.clearContents();
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
    }

    const now = new Date().toISOString();
    const data = rows.map(r => [
      String(r.skuCode || '').trim(),
      String(r.skuName || '').trim(),
      String(r.batch || '').trim(),
      "'" + String(r.mfgDate || '').trim(),    // PASS AS-IS + force text
      "'" + String(r.expDate || '').trim(),    // PASS AS-IS + force text
      String(r.mrp || '').trim(),
      normalizeBinIdForMaster_(r.bin),
      String(r.stockType || 'GOOD').trim().toUpperCase(),
      String(r.skuStatus || 'ACTIVE').trim().toUpperCase(),
      Number(r.qty || 0),
      String(r.ean || '').trim(),
      now,
      'WMS App'
    ]).filter(r => r[0] && Number(r[9]) > 0);

    if (data.length) {
      const startRow = sh.getLastRow() + 1;
      // Set MFG Date (col 4) and EXP Date (col 5) to text format BEFORE writing data
      sh.getRange(startRow, 4, data.length, 2).setNumberFormat('@');
      sh.getRange(startRow, 1, data.length, headers.length).setValues(data);
    }

    // Also sync bin occupancy to BIN_MASTER
    try { updateBinMasterFromDump_(data); } catch(e2) { /* non-fatal */ }

    return { success: true, message: 'Inventory Dump saved', count: data.length };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/** Updates BIN_MASTER.Status from freshly uploaded inventory dump */
function updateBinMasterFromDump_(dumpData) {
  const binSh = SS.getSheetByName('BIN_MASTER');
  if (!binSh) return;
  const binData = binSh.getDataRange().getDisplayValues();

  // Build set of bins that have stock
  const occupiedBins = new Set();
  dumpData.forEach(r => {
    const bin = normalizeBinIdForMaster_(r[6]);
    if (bin && Number(r[9]) > 0) occupiedBins.add(bin);
  });

  for (let i = 1; i < binData.length; i++) {
    const bin = normalizeBinIdForMaster_(binData[i][0]);
    const blocked = String(binData[i][3] || '').trim().toUpperCase();
    if (!bin || blocked === 'YES') continue;
    const currentStatus = String(binData[i][2] || '').trim().toUpperCase();
    const newStatus = occupiedBins.has(bin) ? 'OCCUPIED' : (currentStatus === 'OCCUPIED' ? 'EMPTY' : currentStatus);
    if (newStatus !== currentStatus) binSh.getRange(i + 1, 3).setValue(newStatus);
  }
}

function normalizeBinIdForMaster_(value) {
  let bin = String(value || '').trim().toUpperCase();
  if (!bin) return '';
  bin = bin.replace(/\s+/g, '');
  bin = bin.replace(/[.。]+$/g, '');
  const m = bin.match(/^R(\d{1,3})-C(\d{1,3})-(\d{1,3})$/);
  if (!m) return '';
  const row = Number(m[1]);
  const col = Number(m[2]);
  const level = Number(m[3]);
  if (row < 1 || col < 1 || level < 1) return '';
  return 'R' + row + '-C' + col + '-' + String(level).padStart(3, '0');
}

/**
 * Save Gatepass records from HTML upload. Always APPENDS.
 * rows: [{gpNo, date, skuCode, skuName, bin, qty, reason}]
 */
function saveGatepasses(rows) {
  _invalidateCache_('gatepasses');
  try {
    if (!rows || !rows.length) return { success: false, message: 'No gatepass rows received' };
    const sh = SS.getSheetByName('GATEPASS') || SS.insertSheet('GATEPASS');
    const headers = ['Gatepass No', 'Date', 'SKU Code', 'SKU Name', 'Bin / Shelf',
                     'Qty', 'Reason', 'Uploaded At', 'Uploaded By'];

    // Write header if sheet is new/empty
    if (sh.getLastRow() < 1) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
    }

    const now = new Date().toISOString();
    const data = rows.filter(r => r.skuCode && Number(r.qty) > 0).map(r => [
      String(r.gpNo || '').trim(),
      String(r.date || '').trim(),
      String(r.skuCode || '').trim(),
      String(r.skuName || '').trim(),
      String(r.bin || '').trim(),
      Number(r.qty || 0),
      String(r.reason || '').trim(),
      now,
      'WMS App'
    ]);

    if (data.length) sh.getRange(sh.getLastRow() + 1, 1, data.length, headers.length).setValues(data);
    return { success: true, message: 'Gatepasses saved', count: data.length };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Save Cycle Count log from HTML.
 * rows: [{time, bin, skuCode, skuName, batch, mfgDate, expDate, stockType,
 *         systemQty, gpBlock, countedQty, diff, remarks, by}]
 */
function saveCycleCountLog(rows) {
  try {
    if (!rows || !rows.length) return { success: false, message: 'No cycle count rows received' };
    const sh = SS.getSheetByName('CYCLE_COUNT_LOG') || SS.insertSheet('CYCLE_COUNT_LOG');
    const headers = ['Date', 'Time', 'Bin / Shelf', 'SKU Code', 'SKU Name', 'Batch No',
                     'MFG Date', 'EXP Date', 'Stock Type', 'System Qty', 'GP Block Qty',
                     'Counted Qty', 'Difference', 'Status', 'Remarks', 'Counted By'];

    if (sh.getLastRow() < 1) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
    }

    const tz = Session.getScriptTimeZone();
    const data = rows.map(r => {
      const dt   = r.time ? new Date(r.time) : new Date();
      const diff = Number(r.countedQty || 0) - Number(r.systemQty || 0);
      const status = diff === 0 ? 'MATCH' : diff > 0 ? 'EXCESS' : 'SHORT';
      return [
        Utilities.formatDate(dt, tz, 'dd-MM-yyyy'),          // ⑫ proper date format
        Utilities.formatDate(dt, tz, 'HH:mm:ss'),
        String(r.bin || '').trim(),
        String(r.skuCode || '').trim(),
        String(r.skuName || '').trim(),
        String(r.batch || '').trim(),
        "'" + String(r.mfgDate || '').trim(),                // PASS AS-IS + force text format
        "'" + String(r.expDate || '').trim(),                // PASS AS-IS + force text format
        String(r.stockType || 'GOOD').trim(),
        Number(r.systemQty || 0),
        Number(r.gpBlock || 0),
        Number(r.countedQty || 0),
        diff,
        status,
        String(r.remarks || '').trim(),
        String(r.by || '').trim()
      ];
    }).filter(r => r[3]);  // must have SKU Code

    if (data.length) {
      const startRow = sh.getLastRow() + 1;
      // Set MFG Date (col 7) and EXP Date (col 8) to text format BEFORE writing
      sh.getRange(startRow, 7, data.length, 2).setNumberFormat('@');
      sh.getRange(startRow, 1, data.length, headers.length).setValues(data);
    }
    return { success: true, message: 'Cycle Count log saved', count: data.length };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Save GRN to Google Sheet.
 * grn: {grnNo, date, supplier, vehicle, lines:[{sku,skuName,batch,qty,casePack,boxPallet,cls}], savedBy}
 */
function saveGRNToSheet(grn) {
  try {
    if (!grn || !grn.grnNo) return { success: false, message: 'GRN data missing' };
    const sh = SS.getSheetByName('GRN_LOG') || SS.insertSheet('GRN_LOG');
    const headers = ['GRN No', 'Date', 'Supplier', 'Vehicle No', 'SKU Code', 'SKU Name',
                     'Batch No', 'Qty', 'Case Pack', 'Box/Pallet', 'Classification',
                     'Saved By', 'Saved At'];

    if (sh.getLastRow() < 1) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
    }

    const now = new Date().toISOString();
    const lines = grn.lines || [];
    if (!lines.length) return { success: false, message: 'No GRN lines to save' };

    const data = lines.map(l => [
      grn.grnNo, grn.date || '', grn.supplier || '', grn.vehicle || '',
      String(l.sku || l.skuCode || '').trim(),
      String(l.skuName || '').trim(),
      String(l.batch || '').trim(),
      Number(l.qty || 0),
      Number(l.casePack || '') || '',
      Number(l.boxPallet || '') || '',
      String(l.cls || 'B').trim(),
      grn.savedBy || 'WMS App',
      now
    ]);

    sh.getRange(sh.getLastRow() + 1, 1, data.length, headers.length).setValues(data);
    return { success: true, message: 'GRN saved to sheet', count: data.length };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Clear all gatepasses from sheet (called when user clicks "Clear All" in HTML)
 */
function clearGatepasses() {
  try {
    const sh = SS.getSheetByName('GATEPASS');
    if (!sh) return { success: true, message: 'No gatepass sheet found' };
    const headers = sh.getRange(1, 1, 1, 9).getValues()[0];
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
    return { success: true, message: 'Gatepasses cleared' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Get SKU Master from sheet (for loading into HTML app on startup)
 */
function getSKUMasterFromSheet() {
  try {
    const sh = SS.getSheetByName('SKU_MASTER');
    if (!sh || sh.getLastRow() < 2) return [];
    const lastRow = sh.getLastRow();
    const data = sh.getRange(1, 1, lastRow, 10).getDisplayValues(); // 10 cols, no getDataRange
    // Cols: 0=SKU Code,1=SKU Name,2=Brand,3=Case Pack,4=Box/Pallet,5=Class,6=MRP,7=Status
    return data.slice(1).filter(r => r[0]).map(r => ({
      code:      r[0], name: r[1],
      brand:     r[2] || inferBrandFromSKUCode_(String(r[0])),
      casePack:  Number(r[3]) || null,
      boxPallet: Number(r[4]) || null,
      cls:       r[5] || 'B',
      mrp:       r[6] || '',
      skuStatus: r[7] || 'ACTIVE'
    }));
  } catch(e) { return []; }
}

/**
 * Get Bin Master from sheet
 */
function getBinMasterFromSheet() {
  try {
    const sh = SS.getSheetByName('BIN_MASTER_WMS');
    if (!sh || sh.getLastRow() < 2) return [];
    const lastRow = sh.getLastRow();
    const data = sh.getRange(1, 1, lastRow, 11).getDisplayValues(); // 11 cols, no getDataRange
    // BIN_MASTER_WMS columns (10 total): 0=BinID 1=Row 2=ColNo 3=Level 4=BrandZone 5=FSN 6=Status 7=SKUinBin 8=UpdatedAt 9=UpdatedBy
    const byBin = {};
    const statusRank = { EMPTY: 1, OCCUPIED: 2, RESERVED: 3, BLOCKED: 4 };
    data.slice(1).forEach(r => {
      const id = normalizeBinIdForMaster_(r[0]);
      if (!id) return;
      const parsed = id.match(/^R(\d+)-C(\d+)-(\d+)$/);
      const status = String(r[6] || 'Empty').trim() || 'Empty';
      const item = {
        id,
        row:         Number(r[1]) || (parsed ? Number(parsed[1]) : 0),
        col:         Number(r[2]) || (parsed ? Number(parsed[2]) : 0),
        level:       Number(r[3]) || (parsed ? Number(parsed[3]) : 0),
        brand:       String(r[4] || 'Unknown').trim(),
        fsn:         String(r[5] || 'B').trim().toUpperCase(),
        status,
        skuFromDump: String(r[7] || '').trim(),
      };
      const prev = byBin[id];
      if (!prev || (statusRank[status.toUpperCase()] || 1) > (statusRank[String(prev.status).toUpperCase()] || 1)) {
        byBin[id] = item;
      }
    });
    return Object.values(byBin);
  } catch(e) { return []; }
}

/**
 * Get Gatepass records from sheet
 */
function getGatepassesFromSheet() {
  try {
    const sh = SS.getSheetByName('GATEPASS');
    if (!sh || sh.getLastRow() < 2) return [];
    const lastRow = sh.getLastRow();
    const data = sh.getRange(2, 1, lastRow - 1, 12).getDisplayValues();
    return data.filter(r => r[2]).map(r => ({
      gpNo: r[0], date: r[1], skuCode: r[2], skuName: r[3],
      bin: r[4], qty: Number(r[5]) || 0, reason: r[6]
    }));
  } catch(e) { return []; }
}

/**
 * Load all WMS data on app startup — called once after login
 */
function loadAllWMSData() {
  // Returns masters for login initialization — binMaster served from cache when warm.
  try {
    const binRows = _cachedRead_('binMaster', getBinMasterFromSheet, 6);
    const binSummary = binRows.reduce((acc, b) => {
      acc.total++;
      const st = String(b.status || '').trim().toUpperCase();
      if (st === 'EMPTY') acc.empty++;
      else if (st === 'OCCUPIED') acc.occupied++;
      else if (st === 'RESERVED') acc.reserved++;
      else if (st === 'BLOCKED') acc.blocked++;
      return acc;
    }, { total: 0, empty: 0, occupied: 0, reserved: 0, blocked: 0 });

    return {
      success:   true,
      skuMaster: getSKUMasterFromSheet(),
      binSummary
    };
  } catch(e) {
    return { success: false, message: e.message, skuMaster: [], binSummary: null };
  }
}


/**
 * ════════════════════════════════════════════════════════
 * FIXED v5: saveAdjustmentLog — APPENDS rows (never replaces).
 * Each stock transfer logs 2 rows: REMOVE (negative qty) + ADD (positive qty).
 * Each new-stock addition logs 1 row: ADD (positive qty).
 * All entries are permanently stored — old data is never wiped.
 * rows: [{time, bin, skuCode, skuName, batch, mfgDate, expDate,
 *         stockType, qty, action, ref, by}]
 * ════════════════════════════════════════════════════════
 */
