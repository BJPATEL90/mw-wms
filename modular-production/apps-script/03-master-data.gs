/**
 * Modular split from production-ready/Code.gs
 * File: 03-master-data.gs
 */

function getNextSr_(sheetName, colNo) {
  const sh = SS.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return 1;
  const vals = sh.getRange(2, colNo, sh.getLastRow() - 1, 1).getValues().flat();
  const nums  = vals.map(Number).filter(n => !isNaN(n));
  return nums.length ? Math.max(...nums) + 1 : 1;
}


/* ════════════════════════════════════════════════════════
   DATA SYNC FUNCTIONS
   Called by google.script.run from the HTML web app.
   Every function returns {success, message, count}.
   ════════════════════════════════════════════════════════ */

/**
 * Bulk-save SKU Master from HTML app.
 * Clears the sheet and rewrites all rows.
 * rows: [{skuCode, skuName, casePack, boxPallet, cls, mrp, skuStatus}]
 */
function bulkSaveSKUMaster(rows) {
  _invalidateCache_('skuMaster');
  try {
    if (!rows || !rows.length) return { success: false, message: 'No SKU rows received' };
    const sh = SS.getSheetByName('SKU_MASTER') || SS.insertSheet('SKU_MASTER');
    const headers = ['SKU Code', 'SKU Name', 'Brand', 'Case Pack (Units/Box)', 'Box/Pallet (Boxes/Shelf)',
                     'Classification (A/B/C)', 'MRP', 'SKU Status', 'Updated At', 'Updated By'];
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');

    const now = new Date().toISOString();
    const data = rows.map(r => {
      const code  = String(r['SKU Code'] || r.skuCode || '').trim();
      const brand = r['Brand'] || r.brand || inferBrandFromSKUCode_(code);
      return [
        code,
        String(r['SKU Name'] || r.skuName || '').trim(),
        brand,
        Number(r['Case Pack (Units/Box)'] || r.casePack || '') || '',
        Number(r['Box/Pallet (Boxes/Shelf)'] || r.boxPallet || '') || '',
        String(r['Classification (A/B/C)'] || r.cls || 'B').trim(),
        String(r.mrp || r.MRP || '').trim(),
        String(r.skuStatus || r['SKU Status'] || 'ACTIVE').trim(),
        now,
        'WMS App'
      ];
    }).filter(r => r[0]);

    if (data.length) sh.getRange(2, 1, data.length, headers.length).setValues(data);
    sh.setFrozenRows(1);
    return { success: true, message: 'SKU Master saved', count: data.length };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Bulk-save Bin Master from HTML app.
 * rows: [{id/binId, zone, fsn/fsnClass, capacity, status, skuFromDump/skuInBin}]
 */
function bulkSaveBinMaster(rows) {
  _invalidateCache_('binMaster');
  try {
    if (!rows || !rows.length) return { success: false, message: 'No Bin rows received' };
    if (rows.length > 10000) {
      return {
        success: false,
        message: 'Bin Master save blocked: received ' + rows.length + ' rows. Upload the clean Bin Master file again; inventory/browser cache is not allowed to rewrite Bin Master.'
      };
    }
    const sh = SS.getSheetByName('BIN_MASTER_WMS') || SS.insertSheet('BIN_MASTER_WMS');
    const headers = ['Bin ID', 'Row', 'Column No', 'Level', 'Brand Zone',
                     'FSN Class (A/B/C)', 'Status', 'SKU in Bin', 'Updated At', 'Updated By'];
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');

    const now = new Date().toISOString();
    const dataByBin = {};
    rows.forEach(r => {
      const binId = normalizeBinIdForMaster_(r['Bin ID'] || r.id || r.binId);
      if (!binId) return;
      // Parse Row/Col/Level from Bin ID if not provided separately
      const parsed = binId.match(/^R(\d+)-C(\d+)-(\d+)$/i);
      const rowNo  = r['Row']       || (parsed ? parseInt(parsed[1]) : '');
      const colNo  = r['Column No'] || (parsed ? parseInt(parsed[2]) : '');
      const level  = r['Level']     || (parsed ? parseInt(parsed[3]) : '');
      // Auto-assign brand zone if not provided
      const brand  = r['Brand Zone'] || r.brand || inferBrandFromBin_(binId);
      const fsn    = String(r['FSN Class'] || r.fsn || 'B').trim().toUpperCase();
      const status = String(r.Status || r.status || 'Empty').trim();
      dataByBin[binId] = [binId, rowNo, colNo, level, brand, fsn, status, '', now, 'WMS App'];
    });
    const data = Object.values(dataByBin);

    if (data.length) sh.getRange(2, 1, data.length, headers.length).setValues(data);
    sh.setFrozenRows(1);
    sh.setTabColor('#4f46e5');

    return { success: true, message: 'Bin Master saved (' + data.length + ' bins)', count: data.length };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/** Infer brand from SKU code prefix */
function inferBrandFromSKUCode_(code) {
  code = String(code || '').trim().toUpperCase();
  if (code.startsWith('MWBW')) return 'Bodywise';
  if (code.startsWith('MWLJ')) return 'LittleJoys';
  if (code.startsWith('MWMM')) return 'ManMatters';
  if (code.startsWith('OWN'))  return 'OWN';
  return '';
}

/** Auto-infer brand zone from Bin ID using BRAND_ZONES map */
function inferBrandFromBin_(binId) {
  const m = String(binId).match(/^R(\d+)-C(\d+)/i);
  if (!m) return 'Unknown';
  const row = parseInt(m[1]), col = parseInt(m[2]);
  for (const [brand, zones] of Object.entries(BRAND_ZONES)) {
    if (zones.some(z => row >= z.rowMin && row <= z.rowMax &&
                        col >= z.colMin && col <= z.colMax)) return brand;
  }
  if (row >= 27 && row <= 35) return 'Overflow';
  return 'Unknown';
}

/** Updates BIN_MASTER Status column from WMS Bin data */
