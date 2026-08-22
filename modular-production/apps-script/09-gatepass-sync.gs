/**
 * Modular split from Code.gs
 * File: 09-gatepass-sync.gs
 */

function syncGatepasses() {
  const log = [];
  const t0  = new Date();
  function L(m){ log.push(m); Logger.log(m); }
  L('=== Gatepass Sync: ' + uwFmtTs_(t0) + ' ===');
  try {
    const email = gpGetLatestEmail_();
    if (!email) {
      L('No new Gatepass email found');
      return { success: true, status: 'SKIPPED', message: 'No new email', rowCount: 0 };
    }
    L('Email: ' + uwFmtTs_(email.date) + ' | URL: ' + email.url);
    const rows = gpFetchAndFilter_(email.url, log);
    L('Valid rows after filter: ' + rows.length);
    gpWriteSheet_(rows);
    L('GATEPASS_DUMP written: ' + rows.length + ' rows');
    return { success: true, status: 'SUCCESS', rowCount: rows.length,
             timestamp: uwFmtTs_(t0), message: 'Synced ' + rows.length + ' gatepass rows' };
  } catch(e) {
    L('ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

function forceReimportGatepass() {
  const log = [];
  function L(m){ log.push(m); Logger.log(m); }
  L('=== Force Re-import Gatepass ===');
  try {
    const q       = 'subject:"' + UW_CONFIG.GP_SUBJECT + '" newer_than:7d';
    const threads = GmailApp.search(q, 0, 10);
    L('Threads found: ' + threads.length);
    if (!threads.length) return { success: false, message: 'No Gatepass email found in last 7 days' };

    const candidates = [];
    for (const thread of threads) {
      for (const msg of thread.getMessages()) {
        const plain = msg.getPlainBody();
        const html  = msg.getBody();
        const pm    = plain.match(/(https:\/\/[a-zA-Z0-9\-\.]+\.cloudfront\.net\/[^\s<>"']+\.csv)/i)
                   || html.match(/(https:\/\/[a-zA-Z0-9\-\.]+\.cloudfront\.net\/[^\s<>"']+\.csv)/i);
        if (!pm) continue;
        candidates.push({ url: decodeURIComponent(pm[1].trim()), date: msg.getDate() });
      }
    }
    candidates.sort((a, b) => b.date - a.date);
    if (!candidates.length) return { success: false, message: 'No CSV URL found in any Gatepass email' };

    const email = candidates[0];
    L('Importing from: ' + uwFmtTs_(email.date) + ' | ' + email.url);
    const rows = gpFetchAndFilter_(email.url, log);
    L('Valid rows: ' + rows.length);
    gpWriteSheet_(rows);
    return { success: true, rowCount: rows.length,
             message: 'Force-imported ' + rows.length + ' gatepass rows from ' + uwFmtTs_(email.date) };
  } catch(e) {
    L('ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

/** Find latest unprocessed Gatepass email */
function gpGetLatestEmail_() {
  const q       = 'subject:"' + UW_CONFIG.GP_SUBJECT + '" newer_than:1d';
  const threads = GmailApp.search(q, 0, 10);
  if (!threads.length) return null;

  const processedIds = uwGetProcessedIds_(); // reuse same SYNC_LOG tracker
  const candidates   = [];
  for (const thread of threads) {
    for (const msg of thread.getMessages()) {
      if (processedIds.has(msg.getId())) continue;
      const plain = msg.getPlainBody();
      const html  = msg.getBody();
      const pm    = plain.match(/(https:\/\/[a-zA-Z0-9\-\.]+\.cloudfront\.net\/[^\s<>"']+\.csv)/i)
                 || html.match(/(https:\/\/[a-zA-Z0-9\-\.]+\.cloudfront\.net\/[^\s<>"']+\.csv)/i);
      if (!pm) continue;
      candidates.push({ url: decodeURIComponent(pm[1].trim()), date: msg.getDate(), messageId: msg.getId() });
    }
  }
  if (!candidates.length) return null;
  candidates.sort((a, b) => b.date - a.date);
  return candidates[0];
}

/** Fetch CSV + apply filters: From Party + Status */
function gpFetchAndFilter_(url, log) {
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (resp.getResponseCode() !== 200) throw new Error('HTTP ' + resp.getResponseCode());

  const text    = resp.getContentText('UTF-8').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines   = text.split('\n');
  if (lines.length < 2) return [];

  // Parse using | delimiter (Gatepass CSV uses pipe)
  const rawHeaders = lines[0].split('|').map(h => h.trim().replace(/^ï»¿/, ''));
  if(log) log.push('GP CSV headers: ' + rawHeaders.join(' | '));

  const cm = UW_CONFIG.GP_COL_MAP;
  const hi = {}; // column name → index
  rawHeaders.forEach((h, i) => { hi[h] = i; });

  const colIdx = (name) => hi[name] !== undefined ? hi[name] : -1;

  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    const vals = line.split('|');
    const get  = (name) => {
      const idx = colIdx(name);
      return idx >= 0 ? String(vals[idx] || '').trim() : '';
    };

    // Apply filters
    const fromParty = get(cm.fromParty);
    const gpStatus  = get(cm.gpStatus);
    if (fromParty !== UW_CONFIG.GP_FROM_PARTY) continue;
    if (!UW_CONFIG.GP_STATUSES.includes(gpStatus)) continue;

    const skuCode = get(cm.skuCode);
    const qty     = parseFloat(get(cm.qty).replace(/,/g, '')) || 0;
    if (!skuCode || qty <= 0) continue;

    rows.push([
      get(cm.gpCode),
      gpStatus,
      get(cm.gpCreatedAt),
      get(cm.toParty),
      get(cm.fromParty),
      get(cm.itemName),
      skuCode,
      get(cm.batch),
      qty,
      get(cm.bin),
      UW_CONFIG.STOCK_TYPE_MAP[get(cm.stockType)] || get(cm.stockType) || 'GOOD',
      get(cm.itemStatus),
      uwNormDate_(get(cm.mfgDate)),
      uwNormDate_(get(cm.expDate)),
    ]);
  }
  return rows;
}

/** Write filtered rows to GATEPASS_DUMP sheet */
function gpWriteSheet_(rows) {
  const ss  = SS; // SS not openById
  let sh    = ss.getSheetByName(UW_CONFIG.GP_SHEET) || ss.insertSheet(UW_CONFIG.GP_SHEET);
  sh.clearContents();
  const headers = ['Gatepass Code', 'GP Status', 'Created At', 'To Party', 'From Party',
                   'Item Name', 'SKU Code', 'Batch No', 'Qty', 'Shelf/Bin',
                   'Stock Type', 'Item Status', 'MFG Date', 'EXP Date'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.setTabColor('#f59e0b');
  SpreadsheetApp.flush();
}

/** Return gatepass data to frontend for diff check enrichment */
function getGatepasses() {
  try {
    const ss = SS; // SS not openById
    const sh = ss.getSheetByName(UW_CONFIG.GP_SHEET);
    if (!sh || sh.getLastRow() < 2) return { success: true, rows: [] };
    const data = sh.getDataRange().getDisplayValues().slice(1);
    const rows = data.filter(r => r[0]).map(r => ({
      gpCode:     r[0], gpStatus: r[1], createdAt: r[2],
      toParty:    r[3], fromParty: r[4], itemName: r[5],
      skuCode:    r[6], batch: r[7],
      qty:        Number(r[8]) || 0,
      bin:        r[9], stockType: r[10], itemStatus: r[11],
      mfgDate:    r[12], expDate: r[13],
    }));
    return { success: true, rows, count: rows.length };
  } catch(e) { return { success: false, rows: [], message: e.message }; }
}

function resetProcessedEmails() {
  // Clears all SUCCESS rows from SYNC_LOG so emails can be re-imported.
  // Use this if you need to re-run a past sync (e.g. after a bad import).
  const ss = SpreadsheetApp.openById(UW_CONFIG.SHEET_ID);
  const sh = ss.getSheetByName(UW_CONFIG.SHEET_SYNC_LOG);
  if (!sh || sh.getLastRow() < 2) { Logger.log('SYNC_LOG empty — nothing to reset'); return; }
  const data = sh.getDataRange().getValues();
  // Keep header + non-SUCCESS rows
  const keep = [data[0]].concat(data.slice(1).filter(r => String(r[1]).toUpperCase() !== 'SUCCESS'));
  sh.clearContents();
  sh.getRange(1, 1, keep.length, keep[0].length).setValues(keep);
  Logger.log('✅ Reset: kept ' + (keep.length-1) + ' non-success rows. Emails can now be re-imported.');
}

/* ════════════════════════════════════════════════════════
   PUTAWAY PLAN — server-side functions
   ════════════════════════════════════════════════════════ */

function savePutawayPlan(rows, grnNo, createdBy) {
  try {
    if (!rows || !rows.length) return { success: false, message: 'No plan rows' };
    const sh = SS.getSheetByName('PUTAWAY_PLAN') || SS.insertSheet('PUTAWAY_PLAN');
    const headers = ['Plan ID','GRN No','Line No','Pallet No','SKU Code','SKU Name',
                     'Batch','Bin ID','Alloc Qty','Alloc Boxes','Status','Warning',
                     'Is Full Pallet','SKU Class','Brand','Created By','Created At',
                     'Confirmed By','Confirmed At'];
    if (sh.getLastRow() < 1) {
      sh.getRange(1,1,1,headers.length).setValues([headers])
        .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    // Cancel any existing active plan first
    _cancelActivePlanRows_(sh, grnNo);

    const now = new Date().toISOString();
    const planId = String(grnNo || '').trim();
    const data = rows.map(r => [
      planId, planId,
      Number(r.lineNo || 0), Number(r.palletNo || 1),
      String(r.sku || r.skuCode || ''), String(r.skuName || ''),
      String(r.batch || ''), String(r.binId || 'NO_BIN'),
      Number(r.allocQty || 0), Number(r.allocBoxes || 0),
      'PENDING', r.warning ? 'YES' : 'NO',
      r.isFullPallet ? 'YES' : 'NO',
      String(r.cls || ''), String(r.binBrand || ''),
      String(createdBy || 'WMS App'), now, '', ''
    ]);
    if (data.length) sh.getRange(sh.getLastRow()+1, 1, data.length, headers.length).setValues(data);

    // Mark bins Reserved in BIN_MASTER_WMS
    const binSh = SS.getSheetByName('BIN_MASTER_WMS');
    if (binSh) {
      const binsToReserve = [...new Set(rows.filter(r=>r.binId&&r.binId!=='NO_BIN'&&!r.warning).map(r=>r.binId))];
      _setBinStatuses_(binSh, binsToReserve, 'Reserved', createdBy);
    }

    cleanOldPutawayRows_(sh);
    return { success: true, planId, count: data.length };
  } catch(e) { return { success: false, message: e.message }; }
}

function loadActivePutawayPlan() {
  try {
    const sh = SS.getSheetByName('PUTAWAY_PLAN');
    if (!sh || sh.getLastRow() < 2) return { success: true, rows: [], planId: '' };
    const data = sh.getDataRange().getDisplayValues();
    const rows = [];
    let planId = '';
    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][10] || '').trim().toUpperCase();
      if (status === 'CANCELLED') continue;
      if (!planId && status === 'PENDING') planId = String(data[i][0] || '').trim();
      rows.push({
        planId:       String(data[i][0] || ''),
        grnNo:        String(data[i][1] || ''),
        lineNo:       Number(data[i][2] || 0),
        palletNo:     Number(data[i][3] || 1),
        sku:          String(data[i][4] || ''),
        skuName:      String(data[i][5] || ''),
        batch:        String(data[i][6] || ''),
        binId:        String(data[i][7] || ''),
        allocQty:     Number(data[i][8] || 0),
        allocBoxes:   Number(data[i][9] || 0),
        status:       String(data[i][10] || 'PENDING'),
        warning:      String(data[i][11] || '') === 'YES',
        isFullPallet: String(data[i][12] || '') === 'YES',
        cls:          String(data[i][13] || ''),
        binBrand:     String(data[i][14] || ''),
        createdBy:    String(data[i][15] || ''),
        createdAt:    String(data[i][16] || ''),
        confirmedBy:  String(data[i][17] || ''),
        confirmedAt:  String(data[i][18] || ''),
      });
    }
    // Only return rows for the active plan (most recent PENDING planId)
    const filtered = planId ? rows.filter(r => r.planId === planId) : rows;
    return { success: true, rows: filtered, planId };
  } catch(e) { return { success: false, rows: [], message: e.message }; }
}

function confirmPutawayRow(planId, lineNo, binId, confirmedBy) {
  try {
    planId = String(planId || '').trim();
    lineNo = Number(lineNo || 0);
    const sh = SS.getSheetByName('PUTAWAY_PLAN');
    if (!sh) return { success: false, message: 'PUTAWAY_PLAN sheet missing' };
    const data = sh.getDataRange().getDisplayValues();
    const now = new Date().toISOString();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === planId && Number(data[i][2]) === lineNo) {
        sh.getRange(i+1, 11).setValue('CONFIRMED');
        sh.getRange(i+1, 18).setValue(confirmedBy || '');
        sh.getRange(i+1, 19).setValue(now);
        found = true;
        break;
      }
    }
    if (!found) return { success: false, message: 'Row not found' };
    // Mark bin Occupied
    if (binId && binId !== 'NO_BIN') {
      const binSh = SS.getSheetByName('BIN_MASTER_WMS');
      if (binSh) _setBinStatuses_(binSh, [binId], 'Occupied', confirmedBy);
    }
    // Check if all rows for this plan are confirmed
    const planRows = sh.getDataRange().getDisplayValues().slice(1)
      .filter(r => String(r[0]).trim() === planId);
    const allDone = planRows.every(r => ['CONFIRMED','CANCELLED'].includes(String(r[10]).trim().toUpperCase()));
    return { success: true, allConfirmed: allDone };
  } catch(e) { return { success: false, message: e.message }; }
}

function swapPutawayBin(planId, lineNo, oldBinId, newBinId, updatedBy) {
  try {
    planId = String(planId || '').trim();
    lineNo = Number(lineNo || 0);
    const sh = SS.getSheetByName('PUTAWAY_PLAN');
    if (!sh) return { success: false, message: 'PUTAWAY_PLAN sheet missing' };
    const data = sh.getDataRange().getDisplayValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === planId && Number(data[i][2]) === lineNo) {
        sh.getRange(i+1, 8).setValue(newBinId);
        sh.getRange(i+1, 11).setValue('PENDING');
        sh.getRange(i+1, 12).setValue('NO');
        found = true;
        break;
      }
    }
    if (!found) return { success: false, message: 'Row not found' };
    const binSh = SS.getSheetByName('BIN_MASTER_WMS');
    if (binSh) {
      if (oldBinId && oldBinId !== 'NO_BIN') _setBinStatuses_(binSh, [oldBinId], 'Empty', updatedBy);
      if (newBinId && newBinId !== 'NO_BIN') _setBinStatuses_(binSh, [newBinId], 'Reserved', updatedBy);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function cancelPutawayPlan(planId, releasedBy) {
  try {
    planId = String(planId || '').trim();
    const sh = SS.getSheetByName('PUTAWAY_PLAN');
    if (!sh || sh.getLastRow() < 2) return { success: true };
    const data = sh.getDataRange().getDisplayValues();
    const binsToRelease = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === planId) {
        sh.getRange(i+1, 11).setValue('CANCELLED');
        const bin = String(data[i][7] || '').trim();
        if (bin && bin !== 'NO_BIN') binsToRelease.push(bin);
      }
    }
    const binSh = SS.getSheetByName('BIN_MASTER_WMS');
    if (binSh && binsToRelease.length) _setBinStatuses_(binSh, binsToRelease, 'Empty', releasedBy);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function confirmAllPutaway(planId, confirmedBy) {
  try {
    planId = String(planId || '').trim();
    const sh = SS.getSheetByName('PUTAWAY_PLAN');
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'No plan found' };
    const data = sh.getDataRange().getDisplayValues();
    const now = new Date().toISOString();
    const binsToOccupy = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== planId) continue;
      if (String(data[i][10]).trim().toUpperCase() === 'PENDING') {
        sh.getRange(i+1, 11).setValue('CONFIRMED');
        sh.getRange(i+1, 18).setValue(confirmedBy || '');
        sh.getRange(i+1, 19).setValue(now);
        const bin = String(data[i][7] || '').trim();
        if (bin && bin !== 'NO_BIN') binsToOccupy.push(bin);
      }
    }
    const binSh = SS.getSheetByName('BIN_MASTER_WMS');
    if (binSh && binsToOccupy.length) _setBinStatuses_(binSh, binsToOccupy, 'Occupied', confirmedBy);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateBinStatus(binId, status, updatedBy) {
  try {
    binId = String(binId || '').trim();
    if (!binId) return { success: false, message: 'Bin ID required' };
    const sh = SS.getSheetByName('BIN_MASTER_WMS');
    if (!sh) return { success: false, message: 'BIN_MASTER_WMS missing' };
    _setBinStatuses_(sh, [binId], status, updatedBy);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function addOrUpdateSKURow(row) {
  _invalidateCache_('skuMaster');
  try {
    const code = String(row['SKU Code'] || row.code || '').trim();
    if (!code) return { success: false, message: 'SKU Code required' };
    const sh = SS.getSheetByName('SKU_MASTER') || SS.insertSheet('SKU_MASTER');
    const headers = ['SKU Code','SKU Name','Brand','Case Pack (Units/Box)','Box/Pallet (Boxes/Shelf)',
                     'Classification (A/B/C)','MRP','SKU Status','Updated At','Updated By'];
    if (sh.getLastRow() < 1) {
      sh.getRange(1,1,1,headers.length).setValues([headers])
        .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const data = sh.getDataRange().getDisplayValues();
    const now = new Date().toISOString();
    const newRow = [
      code,
      String(row['SKU Name'] || row.name || ''),
      String(row.Brand || row.brand || inferBrandFromSKUCode_(code)),
      Number(row['Case Pack (Units/Box)'] || row.casePack || '') || '',
      Number(row['Box/Pallet (Boxes/Shelf)'] || row.boxPallet || '') || '',
      String(row['Classification (A/B/C)'] || row.cls || 'B'),
      String(row.mrp || ''),
      String(row.skuStatus || 'ACTIVE'),
      now, 'WMS App'
    ];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === code) {
        sh.getRange(i+1, 1, 1, headers.length).setValues([newRow]);
        return { success: true, action: 'updated' };
      }
    }
    sh.appendRow(newRow);
    return { success: true, action: 'added' };
  } catch(e) { return { success: false, message: e.message }; }
}

function addOrUpdateBinRow(row) {
  _invalidateCache_('binMaster');
  try {
    const binId = normalizeBinIdForMaster_(row['Bin ID'] || row.id || row.binId);
    if (!binId) return { success: false, message: 'Bin ID required' };
    const sh = SS.getSheetByName('BIN_MASTER_WMS') || SS.insertSheet('BIN_MASTER_WMS');
    const headers = ['Bin ID','Row','Column No','Level','Brand Zone',
                     'FSN Class (A/B/C)','Status','SKU in Bin','Updated At','Updated By'];
    if (sh.getLastRow() < 1) {
      sh.getRange(1,1,1,headers.length).setValues([headers])
        .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const parsed = binId.match(/^R(\d+)-C(\d+)-(\d+)$/i);
    const brand = String(row['Brand Zone'] || row.brand || row.zone || inferBrandFromBin_(binId));
    const fsn   = String(row['FSN Class'] || row.fsn || 'B').toUpperCase();
    const status = String(row.Status || row.status || 'Empty');
    const now = new Date().toISOString();
    const newRow = [
      binId,
      row.row || (parsed ? parseInt(parsed[1]) : ''),
      row.col || (parsed ? parseInt(parsed[2]) : ''),
      row.level || (parsed ? parseInt(parsed[3]) : ''),
      brand, fsn, status, '', now, 'WMS App'
    ];
    const data = sh.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (normalizeBinIdForMaster_(data[i][0]) === binId) {
        sh.getRange(i+1, 1, 1, headers.length).setValues([newRow]);
        return { success: true, action: 'updated' };
      }
    }
    sh.appendRow(newRow);
    return { success: true, action: 'added' };
  } catch(e) { return { success: false, message: e.message }; }
}

/* ── Private helpers ── */

function _setBinStatuses_(sh, binIds, status, updatedBy) {
  if (!binIds || !binIds.length) return;
  _invalidateCache_('binMaster'); // invalidate before write so next read is fresh
  const binSet = new Set(binIds.map(b => normalizeBinIdForMaster_(b)).filter(Boolean));
  const data = sh.getDataRange().getDisplayValues();
  const now = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    const id = normalizeBinIdForMaster_(data[i][0]);
    if (binSet.has(id)) {
      sh.getRange(i+1, 7).setValue(status);
      sh.getRange(i+1, 9).setValue(now);
      sh.getRange(i+1, 10).setValue(updatedBy || 'WMS App');
    }
  }
}

function _cancelActivePlanRows_(sh, excludePlanId) {
  if (sh.getLastRow() < 2) return;
  const data = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    const pid = String(data[i][0] || '').trim();
    const st  = String(data[i][10] || '').trim().toUpperCase();
    if (st === 'PENDING' && pid !== String(excludePlanId || '').trim()) {
      sh.getRange(i+1, 11).setValue('CANCELLED');
    }
  }
}

function cleanOldPutawayRows_(sh) {
  if (!sh) sh = SS.getSheetByName('PUTAWAY_PLAN');
  if (!sh || sh.getLastRow() < 2) return;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);
  const data = sh.getDataRange().getDisplayValues();
  // Delete from bottom up to preserve row indices
  for (let i = data.length - 1; i >= 1; i--) {
    const st = String(data[i][10] || '').trim().toUpperCase();
    const createdAt = new Date(data[i][16] || '');
    if (st === 'CONFIRMED' && !isNaN(createdAt.getTime()) && createdAt < cutoff) {
      sh.deleteRow(i + 1);
    }
  }
}
