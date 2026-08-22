/**
 * Modular split from production-ready/Code.gs
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
