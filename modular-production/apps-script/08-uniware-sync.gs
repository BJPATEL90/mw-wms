/**
 * Modular split from Code.gs
 * File: 08-uniware-sync.gs
 */

const UW_CONFIG = {
  SHEET_ID:       '1iBQgbFq6D5iKXSLP6NzfdqxkUuGAFHWe_9Wl7sKZ0jA',
  SUBJECT_FILTER: 'Export Job Complete - Shelfwise Inventory',
  SEARCH_DAYS:    1,   // look back 24 hours only — always use latest email
  TZ:             'Asia/Kolkata',
  URL_REGEX:      /(https:\/\/[a-zA-Z0-9\-\.]+\.cloudfront\.net\/[^\s<>"']+\.csv)/i,
  SHEET_DUMP:     'INVENTORY_DUMP',
  SHEET_PREV:     'INVENTORY_DUMP_PREV',
  SHEET_SYNC_LOG: 'SYNC_LOG',
  TRIGGER_MINS:   30,
  STOCK_TYPE_MAP: {
    'GOOD_INVENTORY': 'GOOD',
    'BAD_INVENTORY':  'BAD',
  },
  DUMP_HEADERS: [
    'SKU Code','SKU Name','Batch No','MFG Date','EXP Date',
    'MRP','Bin / Shelf','Stock Type','SKU Status','Qty',
    'EAN','Uploaded At','Uploaded By'
  ],
  COL_MAP: {
    skuCode:      'Item Type SKU Code',
    skuName:      'Item Type Name',
    batch:        'Vendor batch code',
    mfgDate:      'Manufacturing',
    expDate:      'Expiry',
    mrp:          'MRP',
    bin:          'Shelf',
    stockType:    'Inventory Type',
    skuStatus:    'Batch Status',
    qty:          'Quantity',
    ean:          'EAN',
    qtyBlocked:   'Quantity Blocked',
    qtyDamaged:   'Quantity Damaged',
    facility:     'Facility',
    section:      'Section',
  },

  // ── Gatepass sync config ──
  GP_SUBJECT:    'Export Job Complete - Gatepass All Facility',
  GP_SHEET:      'GATEPASS_DUMP',   // sheet where raw gatepass data is written
  GP_FROM_PARTY: 'SL Mother Hub',   // filter: only fetch gatepasses from this facility
  GP_STATUSES:   ['CREATED', 'RETURN_AWAITED'], // filter: only these header statuses
  GP_COL_MAP: {
    gpCode:      'Gatepass Code',
    gpStatus:    'Gatepass Status',
    gpCreatedAt: 'Gatepass Created At',
    toParty:     'To Party',
    fromParty:   'From Party',
    itemName:    'Item Name',
    skuCode:     'Item SkuCode',
    batch:       'Vendor Batch No',
    qty:         'Quantity',
    bin:         'Shelf',
    stockType:   'Inventory Type',
    itemStatus:  'Gatepass Item Status',
    mfgDate:     'Manufacturing date',
    expDate:     'Expiry Date',
  },
};

/* ── Main trigger function ── */
function runUniwareSync() {
  const log = [];
  const t0  = new Date();
  try {
    log.push('=== Uniware Sync: ' + uwFmtTs_(t0) + ' ===');
    const email = uwGetLatestEmail_();
    if (!email) {
      log.push('No new email found — skipping');
      uwWriteLog_('SKIPPED', 0, 'No new email', t0, log);
      return;
    }
    log.push('Email: ' + email.subject + ' | ' + uwFmtTs_(email.date));
    log.push('URL: ' + email.url);

    const csvRows = uwFetchCSV_(email.url);
    if (!csvRows.length) {
      log.push('ERROR: CSV empty');
      uwWriteLog_('FAILED', 0, 'CSV empty', t0, log);
      return;
    }
    log.push('CSV rows: ' + csvRows.length);

    const mapped = uwMapSchema_(csvRows);
    const valid  = mapped.filter(r => r[0] && Number(r[9]) > 0);
    log.push('Valid rows: ' + valid.length);

    if (!valid.length) {
      log.push('ERROR: No valid rows — aborting to protect existing data');
      uwWriteLog_('FAILED', 0, 'No valid rows', t0, log);
      return;
    }

    uwBackup_();
    log.push('Backup done');
    uwWriteDump_(valid);
    log.push('INVENTORY_DUMP replaced: ' + valid.length + ' rows');
    uwMarkProcessed_(email.messageId); // no-op now — SYNC_LOG is the record
    // Store URL + messageId in detail so we never re-import same email
    const detail = email.url + ' | msgid:' + email.messageId;
    uwWriteLog_('SUCCESS', valid.length, detail, t0, log);
    log.push('=== Done in ' + Math.round((new Date()-t0)/1000) + 's ===');
    Logger.log(log.join('\n'));
  } catch(e) {
    log.push('FATAL: ' + e.message);
    uwWriteLog_('ERROR', 0, e.message, t0, log);
    Logger.log(log.join('\n'));
  }
}

/* ── Find latest unprocessed Shelfwise email ── */
function uwGetLatestEmail_() {
  const q = 'subject:"' + UW_CONFIG.SUBJECT_FILTER + '" newer_than:' + UW_CONFIG.SEARCH_DAYS + 'd';
  const threads = GmailApp.search(q, 0, 20);
  if (!threads.length) return null;
  const processedIds = uwGetProcessedIds_(); // load once, check against SYNC_LOG
  Logger.log('Already-processed message IDs: ' + processedIds.size);
  const candidates = [];
  for (const thread of threads) {
    for (const msg of thread.getMessages()) {
      if (uwIsProcessed_(msg, processedIds)) {
        Logger.log('Skipping already-imported: ' + msg.getSubject() + ' (' + msg.getId() + ')');
        continue;
      }
      const html  = msg.getBody();
      const plain = msg.getPlainBody();
      let url = null;
      // Uniware sends URL as plain text in <td> — NOT as href attribute
      // Search plain body first (cleanest), then HTML body
      const plainMatch = plain.match(/(https:\/\/[a-zA-Z0-9\-\.]+\.cloudfront\.net\/[^\s<>"']+\.csv)/i);
      const htmlMatch  = html.match(/(https:\/\/[a-zA-Z0-9\-\.]+\.cloudfront\.net\/[^\s<>"']+\.csv)/i);
      const match = plainMatch || htmlMatch;
      if (match) url = decodeURIComponent(match[1].trim());
      if (!url) continue;
      candidates.push({ url, date: msg.getDate(), subject: msg.getSubject(), messageId: msg.getId() });
    }
  }
  if (!candidates.length) return null;

  // Sort by date DESC — always pick the MOST RECENT unprocessed email
  candidates.sort((a, b) => b.date - a.date);
  const chosen = candidates[0];
  Logger.log('Selected email: ' + chosen.subject + ' | Date: ' + uwFmtTs_(chosen.date));
  Logger.log('URL to fetch: ' + chosen.url);
  Logger.log('Total candidates found: ' + candidates.length + ' (picking most recent)');
  return chosen;
}

/* ── Download + parse CSV ── */
function uwFetchCSV_(url) {
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (resp.getResponseCode() !== 200) throw new Error('HTTP ' + resp.getResponseCode());
  const text  = resp.getContentText('UTF-8').replace(/\r\n/g,'\n').replace(/\r/g,'\n');
  const lines = text.split('\n');
  if (lines.length < 2) return [];
  // Trim + preserve exact case (COL_MAP uses exact case from CSV)
  // Also strip BOM character that some CSV exports prepend
  const headers = uwParseLine_(lines[0]).map(h => h.trim().replace(/^\uFEFF/, ''));
  Logger.log('CSV headers found (' + headers.length + '): ' + headers.join(' | '));
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    const vals = uwParseLine_(line);
    const obj  = {};
    headers.forEach((h, idx) => { obj[h] = (vals[idx] !== undefined ? vals[idx] : '').trim(); });
    rows.push(obj);
  }
  return rows;
}

function uwParseLine_(line) {
  const res = []; let cur = ''; let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') { if (inQ && line[i+1] === '"') { cur += '"'; i++; } else inQ = !inQ; }
    else if (c === ',' && !inQ) { res.push(cur); cur = ''; }
    else cur += c;
  }
  res.push(cur);
  return res;
}

/* ── Map Shelfwise columns → WMS INVENTORY_DUMP schema ── */
function uwMapSchema_(rows) {
  const cm  = UW_CONFIG.COL_MAP;
  const now = Utilities.formatDate(new Date(), UW_CONFIG.TZ, 'dd-MM-yyyy HH:mm:ss');
  return rows.map(r => {
    // Stock type normalisation
    const stRaw     = String(r[cm.stockType] || '').trim();
    const stockType = UW_CONFIG.STOCK_TYPE_MAP[stRaw] || stRaw || 'GOOD';

    // Quantity — use main Quantity column
    // Note: Quantity Blocked and Quantity Damaged are available for future use
    const qty = parseFloat(String(r[cm.qty] || '0').replace(/,/g, '')) || 0;

    // EAN — now properly mapped from CSV
    const ean = String(r[cm.ean] || '').trim();

    return [
      String(r[cm.skuCode]  || '').trim(),   // SKU Code
      String(r[cm.skuName]  || '').trim(),   // SKU Name
      String(r[cm.batch]    || '').trim(),   // Batch No  ← FIXED
      uwNormDate_(r[cm.mfgDate]),            // MFG Date  ← FIXED
      uwNormDate_(r[cm.expDate]),            // EXP Date
      String(r[cm.mrp]      || '').trim(),   // MRP
      normalizeBinIdForMaster_(r[cm.bin]),   // Bin / Shelf
      stockType,                             // Stock Type
      String(r[cm.skuStatus]|| '').trim(),   // SKU Status
      qty,                                   // Qty
      ean,                                   // EAN  ← now populated
      now,                                   // Uploaded At
      'Uniware Auto-Sync',                   // Uploaded By
    ];
  });
}

function uwNormDate_(raw) {
  if (!raw) return '';
  raw = String(raw).trim();
  if (!raw) return '';

  // Already DD-MM-YYYY or DD/MM/YYYY — safe format, return as-is
  if (/^\d{2}[-\/]\d{2}[-\/]\d{4}$/.test(raw)) return raw.replace(/\//g, '-');

  // YYYY-MM-DD (ISO format — what Shelfwise sends e.g. "2026-01-03")
  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) {
    const p = raw.substring(0, 10).split('-');
    // Return DD-MM-YYYY as plain string — Google Sheets won't auto-convert this
    return p[2] + '-' + p[1] + '-' + p[0];
  }

  // Excel serial number (5-digit number)
  if (/^\d{5}(\.\d+)?$/.test(raw)) {
    const d = new Date((parseInt(raw) - 25569) * 86400000);
    return Utilities.formatDate(d, UW_CONFIG.TZ, 'dd-MM-yyyy');
  }

  // JavaScript Date object converted to string (e.g. "Sat Jan 03 2026 00:00:00 GMT+0530")
  // This happens when GAS parses the CSV cell as a Date automatically
  if (/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)/.test(raw)) {
    try {
      const d = new Date(raw);
      if (!isNaN(d.getTime())) {
        return Utilities.formatDate(d, UW_CONFIG.TZ, 'dd-MM-yyyy');
      }
    } catch(e) {}
  }

  // Generic date string — try parsing
  const d = new Date(raw);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, UW_CONFIG.TZ, 'dd-MM-yyyy');
  }

  return raw; // return as-is if nothing matched
}

/* ── Backup existing INVENTORY_DUMP ── */
function uwBackup_() {
  const ss    = SS; // SS not openById
  const srcSh = ss.getSheetByName(UW_CONFIG.SHEET_DUMP);
  if (!srcSh || srcSh.getLastRow() < 2) return;
  let prev = ss.getSheetByName(UW_CONFIG.SHEET_PREV) || ss.insertSheet(UW_CONFIG.SHEET_PREV);
  const data = srcSh.getDataRange().getValues();
  prev.clearContents();
  prev.getRange(1,1,data.length,data[0].length).setValues(data);
  prev.getRange(1,1,1,data[0].length).setBackground('#7f1d1d').setFontColor('#fff').setFontWeight('bold');
  prev.setTabColor('#dc2626');
}

/* ── Write INVENTORY_DUMP (full replace) ── */
function uwWriteDump_(rows) {
  const ss  = SS; // SS not openById
  let sh    = ss.getSheetByName(UW_CONFIG.SHEET_DUMP) || ss.insertSheet(UW_CONFIG.SHEET_DUMP);
  sh.clearContents();
  const h = UW_CONFIG.DUMP_HEADERS;

  // Write header
  sh.getRange(1,1,1,h.length).setValues([h])
    .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
  sh.setFrozenRows(1);
  sh.setTabColor('#dc2626');

  if (!rows.length) { SpreadsheetApp.flush(); return; }

  // Write data rows
  sh.getRange(2, 1, rows.length, h.length).setValues(rows);

  // Force MFG Date (col 4) and EXP Date (col 5) to plain text
  // Prevents Google Sheets auto-converting "03-01-2026" to a date serial
  sh.getRange(2, 4, rows.length, 2).setNumberFormat('@STRING@');

  SpreadsheetApp.flush();

  // ── Update BIN_MASTER_WMS status from this fresh dump ──
  uwUpdateBinStatus_(rows);
}

/**
 * Updates BIN_MASTER_WMS.Status column based on INVENTORY_DUMP content.
 * Logic:
 *   - Bin found in dump with qty > 0  → Occupied
 *   - Bin in master but NOT in dump   → Empty
 *   - Bin in dump but NOT in master   → ignored
 *
 * Runs as a single batch read + single batch write for performance.
 */
function uwUpdateBinStatus_(dumpRows) {
  const binSh = SS.getSheetByName('BIN_MASTER_WMS'); // Use SS not openById
  if (!binSh) {
    Logger.log('BIN_MASTER_WMS not found — skipping bin status update');
    return;
  }

  // ── Build occupied bins map + SKU-in-Bin from dump rows ──
  // dumpRows: 0=SKU,1=Name,2=Batch,3=MFG,4=EXP,5=MRP,6=Bin,7=StockType,8=Status,9=Qty
  const occupiedBins = new Set();
  const binSkuMap    = {};   // bin → first SKU code found
  dumpRows.forEach(r => {
    const bin = normalizeBinIdForMaster_(r[6]);
    const qty = Number(r[9] || 0);
    if (bin && qty > 0) {
      occupiedBins.add(bin);
      if (!binSkuMap[bin]) binSkuMap[bin] = String(r[0] || '').trim();
    }
  });
  Logger.log('Bins to mark Occupied: ' + occupiedBins.size);

  // ── Read all existing BIN_MASTER_WMS rows ──
  const lastRow  = binSh.getLastRow();
  if (lastRow < 2) {
    Logger.log('BIN_MASTER_WMS is empty — skipped inventory-driven bin creation. Upload Bin Master first.');
    return;
  }

  const binData  = binSh.getRange(2, 1, lastRow - 1, 10).getValues(); // 10 cols: BinID,Row,Col,Level,Brand,FSN,Status,SKUinBin,UpdatedAt,UpdatedBy
  const binIndex = {}; // binId → row index in binData (0-based)
  binData.forEach((row, i) => {
    const id = normalizeBinIdForMaster_(row[0]);
    if (id) binIndex[id] = i;
  });

  // ── Update status for each existing bin ──
  const now2 = Utilities.formatDate(new Date(), UW_CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss'Z'");
  const updatedData = binData.map(row => {
    const id  = normalizeBinIdForMaster_(row[0]);
    if (!id) return row;
    if (String(row[6]||'').trim() === 'Blocked') return row; // Never overwrite Blocked
    const newStatus = occupiedBins.has(id) ? 'Occupied' : 'Empty';
    const updated   = [...row];
    // Ensure row has 10 columns: BinID,Row,Col,Level,Brand,FSN,Status,SKUinBin,UpdatedAt,UpdatedBy
    while (updated.length < 10) updated.push('');
    // If Brand Zone not set yet, infer it
    if (!updated[4]) updated[4] = inferBrandFromBin_(id);
    // Parse Row/Col/Level if missing
    if (!updated[1]) {
      const m = id.match(/^R(\d+)-C(\d+)-(\d+)$/i);
      if (m) { updated[1]=parseInt(m[1]); updated[2]=parseInt(m[2]); updated[3]=parseInt(m[3]); }
    }
    updated[6]  = newStatus;                          // col G (idx 6) = Status ✅
    updated[7]  = binSkuMap[id] || updated[7] || '';  // col H (idx 7) = SKU in Bin ✅
    updated[8]  = now2;                               // col I (idx 8) = Updated At ✅
    updated[9]  = 'Uniware Auto-Sync';                // col J (idx 9) = Updated By ✅
    return updated;
  });

  // ── Write all updates in one batch call ──
  binSh.getRange(2, 1, updatedData.length, 10).setValues(updatedData);

  // Auto-create DISABLED — bin master is user-uploaded, never auto-grown
  const existingBinIds = new Set(Object.keys(binIndex));
  const missingBins    = new Set([...occupiedBins].filter(b => !existingBinIds.has(b)));
  if(missingBins.size>0) Logger.log('INFO: '+missingBins.size+' dump bins not in master (ignored): '+[...missingBins].slice(0,3).join(', '));

  SpreadsheetApp.flush();
  Logger.log('BIN_MASTER_WMS updated: ' + updatedData.length + ' existing bins, ' + missingBins.size + ' dump bins ignored');
}

/**
 * Retained as a no-op compatibility stub. Bin Master is user-uploaded only.
 */
function uwCreateMissingBins_(binSh, missingBins, existingBinIds) {
  Logger.log('uwCreateMissingBins_ skipped: Bin Master must not be auto-grown from inventory dump.');
}

/* ── Gmail label helpers ──
   We track processed emails by writing their Message-ID to SYNC_LOG.
   This avoids any conflict with Gmail filters that may already apply
   labels like 'wms-imported' to incoming Shelfwise emails.
   ── */
function uwGetProcessedIds_() {
  try {
    const sh = SS.getSheetByName(UW_CONFIG.SHEET_SYNC_LOG); // Use SS
    if (!sh || sh.getLastRow() < 2) return new Set();
    // Message IDs stored in column 5 (Detail column) as "msgid:XXXXX"
    const vals = sh.getRange(2, 5, sh.getLastRow()-1, 1).getValues().flat();
    const ids  = new Set();
    vals.forEach(v => {
      const m = String(v).match(/msgid:([a-zA-Z0-9]+)/);
      if (m) ids.add(m[1]);
    });
    return ids;
  } catch(e) { return new Set(); }
}

function uwMarkProcessed_(messageId) {
  // Message ID is already stored in SYNC_LOG detail field (see uwWriteLog_)
  // Nothing extra needed — uwWriteLog_ writes "msgid:XXXXX" in detail column
}

function uwIsProcessed_(msg, processedIds) {
  // Check against our own SYNC_LOG record — NOT Gmail labels
  // This avoids false-positive matches from Gmail filters like 'wms-imported'
  return processedIds && processedIds.has(msg.getId());
}

/* ── Sync log ── */
function uwWriteLog_(status, count, detail, t0, log) {
  const ss = SS; // SS not openById
  let sh   = ss.getSheetByName(UW_CONFIG.SHEET_SYNC_LOG);
  if (!sh) {
    sh = ss.insertSheet(UW_CONFIG.SHEET_SYNC_LOG);
    sh.getRange(1,1,1,6).setValues([['Timestamp','Status','Rows','Duration(s)','Detail','Log']])
      .setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setTabColor('#0891b2');
  }
  const dur = Math.round((new Date()-t0)/1000);
  const ts  = Utilities.formatDate(t0, UW_CONFIG.TZ, 'dd-MM-yyyy HH:mm:ss');
  // Append msgid to detail so uwGetProcessedIds_ can find it later
  sh.appendRow([ts, status, count, dur, detail, log.join('\n')]);
  const bg = status==='SUCCESS'?'#f0fdf4':status==='SKIPPED'?'#fefce8':'#fff5f5';
  sh.getRange(sh.getLastRow(),1,1,6).setBackground(bg);
}

/* ── Timestamp helper ── */
function uwFmtTs_(d) {
  return Utilities.formatDate(d||new Date(), UW_CONFIG.TZ, 'dd-MM-yyyy HH:mm:ss');
}

/* ── Public API called by WMS frontend ── */
function getLastSyncStatus() {
  try {
    const sh = SS.getSheetByName(UW_CONFIG.SHEET_SYNC_LOG); // Use SS not openById
    if (!sh || sh.getLastRow() < 2) return { success:true, status:'NEVER', timestamp:'', rowCount:0, syncId:'' };
    const d      = sh.getRange(sh.getLastRow(),1,1,5).getValues()[0];
    const detail = String(d[4]||'');
    // Extract msgid for reliable frontend change detection
    const mid    = detail.match(/msgid:([a-zA-Z0-9]+)/);
    const syncId = mid ? mid[1] : String(d[0]||'');
    return { success:true, timestamp:String(d[0]||''), status:String(d[1]||''), rowCount:Number(d[2]||0), duration:Number(d[3]||0), detail:detail, syncId:syncId };
  } catch(e) { return { success:false, message:e.message }; }
}

function manualUniwareSync() {
  try { runUniwareSync(); return getLastSyncStatus(); }
  catch(e) { return { success:false, message:e.message }; }
}

/**
 * Force re-import from the latest Shelfwise email regardless of
 * whether it was already imported. Use after FLUSH to restore data.
 * Does NOT check SYNC_LOG — always imports the most recent email found.
 */
function forceReimportLatest() {
  const log = [];
  const t0  = new Date();

  // Helper — logs AND writes to Logger immediately so nothing is lost on early return
  function L(msg) { log.push(msg); Logger.log(msg); }

  L('=== Force Re-import started: ' + uwFmtTs_(t0) + ' ===');
  L('Subject filter: ' + UW_CONFIG.SUBJECT_FILTER);

  try {
    // Search last 7 days regardless of processed status
    const q       = 'subject:"' + UW_CONFIG.SUBJECT_FILTER + '" newer_than:7d';
    L('Gmail query: ' + q);
    const threads = GmailApp.search(q, 0, 20);
    L('Threads found: ' + threads.length);
    if (!threads.length) {
      L('❌ No threads found — check subject line and inbox');
      return { success:false, message:'No Shelfwise email found in last 7 days' };
    }

    // Collect ALL messages with CSV URLs — ignore processed check
    const candidates = [];
    for (const thread of threads) {
      const msgs = thread.getMessages();
      L('Thread has ' + msgs.length + ' message(s)');
      for (const msg of msgs) {
        const body = msg.getBody();
        // URL is plain text in <td> — NOT in href attribute
        // Try plain text body first (cleaner), fall back to HTML body
        const plain = msg.getPlainBody();
        const pm = plain.match(/(https:\/\/[a-zA-Z0-9\-\.]+\.cloudfront\.net\/[^\s<>"']+\.csv)/i)
                || body.match(/(https:\/\/[a-zA-Z0-9\-\.]+\.cloudfront\.net\/[^\s<>"']+\.csv)/i);
        L('  Msg: ' + msg.getSubject() + ' | Date: ' + uwFmtTs_(msg.getDate()) + ' | Has URL: ' + !!pm);
        if (!pm) continue;
        const url = decodeURIComponent(pm[1].trim());
        L('  URL: ' + url);
        candidates.push({ url, date: msg.getDate(), messageId: msg.getId() });
      }
    }

    L('Total candidates with URL: ' + candidates.length);
    if (!candidates.length) {
      L('❌ No CSV URL found in any email body');
      return { success:false, message:'No CSV URL found in any recent email' };
    }

    candidates.sort((a,b) => b.date - a.date);
    const email = candidates[0];
    L('Selected (most recent): ' + uwFmtTs_(email.date));
    L('URL: ' + email.url);

    L('Fetching CSV...');
    const csvRows = uwFetchCSV_(email.url);
    L('CSV rows downloaded: ' + csvRows.length);
    if (!csvRows.length) {
      L('❌ CSV empty — URL may have expired');
      return { success:false, message:'CSV downloaded but empty — URL may have expired' };
    }

    if (csvRows.length > 0) {
      L('CSV columns: ' + Object.keys(csvRows[0]).join(' | '));
      L('Sample qty value: ' + csvRows[0][UW_CONFIG.COL_MAP.qty]);
      L('Sample SKU value: ' + csvRows[0][UW_CONFIG.COL_MAP.skuCode]);
    }

    const mapped = uwMapSchema_(csvRows);
    L('Mapped rows: ' + mapped.length);

    const valid = mapped.filter(r => r[0] && Number(r[9]) > 0);
    L('Valid rows (has SKU + qty > 0): ' + valid.length);

    if (!valid.length && mapped.length > 0) {
      L('DEBUG — first mapped row: ' + JSON.stringify(mapped[0]));
      L('❌ 0 valid rows — check COL_MAP matches actual CSV column names');
      return { success:false, message:'0 valid rows after mapping. See execution log for details.' };
    }

    L('Backing up existing dump...');
    uwBackup_();
    L('Writing ' + valid.length + ' rows to INVENTORY_DUMP...');
    uwWriteDump_(valid);
    L('✅ INVENTORY_DUMP replaced: ' + valid.length + ' rows');

    const detail = email.url + ' | FORCE-REIMPORT | ' + uwFmtTs_(email.date);
    uwWriteLog_('SUCCESS', valid.length, detail, t0, log);
    return { success:true, message:'Re-imported ' + valid.length + ' rows from ' + uwFmtTs_(email.date), rowCount: valid.length };

  } catch(e) {
    L('❌ EXCEPTION: ' + e.message);
    L('Stack: ' + (e.stack || 'none'));
    uwWriteLog_('ERROR', 0, 'forceReimportLatest: ' + e.message, t0, log);
    return { success:false, message:e.message };
  }
}

function restorePreviousDump() {
  try {
    const ss   = SS; // SS not openById
    const prev = ss.getSheetByName(UW_CONFIG.SHEET_PREV);
    if (!prev || prev.getLastRow() < 2) return { success:false, message:'No backup found in '+UW_CONFIG.SHEET_PREV };
    const data = prev.getDataRange().getValues();
    let sh = ss.getSheetByName(UW_CONFIG.SHEET_DUMP) || ss.insertSheet(UW_CONFIG.SHEET_DUMP);
    sh.clearContents();
    sh.getRange(1,1,data.length,data[0].length).setValues(data);
    SpreadsheetApp.flush();
    return { success:true, message:'Restored '+(data.length-1)+' rows from backup' };
  } catch(e) { return { success:false, message:e.message }; }
}

/* ── ONE-TIME SETUP — run from Apps Script editor ── */
function setupUniwareSync() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runUniwareSync')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runUniwareSync').timeBased().everyMinutes(UW_CONFIG.TRIGGER_MINS).create();
  // Note: Gmail label tracking removed — processed emails tracked via SYNC_LOG instead
  const ss = SpreadsheetApp.openById(UW_CONFIG.SHEET_ID);
  Logger.log('✅ Connected to sheet: ' + ss.getName());
  Logger.log('✅ Trigger: runUniwareSync every ' + UW_CONFIG.TRIGGER_MINS + ' minutes');
  Logger.log('✅ Gmail label created: WMS/Imported');
  Logger.log('Run testUniwareSyncNow() to test with a real email.');
}

function testUniwareSyncNow() {
  Logger.log('Running sync...');
  runUniwareSync();
  const status = getLastSyncStatus();
  Logger.log('Result: ' + status.status + ' | Rows: ' + status.rowCount);
  if (status.status === 'SKIPPED') {
    Logger.log('Sync skipped (no new email). If you need to re-import, run forceReimportLatest()');
  }
}

/** Run after testUniwareSyncNow — shows sample rows to verify column mapping */
function verifyMapping() {
  const ss = SpreadsheetApp.openById(UW_CONFIG.SHEET_ID);
  const sh = ss.getSheetByName(UW_CONFIG.SHEET_DUMP);
  if (!sh || sh.getLastRow() < 2) { Logger.log('INVENTORY_DUMP empty'); return; }

  const headers = sh.getRange(1, 1, 1, 13).getValues()[0];
  const sample  = sh.getRange(2, 1, 3, 13).getValues();

  Logger.log('=== INVENTORY_DUMP Column Check ===');
  Logger.log('Headers: ' + JSON.stringify(headers));
  sample.forEach((row, i) => {
    Logger.log('Row ' + (i+2) + ':');
    Logger.log('  SKU Code : ' + row[0]);
    Logger.log('  SKU Name : ' + row[1]);
    Logger.log('  Batch No : ' + row[2]);   // was blank — should now have value
    Logger.log('  MFG Date : ' + row[3]);   // was blank — should now have value
    Logger.log('  EXP Date : ' + row[4]);
    Logger.log('  MRP      : ' + row[5]);
    Logger.log('  Bin/Shelf: ' + row[6]);
    Logger.log('  StockType: ' + row[7]);
    Logger.log('  SKU Status: ' + row[8]);
    Logger.log('  Qty      : ' + row[9]);
    Logger.log('  EAN      : ' + row[10]);  // was blank — should now have value
  });
}

function debugUniwareEmail() {
  // Searches last 7 days, ignores processed check — shows what WOULD be imported
  Logger.log('Searching Gmail for: ' + UW_CONFIG.SUBJECT_FILTER);
  const q       = 'subject:"' + UW_CONFIG.SUBJECT_FILTER + '" newer_than:7d';
  const threads = GmailApp.search(q, 0, 20);
  Logger.log('Threads found: ' + threads.length);

  if (!threads.length) {
    Logger.log('❌ No emails found. Check subject line matches exactly.');
    return;
  }

  const processedIds = uwGetProcessedIds_();
  Logger.log('Already-processed IDs in SYNC_LOG: ' + processedIds.size);

  // Show ALL emails found, processed or not
  const candidates = [];
  for (const thread of threads) {
    for (const msg of thread.getMessages()) {
      const html  = msg.getBody();
      const plain2 = msg.getPlainBody();
      const hm2    = plain2.match(/(https:\/\/[a-zA-Z0-9\-\.]+\.cloudfront\.net\/[^\s<>"']+\.csv)/i)
                  || html.match(/(https:\/\/[a-zA-Z0-9\-\.]+\.cloudfront\.net\/[^\s<>"']+\.csv)/i);
      const url   = hm2 ? decodeURIComponent(hm2[1].trim()) : null;
      const isProc = processedIds.has(msg.getId());
      candidates.push({
        date:      msg.getDate(),
        subject:   msg.getSubject(),
        messageId: msg.getId(),
        hasUrl:    !!url,
        url:       url || '(no CSV URL found)',
        processed: isProc
      });
    }
  }
  candidates.sort((a,b) => b.date - a.date);
  Logger.log('\n=== All matching emails (newest first) ===');
  candidates.forEach((c,i) => {
    Logger.log('[' + (i+1) + '] ' + uwFmtTs_(c.date) +
      ' | Processed: ' + c.processed +
      ' | Has URL: ' + c.hasUrl +
      ' | ID: ' + c.messageId);
    if (c.hasUrl) Logger.log('    URL: ' + c.url);
  });

  const unproc = candidates.filter(c => !c.processed && c.hasUrl);
  Logger.log('\nUnprocessed emails with URL: ' + unproc.length);
  if (unproc.length) {
    Logger.log('Next to import: ' + uwFmtTs_(unproc[0].date));
  } else {
    Logger.log('All emails already imported. Run forceReimportLatest() to re-import the latest one.');
  }
}

/**
 * ONE-TIME DIAGNOSTIC — run once, copy output, then delete.
 * Dumps the raw HTML body of the latest Shelfwise email so we can
 * see exactly how the URL is encoded in the email.
 */
function dumpEmailBody() {
  const q       = 'subject:"Export Job Complete - Shelfwise Inventory" newer_than:7d';
  const threads = GmailApp.search(q, 0, 1);
  if (!threads.length) { Logger.log('No email found'); return; }

  const msg  = threads[0].getMessages()[threads[0].getMessages().length - 1]; // latest msg
  Logger.log('Subject: ' + msg.getSubject());
  Logger.log('Date: '    + msg.getDate());
  Logger.log('From: '    + msg.getFrom());
  Logger.log('');

  const html  = msg.getBody();
  const plain = msg.getPlainBody();

  // Search for cloudfront anywhere in the body
  const cfIdx = html.toLowerCase().indexOf('cloudfront');
  if (cfIdx >= 0) {
    Logger.log('✅ cloudfront found in HTML at pos ' + cfIdx);
    Logger.log('Context (200 chars around it):');
    Logger.log(html.substring(Math.max(0, cfIdx - 100), cfIdx + 200));
  } else {
    Logger.log('❌ cloudfront NOT found in HTML body');
  }

  const cfIdxP = plain.toLowerCase().indexOf('cloudfront');
  if (cfIdxP >= 0) {
    Logger.log('✅ cloudfront found in PLAIN TEXT at pos ' + cfIdxP);
    Logger.log('Context:');
    Logger.log(plain.substring(Math.max(0, cfIdxP - 50), cfIdxP + 300));
  } else {
    Logger.log('❌ cloudfront NOT found in plain text body');
  }

  // Also dump first 500 chars of HTML to see structure
  Logger.log('');
  Logger.log('=== First 500 chars of HTML body ===');
  Logger.log(html.substring(0, 500));
  Logger.log('');
  Logger.log('=== First 500 chars of PLAIN TEXT body ===');
  Logger.log(plain.substring(0, 500));
}


/* ════════════════════════════════════════════════════════
   GATEPASS AUTO-SYNC MODULE
   Fetches "Export Job Complete - Gatepass All Facility" from Gmail
   Filters: From Party = "SL Mother Hub"
            Status = CREATED or RETURN_AWAITED
   Writes to GATEPASS_DUMP sheet
   ════════════════════════════════════════════════════════ */
