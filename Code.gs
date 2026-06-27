/**
 * ══════════════════════════════════════════════════════════════════
 *  MOSAIC WELLNESS WMS — Code.gs  v7 (Enhanced)
 *  Google Apps Script Web App + Spreadsheet Menu
 *
 *  v5 CHANGES:
 *  ① NEW: getInventoryByBin(bin) — live shelf scan with MFG/EXP dates
 *  ② NEW: getBatchDates(skuCode,batch) — auto-fill MFG/EXP from existing batch
 *  ③ FIX: saveAdjustmentLog now APPENDS (not full-replace) — fixes qty tracking
 *  ④ FIX: doPost switch has new cases for getInventoryByBin, getBatchDates
 *
 *  v6 CHANGES:
 *  ⑤ NEW: CURRENT_INVENTORY sheet — always-current aggregated stock view
 *  ⑥ NEW: syncCurrentInventory() — rebuilds Current_Inventory from INVENTORY_DUMP
 *  ⑦ FIX: getInventoryByBin now also returns shelf-specific GP block qty
 *  ⑧ FIX: Movement button duplicate prevention via serverside idempotency key
 *  ⑨ ADD: doPost case for syncCurrentInventory, getGPBlockByBin
 *
 *  v7 CHANGES:
 *  ⑩ NEW: clearSheetData(sheetName) — flush any sheet to headers-only
 *  ⑪ NEW: refreshAllData() — reload SKU/Bin/Gatepass/Inventory in one call
 *  ⑫ FIX: saveCycleCountLog uses Utilities.formatDate for DD-MM-YYYY
 *  ⑬ FIX: saveAdjustmentLog uses Utilities.formatDate for DD-MM-YYYY
 *  ⑭ FIX: saveInventoryDump normalises date strings from Excel serial numbers
 *  ⑮ FIX: cycle count saveCountLines calls syncCurrentInventory after save
 *  ⑯ FIX: getInventoryByBin resolves Excel serial-number dates to DD-MM-YYYY
 *
 *  SETUP (do this ONCE before using the HTML web app):
 *  1. Paste this file → Save → Run → onOpen (grant permissions)
 *  2. WMS Menu → Setup → Diagnose Sheets   ← see what's missing
 *  3. WMS Menu → Setup → Create All Sheets ← creates every sheet
 *  4. WMS Menu → Setup → Add Demo Users    ← adds login accounts
 *  5. Deploy as Web App:
 *       Apps Script → Deploy → New Deployment → Web App
 *       Execute as: Me | Who has access: Anyone
 *       Copy the web app URL and open it in browser
 *
 *  REQUIRED SHEETS (all auto-created by "Create All Sheets"):
 *    USERS | BIN_MASTER | PRODUCT_MASTER | LIVE_STOCK
 *    TO_HEADER | TO_DETAILS | UNLOADING_HEADER | UNLOADING_DETAILS
 *    LOADING_HEADER | LOADING_DETAILS | LOADING_ALLOC
 * ══════════════════════════════════════════════════════════════════
 */

const SS = SpreadsheetApp.openById('1iBQgbFq6D5iKXSLP6NzfdqxkUuGAFHWe_9Wl7sKZ0jA');

/* ════════════════════════════════════════════════════════
   WEB APP ENTRY POINT
   ════════════════════════════════════════════════════════ */
/** Top-level ping — required for google.script.run.ping() */
function ping() { return { success: true, message: 'pong' }; }

function doGet(e) {
  // Handle ping from testGASConnection()
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'ping') {
    return ContentService.createTextOutput(JSON.stringify({ success: true, message: 'pong' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // Do not run ensureAllSheets_() on every page load. It performs many
  // Spreadsheet calls and can make the web app appear frozen.
  // Run setup from the sheet menu, or let write/read actions validate sheets.

  // ?page=mobile → serve mobile floor app
  const page = (e && e.parameter && e.parameter.page) || '';
  if (page === 'mobile') {
    return HtmlService.createHtmlOutputFromFile('mobile')
      .setTitle('MW WMS — Floor App')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Mosaic Wellness WMS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ════════════════════════════════════════════════════════
   doPost — receives fetch() calls from the HTML app
   The HTML's callSheet() sends:
     POST body: JSON { action: "functionName", payload: {...} }
   ════════════════════════════════════════════════════════ */
function doPost(e) {
  const cors = ContentService.createTextOutput()
    .setMimeType(ContentService.MimeType.JSON);

  try {
    const body    = JSON.parse(e.postData.contents || '{}');
    const action  = String(body.action  || '').trim();
    const payload = body.payload;

    let result;

    switch (action) {
      case 'login':               result = login(payload.userId, payload.password);    break;
      case 'ping':               result = { success: true, message: 'pong' };           break;
      case 'bulkSaveSKUMaster':   result = bulkSaveSKUMaster(payload);               break;
      case 'bulkSaveBinMaster':   result = bulkSaveBinMaster(payload);               break;
      case 'saveInventoryDump':   result = saveInventoryDump(payload.rows, payload.mode); break;
      case 'saveGatepasses':      result = saveGatepasses(payload);                  break;
      case 'clearGatepasses':     result = clearGatepasses();                        break;
      case 'saveCycleCountLog':   result = saveCycleCountLog(payload);               break;
      case 'saveGRNToSheet':      result = saveGRNToSheet(payload);                  break;
      case 'loadAllWMSData':      result = loadAllWMSData();                         break;
      case 'getDashboardCounts':  result = getDashboardCounts();                     break;
      case 'ping':                result = { success: true, message: 'pong' };       break;
      case 'saveAdjustmentLog':   result = saveAdjustmentLog(payload);               break;
      // ── NEW in v5 ──
      case 'getInventoryByBin':   result = getInventoryByBin(payload.bin);           break;
      case 'getBatchDates':       result = getBatchDates(payload.skuCode, payload.batch); break;
      // ── NEW in v6 ──
      case 'syncCurrentInventory': result = syncCurrentInventory(payload);            break;
      case 'getGPBlockByBin':     result = getGPBlockByBin(payload.bin);             break;
      case 'saveAdjLogIdempotent': result = saveAdjustmentLogIdempotent(payload.rows, payload.key); break;
      // ── NEW in v7 ──
      case 'clearSheetData':      result = clearSheetData(payload.sheetName);          break;
      case 'refreshAllData':      result = refreshAllData();                            break;
      case 'getGPBlockByBin':     result = getGPBlockByBin(payload.bin);             break;
      // ── Uniware Auto-Sync ──
      case 'getLastSyncStatus':   result = getLastSyncStatus();                        break;
      case 'manualUniwareSync':   result = manualUniwareSync();                        break;
      case 'restorePreviousDump':   result = restorePreviousDump();                    break;
      case 'forceReimportLatest':   result = forceReimportLatest();                    break;
      case 'syncGatepasses':        result = syncGatepasses();                          break;
      case 'forceReimportGatepass': result = forceReimportGatepass();                   break;
      case 'getGatepasses':         result = getGatepasses();                            break;
      case 'getOverflowOptions':  result = getOverflowOptions();                       break;
      case 'savePutawayPlan':      result = savePutawayPlan(payload.rows, payload.grnNo, payload.createdBy); break;
      case 'loadActivePutawayPlan': result = loadActivePutawayPlan(); break;
      case 'confirmPutawayRow':    result = confirmPutawayRow(payload.planId, payload.lineNo, payload.binId, payload.confirmedBy); break;
      case 'swapPutawayBin':       result = swapPutawayBin(payload.planId, payload.lineNo, payload.oldBinId, payload.newBinId, payload.updatedBy); break;
      case 'cancelPutawayPlan':    result = cancelPutawayPlan(payload.planId, payload.releasedBy); break;
      case 'confirmAllPutaway':    result = confirmAllPutaway(payload.planId, payload.confirmedBy); break;
      case 'updateBinStatus':      result = updateBinStatus(payload.binId, payload.status, payload.updatedBy); break;
      case 'addOrUpdateSKURow':    result = addOrUpdateSKURow(payload); break;
      case 'addOrUpdateBinRow':    result = addOrUpdateBinRow(payload); break;
      default:
        result = { success: false, message: 'Unknown action: ' + action };
    }

    cors.setContent(JSON.stringify(result || { success: true }));

  } catch (err) {
    cors.setContent(JSON.stringify({ success: false, message: err.message }));
  }

  return cors;
}

/* doGet also handles ping and simple GETs from the HTML */


/* ════════════════════════════════════════════════════════
   SPREADSHEET MENU (runs when sheet is opened)
   ════════════════════════════════════════════════════════ */
function onOpen() {
  // getUi() only works when triggered by the spreadsheet opening.
  // It will throw if run manually from the Apps Script editor — that is normal.
  try {
    const ui = SpreadsheetApp.getUi();

    const setupMenu = ui.createMenu('Setup')
      .addItem('Diagnose Sheets',       'diagnoseSheetsSetup')
      .addSeparator()
      .addItem('Create All Sheets',     'createAllSheets')
      .addItem('Add Demo Users',        'addDemoUsers')
      .addItem('Add Sample Bins',       'addSampleBins')
      .addItem('Add Sample Products',   'addSampleProducts');

    const reportsMenu = ui.createMenu('Reports')
      .addItem('Refresh Dashboard',     'refreshDashboard')
      .addItem('Stock Report',          'openStockReport');

    ui.createMenu('WMS')
      .addSubMenu(setupMenu)
      .addSeparator()
      .addSubMenu(reportsMenu)
      .addSeparator()
      .addItem('Diagnose Sheets',       'diagnoseSheetsSetup')
      .addToUi();

  } catch (e) {
    // Running from editor — menu cannot be built here.
    // To trigger the menu: open your Google Spreadsheet and reload the page.
    console.log('onOpen: ' + e.message);
  }
}

// ── Run THIS function manually from the editor instead of onOpen ──
// It creates all sheets without needing the spreadsheet UI context.
function setupWMS() {
  // Safe to run directly from the Apps Script editor.
  // Does NOT use getUi() — uses Logger.log() instead.
  ensureAllSheets_();
  Logger.log('✅ All sheets created.');
  addDemoUsers();
  addSampleBins();
  addSampleProducts();
  Logger.log('✅ WMS setup complete!');
  Logger.log('Login credentials: admin/admin123 | wms1/wms123 | manager/mgr123');
  Logger.log('Now open your spreadsheet and reload — the WMS menu will appear.');
}

/* ════════════════════════════════════════════════════════
   SHEET DEFINITIONS — single source of truth
   All sheet names and their column headers live here.
   ════════════════════════════════════════════════════════ */
function getSheetDefs_() {
  // Only sheets actively used by the current WMS + Uniware pipeline.
  // Legacy inbound/outbound sheets (BIN_MASTER, PRODUCT_MASTER, LIVE_STOCK,
  // TO_*, LOADING_*, UNLOADING_*) are retained in the spreadsheet if they exist
  // but are no longer auto-created or managed here.
  return {
    USERS: {
      headers: ['User ID', 'Password', 'Name', 'Role', 'Active (YES/NO)', 'Last Login', 'Allowed Sections'],
      tabColor: '#0d9488'
    },
    // ── WMS App sheets ──
    SKU_MASTER: {
      headers: ['SKU Code', 'SKU Name', 'Brand', 'Case Pack (Units/Box)', 'Box/Pallet (Boxes/Shelf)',
                'Classification (A/B/C)', 'MRP', 'SKU Status', 'Updated At', 'Updated By'],
      tabColor: '#059669'
    },
    BIN_MASTER_WMS: {
      headers: ['Bin ID', 'Row', 'Column No', 'Level', 'Brand Zone',
                'FSN Class (A/B/C)', 'Status', 'SKU in Bin', 'Updated At', 'Updated By'],
      tabColor: '#4f46e5'
    },
    INVENTORY_DUMP: {
      headers: ['SKU Code', 'SKU Name', 'Batch No', 'MFG Date', 'EXP Date', 'MRP',
                'Bin / Shelf', 'Stock Type', 'SKU Status', 'Qty', 'EAN',
                'Uploaded At', 'Uploaded By'],
      tabColor: '#dc2626'
    },
    GATEPASS: {
      headers: ['Gatepass No', 'Date', 'SKU Code', 'SKU Name', 'Bin / Shelf',
                'Qty', 'Reason', 'Uploaded At', 'Uploaded By'],
      tabColor: '#ea580c'
    },
    CYCLE_COUNT_LOG: {
      headers: ['Date', 'Time', 'Bin / Shelf', 'SKU Code', 'SKU Name', 'Batch No',
                'MFG Date', 'EXP Date', 'Stock Type', 'System Qty', 'GP Block Qty',
                'Counted Qty', 'Difference', 'Status', 'Remarks', 'Counted By'],
      tabColor: '#0d9488'
    },
    GRN_LOG: {
      headers: ['GRN No', 'Date', 'Supplier', 'Vehicle No', 'SKU Code', 'SKU Name',
                'Batch No', 'Qty', 'Case Pack', 'Box/Pallet', 'Classification',
                'Saved By', 'Saved At'],
      tabColor: '#d97706'
    },
    ADJUSTMENT_LOG: {
      headers: ['Date', 'Time', 'Bin / Shelf', 'SKU Code', 'SKU Name', 'Batch No',
                'MFG Date', 'EXP Date', 'Stock Type', 'Qty Moved', 'Action',
                'Reference', 'Saved At', 'By'],
      tabColor: '#7c3aed'
    },
    // ⑤ CURRENT_INVENTORY — always-current aggregated view of warehouse stock
    CURRENT_INVENTORY: {
      headers: ['Bin / Shelf', 'SKU Code', 'SKU Name', 'Batch No', 'MFG Date',
                'EXP Date', 'MRP', 'Stock Type', 'SKU Status', 'System Qty',
                'GP Block Qty', 'Last Updated'],
      tabColor: '#0891b2'
    },
    // Uniware sync support sheets (auto-created by sync module)
    SYNC_LOG: {
      headers: ['Timestamp', 'Status', 'Rows', 'Duration(s)', 'Detail', 'Log'],
      tabColor: '#0891b2'
    },
    INVENTORY_DUMP_PREV: {
      headers: ['SKU Code', 'SKU Name', 'Batch No', 'MFG Date', 'EXP Date', 'MRP',
                'Bin / Shelf', 'Stock Type', 'SKU Status', 'Qty', 'EAN',
                'Uploaded At', 'Uploaded By'],
      tabColor: '#dc2626'
    },
    PUTAWAY_PLAN: {
      headers: ['Plan ID','GRN No','Line No','Pallet No','SKU Code','SKU Name',
                'Batch','Bin ID','Alloc Qty','Alloc Boxes','Status','Warning',
                'Is Full Pallet','SKU Class','Brand','Created By','Created At',
                'Confirmed By','Confirmed At'],
      tabColor: '#7c3aed'
    },
  };
}

/* ════════════════════════════════════════════════════════
   SETUP FUNCTIONS
   ════════════════════════════════════════════════════════ */

/** Creates every required sheet if it doesn't exist, and writes headers. */
function createAllSheets() {
  ensureAllSheets_();
  Logger.log('✅ All sheets created! Next: run addDemoUsers(), addSampleBins(), addSampleProducts()');
}

function ensureAllSheets_() {
  const defs = getSheetDefs_();
  Object.entries(defs).forEach(([name, cfg]) => {
    let sh = SS.getSheetByName(name);
    if (!sh) {
      sh = SS.insertSheet(name);
    }
    // Always write header row to ensure columns are correct
    const headerRange = sh.getRange(1, 1, 1, cfg.headers.length);
    headerRange.setValues([cfg.headers]);
    headerRange.setBackground('#1e293b')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setFontSize(11);
    sh.setFrozenRows(1);
    sh.setTabColor(cfg.tabColor);

    // Auto-resize columns
    try { sh.autoResizeColumns(1, cfg.headers.length); } catch(e) {}
  });
}

/** Shows a popup telling you which sheets exist and which are missing. */
function diagnoseSheetsSetup() {
  const defs = getSheetDefs_();
  const allSheets = SS.getSheets().map(s => s.getName());
  const required = Object.keys(defs);

  let report = 'WMS Sheet Diagnosis\n';
  report += 'Spreadsheet: ' + SS.getName() + '\n';
  report += 'ID: ' + SS.getId() + '\n\n';
  report += '── Required Sheets ──\n';

  let allOk = true;
  required.forEach(name => {
    const exists = allSheets.includes(name);
    report += (exists ? '✅' : '❌') + ' ' + name + (exists ? '' : '  ← MISSING') + '\n';
    if (!exists) allOk = false;
  });

  report += '\n── All Tabs in Spreadsheet ──\n';
  allSheets.forEach(n => { report += '  • ' + n + '\n'; });

  if (!allOk) {
    report += '\n⚠ Run: WMS → Setup → Create All Sheets';
  } else {
    report += '\n✅ All sheets exist. Data should flow correctly.';
  }

  Logger.log(report);
  try { SpreadsheetApp.getUi().alert(report); } catch(e) { /* running from editor — see Logs */ }
}

/** Adds demo login users to the USERS sheet. */
function addDemoUsers() {
  const sh = SS.getSheetByName('USERS');
  if (!sh) { Logger.log('ERROR: Run createAllSheets() first'); throw new Error('Run createAllSheets() first'); return; }
  const existing = sh.getDataRange().getValues();
  const existingIds = existing.slice(1).map(r => String(r[0]).trim());

  const users = [
    ['admin',    'admin123', 'Admin User',   'ADMIN',    'YES'],
    ['wms1',     'wms123',   'WMS User 1',   'OPERATOR', 'YES'],
    ['wms2',     'wms123',   'WMS User 2',   'OPERATOR', 'YES'],
    ['manager',  'mgr123',   'WMS Manager',  'MANAGER',  'YES'],
  ];

  let added = 0;
  users.forEach(u => {
    if (!existingIds.includes(u[0])) {
      sh.appendRow(u);
      added++;
    }
  });

  Logger.log('✅ ' + added + ' user(s) added to USERS sheet. Login: admin/admin123, wms1/wms123, manager/mgr123');
}

/** Adds sample bin layout to BIN_MASTER. */
function addSampleBins() {
  const sh = SS.getSheetByName('BIN_MASTER');
  if (!sh) { Logger.log('ERROR: Run createAllSheets() first'); throw new Error('Run createAllSheets() first'); return; }
  if (sh.getLastRow() > 1) {
    Logger.log('BIN_MASTER already has data — appending more bins anyway.');
  }

  const rows = [];
  const sections = ['A', 'B', 'C'];
  for (let col = 1; col <= 17; col++) {
    for (let lvl = 1; lvl <= 12; lvl++) {
      const binId = 'R1-C' + col + '-' + String(lvl).padStart(3, '0');
      const section = sections[(col + lvl) % 3];
      const status = (col <= 9) ? 'OCCUPIED' : (col <= 11 && lvl % 3 === 0) ? 'ALLOCATED' : 'EMPTY';
      rows.push([binId, section, status, 'NO', '']);
    }
  }

  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log('✅ ' + rows.length + ' bins added to BIN_MASTER.');
}

/** Adds sample products to PRODUCT_MASTER. */
function addSampleProducts() {
  const sh = SS.getSheetByName('PRODUCT_MASTER');
  if (!sh) { Logger.log('ERROR: Run createAllSheets() first'); throw new Error('Run createAllSheets() first'); return; }
  if (sh.getLastRow() > 1) {
    Logger.log('PRODUCT_MASTER already has data — appending sample products anyway.');
  }

  const products = [
    ['MWLJNTP.0003.B0_N',       'LJ NutriMix 2+ 350gm Chocolate Jar',      32, 'A', 24],
    ['MWLJNTP.00059.B0_N',      'LJ NutriMix 7+ 350gm Chocolate Jar',      32, 'A', 24],
  ];

  sh.getRange(sh.getLastRow() + 1, 1, products.length, products[0].length).setValues(products);
  Logger.log('✅ ' + products.length + ' products added to PRODUCT_MASTER.');
}

/* ════════════════════════════════════════════════════════
   LOGIN
   ════════════════════════════════════════════════════════ */
function login(userId, password) {
  const sh = SS.getSheetByName('USERS');
  if (!sh) return { success: false, message: 'USERS sheet not found — contact Admin.' };

  userId   = String(userId   || '').trim();
  password = String(password || '').trim();
  if (!userId || !password) return { success: false, message: 'Enter User ID and Password' };

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowId   = String(data[i][0] || '').trim();
    const rowPw   = String(data[i][1] || '').trim();
    const rowName = String(data[i][2] || '').trim();
    const rowRole = String(data[i][3] || '').trim().toUpperCase();
    const active  = String(data[i][4] || '').trim().toUpperCase();

    if (rowId === userId && rowPw === password) {
      if (active !== 'YES') return { success: false, message: 'Account inactive — contact Admin.' };

      // Write last login timestamp
      try {
        const tz = Session.getScriptTimeZone();
        sh.getRange(i + 1, 6).setValue(Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy HH:mm:ss'));
      } catch(e) {}

      const rowSections = String(data[i][6] || '').trim(); // Column G — Allowed Sections
      return {
        success:         true,
        name:            rowName,
        role:            rowRole,
        userId:          rowId,
        allowedSections: rowSections,
      };
    }
  }
  return { success: false, message: 'Wrong User ID or Password' };
}

/* ════════════════════════════════════════════════════════
   DASHBOARD
   ════════════════════════════════════════════════════════ */
function getDashboardCounts() {
  const binSh = SS.getSheetByName('BIN_MASTER');
  if (!binSh) return { error: 'BIN_MASTER sheet missing. Run Setup.' };

  const bin = binSh.getDataRange().getDisplayValues();
  let total = 0, empty = 0, blocked = 0, occupied = 0, allocated = 0;

  for (let i = 1; i < bin.length; i++) {
    if (!String(bin[i][0] || '').trim()) continue;
    total++;
    const status  = String(bin[i][2] || '').trim().toUpperCase();
    const isBlock = String(bin[i][3] || '').trim().toUpperCase() === 'YES';
    if (isBlock)             blocked++;
    else if (status === 'EMPTY')     empty++;
    else if (status === 'OCCUPIED')  occupied++;
    else if (status === 'ALLOCATED') allocated++;
  }

  const toStats = getTOStats_();
  const liveStats = getLiveStockStats_();

  return {
    total, empty, blocked, occupied, allocated,
    pendingTO: toStats.pending,
    confirmedTO: toStats.confirmed,
    totalGoodStock: liveStats.good,
    totalDamageStock: liveStats.damage,
    activeSKUs: liveStats.skus
  };
}

function getTOStats_() {
  const sh = SS.getSheetByName('TO_HEADER');
  if (!sh) return { pending: 0, confirmed: 0 };
  const data = sh.getDataRange().getDisplayValues();
  let pending = 0, confirmed = 0;
  for (let i = 1; i < data.length; i++) {
    const st = String(data[i][4] || '').trim().toUpperCase();
    if (st === 'PENDING')   pending++;
    if (st === 'CONFIRMED') confirmed++;
  }
  return { pending, confirmed };
}

function getLiveStockStats_() {
  const sh = SS.getSheetByName('LIVE_STOCK');
  if (!sh) return { good: 0, damage: 0, skus: 0 };
  const data = sh.getDataRange().getDisplayValues();
  let good = 0, damage = 0;
  const skuSet = new Set();
  for (let i = 1; i < data.length; i++) {
    const mat  = String(data[i][2] || '').trim();
    const bal  = Number(data[i][6] || 0);
    const type = String(data[i][14] || 'GOOD').trim().toUpperCase();
    if (!mat || bal <= 0) continue;
    if (type === 'GOOD')                             { good += bal; skuSet.add(mat); }
    if (type === 'DAMAGE' || type === 'BLOCKED')      damage += bal;
  }
  return { good, damage, skus: skuSet.size };
}

/* ════════════════════════════════════════════════════════
   UNLOADING / INWARD
   ════════════════════════════════════════════════════════ */
function getPendingUnloadingList() {
  const uH  = SS.getSheetByName('UNLOADING_HEADER');
  const toH = SS.getSheetByName('TO_HEADER');
  if (!uH || !toH) return [];

  const uData  = uH.getDataRange().getDisplayValues();
  const toData = toH.getDataRange().getDisplayValues();

  // Build set of inward numbers that already have a TO
  const inwardWithTO = new Set();
  for (let i = 1; i < toData.length; i++) {
    inwardWithTO.add(String(toData[i][1] || '').trim());
  }

  const list = [];
  for (let i = 1; i < uData.length; i++) {
    const inwardNo = String(uData[i][0] || '').trim();
    const status   = String(uData[i][11] || '').trim().toUpperCase();
    if (!inwardNo) continue;
    if (inwardWithTO.has(inwardNo)) continue;
    if (status === 'CONFIRMED' || status === 'TO GENERATED') continue;
    list.push({
      inwardNo,
      invoiceNo:   uData[i][2],
      invoiceDate: uData[i][3],
      vehicleNo:   uData[i][4],
      transport:   uData[i][5],
      status:      uData[i][11] || 'DRAFT'
    });
  }
  return list;
}

function loadPendingUnloading(inwardNo) {
  inwardNo = String(inwardNo || '').trim();
  const uH = SS.getSheetByName('UNLOADING_HEADER');
  const uD = SS.getSheetByName('UNLOADING_DETAILS');
  if (!uH || !uD) return { success: false, message: 'Sheets missing. Run Setup.' };

  const uHData = uH.getDataRange().getDisplayValues();
  const uDData = uD.getDataRange().getDisplayValues();

  const uh = uHData.find((r, i) => i > 0 && String(r[0]).trim() === inwardNo);
  if (!uh) return { success: false, message: 'Inward No not found: ' + inwardNo };

  const oldTO = getTOByInward_(inwardNo);
  if (oldTO) return { success: false, message: 'This inward already has TO No: ' + oldTO };

  const rows = uDData
    .filter((r, i) => i > 0 && String(r[0]).trim() === inwardNo)
    .map(r => ({
      materialCode: r[2],
      materialName: r[3],
      batch:        r[4],
      qty:          Number(r[5] || 0),
      palletSize:   Number(r[6] || 0),
      section:      r[7],
      caseLot:      r[8]
    }));

  return {
    success: true, inwardNo,
    header: {
      invoiceNo:    uh[2], invoiceDate: uh[3],
      vehicleNo:    uh[4], transport:   uh[5],
      lrNo:         uh[6], lrDate:      uh[7],
      fromLocation: uh[8], toLocation:  uh[9]
    },
    items: rows
  };
}

function saveUnloading(header, items, user) {
  if (!items || items.length === 0) return { success: false, message: 'No items to save' };
  const hSh = SS.getSheetByName('UNLOADING_HEADER');
  const dSh = SS.getSheetByName('UNLOADING_DETAILS');
  if (!hSh || !dSh) return { success: false, message: 'Sheets missing. Run Setup.' };

  const inwardNo = new Date().getTime();
  hSh.appendRow([
    inwardNo, new Date(),
    header.invoiceNo || '', header.invoiceDate || '',
    header.vehicleNo || '', header.transport   || '',
    header.lrNo      || '', header.lrDate      || '',
    header.fromLocation || '', header.toLocation || '',
    user || '', 'DRAFT'
  ]);

  items.forEach((it, i) => {
    dSh.appendRow([
      inwardNo, i + 1,
      it.materialCode, it.materialName, it.batch,
      Number(it.qty || 0), Number(it.palletSize || 0),
      it.section || '', it.caseLot || ''
    ]);
  });

  return { success: true, inwardNo };
}

/* ════════════════════════════════════════════════════════
   PRODUCT MASTER LOOKUP
   ════════════════════════════════════════════════════════ */
function getProduct(code) {
  const sh = SS.getSheetByName('PRODUCT_MASTER');
  if (!sh) return { found: false, message: 'PRODUCT_MASTER sheet missing' };
  const data = sh.getDataRange().getDisplayValues();
  code = String(code || '').trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === code) {
      return {
        found:        true,
        materialCode: data[i][0],
        materialName: data[i][1],
        palletSize:   Number(data[i][2] || 0),
        section:      data[i][3],
        caseLot:      Number(data[i][4] || 0)
      };
    }
  }
  return { found: false, message: 'Material not found: ' + code };
}

/* ════════════════════════════════════════════════════════
   TO (TRANSFER ORDER) GENERATION
   ════════════════════════════════════════════════════════ */
function getOpenTOList() {
  const sh = SS.getSheetByName('TO_HEADER');
  if (!sh) return [];
  const data = sh.getDataRange().getDisplayValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const toNo   = String(data[i][0] || '').trim();
    const status = String(data[i][4] || '').trim().toUpperCase();
    if (toNo && status === 'PENDING') {
      list.push({ toNo, inwardNo: data[i][1], status });
    }
  }
  return list;
}

function getNextTONo_() {
  const sh = SS.getSheetByName('TO_HEADER');
  if (!sh || sh.getLastRow() < 2) return 1000;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat();
  const nums  = vals.map(Number).filter(n => !isNaN(n) && n >= 1000);
  return nums.length ? Math.max(...nums) + 1 : 1000;
}

function generateTO(inwardNo, user) {
  const detailSh = SS.getSheetByName('UNLOADING_DETAILS');
  const toHSh    = SS.getSheetByName('TO_HEADER');
  const toDSh    = SS.getSheetByName('TO_DETAILS');
  if (!detailSh || !toHSh || !toDSh) return { success: false, message: 'Sheets missing. Run Setup.' };

  const detail = detailSh.getDataRange().getValues();
  const rows   = detail.filter((r, i) => i > 0 && String(r[0]) === String(inwardNo));
  if (!rows.length) return { success: false, message: 'Save unloading first before generating TO' };

  const oldTO = getTOByInward_(inwardNo);
  if (oldTO) return { success: false, message: 'This inward already has TO No: ' + oldTO };

  const toNo = getNextTONo_();
  toHSh.appendRow([toNo, inwardNo, new Date(), user || '', 'PENDING']);

  let line = 1;
  rows.forEach(r => {
    let qty     = Number(r[5] || 0);
    let pallet  = Number(r[6] || 0) || qty;
    let section = String(r[7] || '').trim();

    while (qty > 0) {
      const pickQty = qty >= pallet ? pallet : qty;
      const bin     = getEmptyBin(section);
      if (bin === 'NO_BIN') throw new Error('No empty bin available for section: ' + section);

      toDSh.appendRow([toNo, line++, r[2], r[3], r[4], bin, pickQty, pallet,
                       'PENDING', '', '', '', '', '', '']);
      markBin_(bin, 'ALLOCATED');
      qty -= pickQty;
    }
  });

  updateInwardStatus_(inwardNo, 'TO GENERATED');
  return { success: true, toNo };
}

function getTOByInward_(inwardNo) {
  const sh = SS.getSheetByName('TO_HEADER');
  if (!sh) return '';
  const data = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').trim() === String(inwardNo).trim()) return data[i][0];
  }
  return '';
}

function getTODetails(toNo) {
  const sh = SS.getSheetByName('TO_DETAILS');
  if (!sh) return [];
  const data = sh.getDataRange().getDisplayValues();
  return data
    .filter((r, i) => i > 0 && String(r[0]).trim() === String(toNo).trim())
    .map(r => ({
      lineNo:      r[1],
      materialCode:r[2],  materialName: r[3],
      batch:       r[4],  binNo:        r[5],
      qty:         Number(r[6]  || 0),
      palletSize:  Number(r[7]  || 0),
      status:      r[8],
      confirmQty:  r[9],  damageQty:    r[10],
      shortageQty: r[11], remark:       r[12],
      excessQty:   r[13], cqaQty:       r[14]
    }));
}

function loadOldTO(toNo) {
  toNo = String(toNo || '').trim();
  const toHSh = SS.getSheetByName('TO_HEADER');
  const uHSh  = SS.getSheetByName('UNLOADING_HEADER');
  if (!toHSh) return { success: false, message: 'TO_HEADER sheet missing' };

  const toHData = toHSh.getDataRange().getDisplayValues();
  const th = toHData.find((r, i) => i > 0 && String(r[0]).trim() === toNo);
  if (!th) return { success: false, message: 'TO not found: ' + toNo };

  const inwardNo = String(th[1] || '').trim();
  const status   = String(th[4] || '').trim().toUpperCase();

  let uh = null;
  if (uHSh) {
    const uHData = uHSh.getDataRange().getDisplayValues();
    uh = uHData.find((r, i) => i > 0 && String(r[0]).trim() === inwardNo);
  }

  return {
    success: true, toNo, inwardNo, status,
    alreadyConfirmed: status === 'CONFIRMED',
    header: {
      invoiceNo:    uh ? uh[2] : '', invoiceDate: uh ? uh[3] : '',
      vehicleNo:    uh ? uh[4] : '', transport:   uh ? uh[5] : '',
      lrNo:         uh ? uh[6] : '', lrDate:      uh ? uh[7] : '',
      fromLocation: uh ? uh[8] : '', toLocation:  uh ? uh[9] : ''
    },
    rows: getTODetails(toNo)
  };
}

/* ════════════════════════════════════════════════════════
   TO CONFIRMATION
   ════════════════════════════════════════════════════════ */
function confirmTO(toNo, rows, header, user) {
  if (!rows || rows.length === 0) return { success: false, message: 'No rows to confirm' };

  const currentStatus = getTOStatus_(toNo);
  if (currentStatus === 'CONFIRMED') return { success: false, message: 'This TO is already CONFIRMED.' };

  const toDSh   = SS.getSheetByName('TO_DETAILS');
  const stockSh = SS.getSheetByName('LIVE_STOCK');
  if (!toDSh || !stockSh) return { success: false, message: 'Sheets missing. Run Setup.' };

  const toData = toDSh.getDataRange().getDisplayValues();
  let sr = getNextSr_('LIVE_STOCK', 2);

  rows.forEach(c => {
    const planQty    = Number(c.qty       || 0);
    const confirmQty = Number(c.confirmQty|| 0);
    const cqaQty     = Number(c.cqaQty   || 0);
    const damageQty  = Number(c.damageQty || 0);
    const shortageQty= Number(c.shortageQty || 0);
    const excessQty  = Number(c.excessQty || 0);
    const remark     = c.remark || '';

    if (confirmQty < 0 || cqaQty < 0 || damageQty < 0 || shortageQty < 0 || excessQty < 0)
      throw new Error('Negative quantities not allowed');
    if (confirmQty + cqaQty + damageQty + shortageQty > planQty)
      throw new Error('Line ' + c.lineNo + ': Confirm + CQA + Damage + Shortage cannot exceed Plan Qty (' + planQty + ')');

    // Update TO_DETAILS row
    for (let i = 1; i < toData.length; i++) {
      if (String(toData[i][0]).trim() === String(toNo).trim() &&
          String(toData[i][1]).trim() === String(c.lineNo).trim()) {
        toDSh.getRange(i + 1, 9).setValue('CONFIRMED');
        toDSh.getRange(i + 1, 10).setValue(confirmQty);
        toDSh.getRange(i + 1, 11).setValue(damageQty);
        toDSh.getRange(i + 1, 12).setValue(shortageQty);
        toDSh.getRange(i + 1, 13).setValue(remark);
        toDSh.getRange(i + 1, 14).setValue(excessQty);
        toDSh.getRange(i + 1, 15).setValue(cqaQty);
        break;
      }
    }

    // Write to LIVE_STOCK
    if (confirmQty > 0) {
      stockSh.appendRow(makeStockRow_(sr++, c, c.binNo, confirmQty, header, toNo, 'GOOD'));
      markBin_(c.binNo, 'OCCUPIED');
    }
    if (cqaQty > 0)      stockSh.appendRow(makeStockRow_(sr++, c, c.binNo,        cqaQty,      header, toNo, 'CQA'));
    if (damageQty > 0)   stockSh.appendRow(makeStockRow_(sr++, c, 'DAMAGE-BIN',   damageQty,   header, toNo, 'DAMAGE'));
    if (shortageQty > 0) stockSh.appendRow(makeStockRow_(sr++, c, 'SHORTAGE-BIN', shortageQty, header, toNo, 'SHORTAGE'));
    if (excessQty > 0)   stockSh.appendRow(makeStockRow_(sr++, c, 'EXCESS-BIN',   excessQty,   header, toNo, 'EXCESS'));

    if (confirmQty <= 0 && cqaQty <= 0) markBin_(c.binNo, 'EMPTY');
  });

  updateTOStatus_(toNo, 'CONFIRMED');
  updateInwardStatusByTO_(toNo, 'CONFIRMED');
  return { success: true, message: 'TO Confirmed. Stock updated (GOOD / CQA / Damage / Shortage).' };
}

function makeStockRow_(sr, c, bin, qty, header, toNo, type) {
  const key = String(c.materialCode) + '|' + String(c.batch) + '|' + String(bin);
  return [
    key, sr,
    c.materialCode, c.materialName, c.batch,
    bin, qty, qty,
    header.vehicleNo    || '',
    header.invoiceNo    || '',
    header.transport    || '',
    new Date(),
    header.fromLocation || '',
    toNo, type
  ];
}

function getTOStatus_(toNo) {
  const sh = SS.getSheetByName('TO_HEADER');
  if (!sh) return '';
  const data = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(toNo).trim()) return String(data[i][4]).trim().toUpperCase();
  }
  return '';
}

/* ════════════════════════════════════════════════════════
   BIN MANAGEMENT
   ════════════════════════════════════════════════════════ */
function getEmptyBin(section) {
  const sh = SS.getSheetByName('BIN_MASTER');
  if (!sh) throw new Error('BIN_MASTER sheet missing. Run Setup.');
  const data = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    const binSection = String(data[i][1] || '').trim();
    const status     = String(data[i][2] || '').trim().toUpperCase();
    const blocked    = String(data[i][3] || '').trim().toUpperCase();
    if (binSection === String(section).trim() &&
        status === 'EMPTY' &&
        blocked !== 'YES') {
      return data[i][0];
    }
  }
  return 'NO_BIN';
}

function markBin_(bin, status) {
  if (['DAMAGE-BIN', 'EXCESS-BIN', 'SHORTAGE-BIN'].includes(String(bin))) return;
  const sh = SS.getSheetByName('BIN_MASTER');
  if (!sh) return;
  const data = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(bin).trim()) {
      sh.getRange(i + 1, 3).setValue(status);
      return;
    }
  }
}

// Keep old name for backward compatibility with original code
function markBin(bin, status) { return markBin_(bin, status); }

function refreshBinsFromLiveStock_() {
  const liveSh = SS.getSheetByName('LIVE_STOCK');
  const binSh  = SS.getSheetByName('BIN_MASTER');
  if (!liveSh || !binSh) return;

  const liveData = liveSh.getDataRange().getDisplayValues();
  const binBal   = {};

  for (let i = 1; i < liveData.length; i++) {
    const bin  = String(liveData[i][5] || '').trim();
    const bal  = Number(liveData[i][6] || 0);
    const type = String(liveData[i][14] || 'GOOD').trim().toUpperCase();
    if (bin && type === 'GOOD') binBal[bin] = (binBal[bin] || 0) + bal;
  }

  const binData = binSh.getDataRange().getDisplayValues();
  for (let i = 1; i < binData.length; i++) {
    const bin     = String(binData[i][0] || '').trim();
    const blocked = String(binData[i][3] || '').trim().toUpperCase();
    if (!bin || blocked === 'YES') continue;
    binSh.getRange(i + 1, 3).setValue((binBal[bin] || 0) > 0 ? 'OCCUPIED' : 'EMPTY');
  }
}

/* ════════════════════════════════════════════════════════
   STATUS UPDATES
   ════════════════════════════════════════════════════════ */
function updateTOStatus_(toNo, status) {
  const sh = SS.getSheetByName('TO_HEADER');
  if (!sh) return;
  const data = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(toNo).trim()) {
      sh.getRange(i + 1, 5).setValue(status);
      return;
    }
  }
}

function updateInwardStatus_(inwardNo, status) {
  const sh = SS.getSheetByName('UNLOADING_HEADER');
  if (!sh) return;
  const data = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(inwardNo).trim()) {
      sh.getRange(i + 1, 12).setValue(status);
      return;
    }
  }
}

function updateInwardStatusByTO_(toNo, status) {
  const sh = SS.getSheetByName('TO_HEADER');
  if (!sh) return;
  const data = sh.getDataRange().getDisplayValues();
  const th = data.find((r, i) => i > 0 && String(r[0]).trim() === String(toNo).trim());
  if (th) updateInwardStatus_(th[1], status);
}

/* ════════════════════════════════════════════════════════
   STOCK REPORTS
   ════════════════════════════════════════════════════════ */
let __runtimeLoadingPending = {};

function getAvailableStockReport() {
  const sh = SS.getSheetByName('LIVE_STOCK');
  if (!sh) return [];

  const data       = sh.getDataRange().getDisplayValues();
  const pendingMap = getPendingPickMap_();
  const map        = {};

  for (let i = 1; i < data.length; i++) {
    const materialCode = String(data[i][2]  || '').trim();
    const materialName = String(data[i][3]  || '').trim();
    const batch        = String(data[i][4]  || '').trim();
    const binNo        = String(data[i][5]  || '').trim();
    const balance      = Number(data[i][6]  || 0);
    const stockType    = String(data[i][14] || 'GOOD').trim().toUpperCase();

    if (!materialCode || !binNo || balance <= 0) continue;

    const key = materialCode + '|' + batch + '|' + binNo + '|' + stockType;
    if (!map[key]) map[key] = { materialCode, materialName, batch, binNo, stockType, stockQty: 0, pickingQty: 0, qty: 0, deliveryNo: '', loadingNo: '' };
    map[key].stockQty += balance;
  }

  Object.keys(map).forEach(key => {
    const r = map[key];
    if (r.stockType === 'GOOD') {
      const p = pendingMap[r.materialCode + '|' + r.batch + '|' + r.binNo];
      r.pickingQty = p ? Number(p.qty || 0) : 0;
      r.deliveryNo = p ? Object.keys(p.deliveryNos || {}).join(', ') : '';
      r.loadingNo  = p ? Object.keys(p.loadingNos  || {}).join(', ') : '';
      r.qty = Math.max(0, r.stockQty - r.pickingQty);
    } else {
      r.qty = r.stockQty;
    }
  });

  return Object.values(map).sort((a, b) =>
    String(a.materialCode).localeCompare(String(b.materialCode), undefined, { numeric: true }) ||
    String(a.batch).localeCompare(String(b.batch), undefined, { numeric: true }) ||
    String(a.binNo).localeCompare(String(b.binNo), undefined, { numeric: true })
  );
}

function getStockMovementReport(filter) {
  filter = filter || {};
  const fMaterial   = String(filter.materialCode || '').trim().toUpperCase();
  const fBatch      = String(filter.batch        || '').trim().toUpperCase();
  const inclPending = String(filter.includePending || 'YES').toUpperCase() !== 'NO';
  const rows = [];

  function match_(mat, bat) {
    if (fMaterial && String(mat || '').trim().toUpperCase() !== fMaterial) return false;
    if (fBatch    && String(bat || '').trim().toUpperCase() !== fBatch)    return false;
    return true;
  }

  const liveSh = SS.getSheetByName('LIVE_STOCK');
  if (liveSh) {
    const d = liveSh.getDataRange().getDisplayValues();
    for (let i = 1; i < d.length; i++) {
      const mc = String(d[i][2] || '').trim(), mn = String(d[i][3] || '').trim();
      const bt = String(d[i][4] || '').trim(), bn = String(d[i][5] || '').trim();
      const bal = Number(d[i][6] || 0), inQty = Number(d[i][7] || 0) || bal;
      const type = String(d[i][14] || 'GOOD').trim().toUpperCase();
      if (!match_(mc, bt) || !mc || inQty <= 0) continue;
      rows.push({ date: String(d[i][11] || ''), movement: 'INWARD', status: type, materialCode: mc, materialName: mn, batch: bt, binNo: bn, inQty, outQty: 0, balanceQty: bal, deliveryNo: '', cfaCode: '', cfaName: '', invoiceNo: String(d[i][9] || ''), vehicleNo: String(d[i][8] || ''), refNo: String(d[i][13] || ''), remark: String(d[i][12] || '') });
    }
  }

  const allocSh = SS.getSheetByName('LOADING_ALLOC');
  if (allocSh) {
    const a = allocSh.getDataRange().getDisplayValues();
    for (let i = 1; i < a.length; i++) {
      const lNo = String(a[i][0] || ''), pl = String(a[i][1] || ''), dn = String(a[i][2] || '');
      const cc = String(a[i][3] || ''), cn = String(a[i][4] || ''), mc = String(a[i][5] || ''), mn = String(a[i][6] || '');
      const bt = String(a[i][7] || ''), bn = String(a[i][8] || '');
      const pq = Number(a[i][9] || 0), st = String(a[i][10] || '').toUpperCase();
      const cq = Number(a[i][11] || 0), sq = Number(a[i][12] || 0);
      const rm = String(a[i][13] || ''), cb = String(a[i][14] || ''), cd = String(a[i][15] || '');
      if (!match_(mc, bt) || !mc || pq <= 0) continue;
      if (st === 'CONFIRMED') rows.push({ date: cd, movement: 'OUTWARD', status: 'CONFIRMED', materialCode: mc, materialName: mn, batch: bt, binNo: bn, inQty: 0, outQty: cq, balanceQty: '', deliveryNo: dn, cfaCode: cc, cfaName: cn, invoiceNo: '', vehicleNo: '', refNo: lNo + ' / Line ' + pl, remark: rm + (sq > 0 ? ' | Short: ' + sq : '') + (cb ? ' | By: ' + cb : '') });
      else if (inclPending) rows.push({ date: '', movement: 'PICKING', status: 'PENDING', materialCode: mc, materialName: mn, batch: bt, binNo: bn, inQty: 0, outQty: pq, balanceQty: '', deliveryNo: dn, cfaCode: cc, cfaName: cn, invoiceNo: '', vehicleNo: '', refNo: lNo + ' / Line ' + pl, remark: 'Reserved' + (rm ? ' | ' + rm : '') });
    }
  }

  rows.sort((a, b) =>
    String(a.materialCode).localeCompare(String(b.materialCode), undefined, { numeric: true }) ||
    String(a.batch).localeCompare(String(b.batch), undefined, { numeric: true }) ||
    String(a.date).localeCompare(String(b.date), undefined, { numeric: true })
  );

  const totalIn = rows.reduce((s, r) => s + Number(r.inQty || 0), 0);
  const totalOut = rows.filter(r => r.movement === 'OUTWARD').reduce((s, r) => s + Number(r.outQty || 0), 0);
  const totalPicking = rows.filter(r => r.movement === 'PICKING').reduce((s, r) => s + Number(r.outQty || 0), 0);
  return { rows, totalIn, totalOut, totalPicking, net: totalIn - totalOut };
}

/* ════════════════════════════════════════════════════════
   LOADING / OUTWARD / DISPATCH
   ════════════════════════════════════════════════════════ */
function saveLoadingOrder(items, user) {
  if (!items || items.length === 0) return { success: false, message: 'No items in loading order' };
  const hSh = SS.getSheetByName('LOADING_HEADER');
  const dSh = SS.getSheetByName('LOADING_DETAILS');
  if (!hSh || !dSh) return { success: false, message: 'Sheets missing. Run Setup.' };

  const nowKey = new Date().getTime();
  const groups = {};

  items.forEach((it, idx) => {
    const deliveryNo = String(it.deliveryNo || '').trim();
    if (!deliveryNo) throw new Error('Delivery No required. Row: ' + (idx + 1));
    if (!groups[deliveryNo]) groups[deliveryNo] = [];
    groups[deliveryNo].push(it);
  });

  const deliveryNos = Object.keys(groups).sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
  const loadingNos = [], slips = [];

  deliveryNos.forEach((deliveryNo, gIdx) => {
    const groupItems = groups[deliveryNo];
    const loadingNo  = 'L' + nowKey + '-' + (gIdx + 1);
    let totalQty     = 0;

    hSh.appendRow([loadingNo, new Date(), user || '', 'DRAFT', groupItems.length, 0]);
    groupItems.forEach((it, idx) => {
      const qty = Number(it.qty || 0);
      totalQty += qty;
      dSh.appendRow([loadingNo, idx + 1, deliveryNo, it.cfaCode || '', it.cfaName || '', it.materialCode || '', it.materialName || '', it.batch || '', qty, 'DRAFT']);
    });
    hSh.getRange(hSh.getLastRow(), 6).setValue(totalQty);
    loadingNos.push(loadingNo);
    slips.push({ loadingNo, deliveryNo, totalLines: groupItems.length, totalQty });
  });

  return { success: true, loadingNo: loadingNos[0], loadingNos, slips, deliveryCount: deliveryNos.length };
}

function getOpenLoadingList() {
  const sh = SS.getSheetByName('LOADING_HEADER');
  if (!sh) return [];
  const data = sh.getDataRange().getDisplayValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][3] || '').trim().toUpperCase();
    if (status === 'DRAFT' || status === 'PICK GENERATED' || status === 'PENDING') {
      list.push({ loadingNo: data[i][0], createdDate: data[i][1], createdBy: data[i][2], status: data[i][3], totalLines: data[i][4], totalQty: data[i][5] });
    }
  }
  return list;
}

function getLoadingDetails(loadingNo) {
  const sh = SS.getSheetByName('LOADING_DETAILS');
  if (!sh) return [];
  const data = sh.getDataRange().getDisplayValues();
  loadingNo = String(loadingNo || '').trim();
  return data.filter((r, i) => i > 0 && String(r[0]).trim() === loadingNo).map(r => ({
    lineNo: r[1], deliveryNo: r[2], cfaCode: r[3], cfaName: r[4],
    materialCode: r[5], materialName: r[6], batch: r[7], qty: Number(r[8] || 0), status: r[9]
  }));
}

function getLoadingAllocRows(loadingNo) {
  const sh = SS.getSheetByName('LOADING_ALLOC');
  if (!sh) return [];
  const data = sh.getDataRange().getDisplayValues();
  loadingNo  = String(loadingNo || '').trim();
  return data.filter((r, i) => i > 0 && String(r[0]).trim() === loadingNo).map(r => ({
    loadingNo: r[0], pickLine: r[1], deliveryNo: r[2], cfaCode: r[3], cfaName: r[4],
    materialCode: r[5], materialName: r[6], batch: r[7], binNo: r[8],
    pickQty: Number(r[9] || 0), status: r[10], confirmQty: r[11], shortQty: r[12], remark: r[13]
  }));
}

function getLoadingAllocRowsMulti(loadingNos) {
  if (!loadingNos || !loadingNos.length) return [];
  return loadingNos.flatMap(no => getLoadingAllocRows(no));
}

function generateLoadingPick(loadingNo, user) {
  loadingNo = String(loadingNo || '').trim();
  if (!loadingNo) return { success: false, message: 'Loading No missing' };

  const status = getLoadingStatus_(loadingNo);
  if (status === 'CONFIRMED') return { success: false, message: 'This loading is already CONFIRMED' };
  if (getLoadingAllocRows(loadingNo).length > 0) return { success: false, message: 'Pick already generated for ' + loadingNo };

  const details = getLoadingDetails(loadingNo);
  if (!details.length) return { success: false, message: 'No loading details found for ' + loadingNo };

  details.sort((a, b) =>
    String(a.deliveryNo).localeCompare(String(b.deliveryNo), undefined, { numeric: true }) ||
    String(a.materialCode).localeCompare(String(b.materialCode), undefined, { numeric: true }) ||
    String(a.batch).localeCompare(String(b.batch), undefined, { numeric: true })
  );

  __runtimeLoadingPending = {};
  const allocSh = SS.getSheetByName('LOADING_ALLOC');
  let pickLine  = 1;

  details.forEach(it => {
    let need = Number(it.qty || 0);
    if (!it.deliveryNo)   throw new Error('Delivery No required. Delivery: ' + it.deliveryNo);
    if (!it.materialCode) throw new Error('Material Code required');
    if (!it.batch)        throw new Error('Batch required');
    if (need <= 0)        throw new Error('Qty must be > 0');

    const palletSize = getLoadingPalletSize_(it.materialCode);
    const fullQty    = palletSize > 0 ? Math.floor(need / palletSize) * palletSize : 0;
    const looseQty   = need - fullQty;

    const totalAvail = getAvailableStockRowsForPick_(it.materialCode, it.batch, palletSize, 'ALL').reduce((s, r) => s + Number(r.available || 0), 0);
    if (totalAvail < need) throw new Error('Not enough stock for Delivery ' + it.deliveryNo + ' | ' + it.materialCode + ' / ' + it.batch + '. Required: ' + need + ', Available: ' + totalAvail);

    if (fullQty  > 0) pickLine = allocateLoadingQty_(allocSh, loadingNo, pickLine, it, fullQty,  palletSize, 'FULL');
    if (looseQty > 0) pickLine = allocateLoadingQty_(allocSh, loadingNo, pickLine, it, looseQty, palletSize, 'LOOSE');
  });

  updateLoadingStatus_(loadingNo, 'PICK GENERATED');
  updateLoadingDetailsStatus_(loadingNo, 'PICK GENERATED');
  return { success: true, message: 'Pick generated', loadingNo, rows: getLoadingAllocRows(loadingNo) };
}

function allocateLoadingQty_(allocSh, loadingNo, pickLine, it, qtyNeeded, palletSize, mode) {
  let need      = Number(qtyNeeded || 0);
  const stockRows = getAvailableStockRowsForPick_(it.materialCode, it.batch, palletSize, mode);

  for (let s = 0; s < stockRows.length && need > 0; s++) {
    const avail = Number(stockRows[s].available || 0);
    let take    = 0;
    if (mode === 'FULL' && palletSize > 0) {
      take = Math.min(need, Math.floor(avail / palletSize) * palletSize);
    } else {
      take = Math.min(need, avail);
    }
    if (take <= 0) continue;

    const pickType = mode === 'FULL' ? 'FULL PALLET' : 'LOOSE PICK';
    allocSh.appendRow([loadingNo, pickLine++, it.deliveryNo, it.cfaCode, it.cfaName, it.materialCode, it.materialName, it.batch, stockRows[s].binNo, take, 'PENDING', '', '', pickType, '', '']);

    const key = String(it.materialCode).trim() + '|' + String(it.batch).trim() + '|' + String(stockRows[s].binNo).trim();
    __runtimeLoadingPending[key] = (__runtimeLoadingPending[key] || 0) + take;
    need -= take;
  }

  if (need > 0) throw new Error('Allocation failed for Delivery ' + it.deliveryNo + ' | ' + it.materialCode + ' / ' + it.batch + '. Unallocated qty: ' + need);
  return pickLine;
}

function confirmLoadingPick(loadingNo, rows, user) {
  loadingNo = String(loadingNo || '').trim();
  if (!rows || rows.length === 0) return { success: false, message: 'No rows to confirm' };
  if (getLoadingStatus_(loadingNo) === 'CONFIRMED') return { success: false, message: 'This loading is already CONFIRMED' };

  const allocSh   = SS.getSheetByName('LOADING_ALLOC');
  if (!allocSh) return { success: false, message: 'LOADING_ALLOC sheet missing' };
  const allocData = allocSh.getDataRange().getDisplayValues();

  rows.forEach(r => {
    const pickQty   = Number(r.pickQty    || 0);
    const confirmQty= Number(r.confirmQty || 0);
    const shortQty  = Number(r.shortQty   || 0);
    if (confirmQty < 0 || shortQty < 0) throw new Error('Negative qty not allowed');
    if (confirmQty + shortQty > pickQty)  throw new Error('Pick line ' + r.pickLine + ': Confirm + Short > Pick Qty');

    if (confirmQty > 0) reduceLiveStock_(r.materialCode, r.batch, r.binNo, confirmQty);

    for (let i = 1; i < allocData.length; i++) {
      if (String(allocData[i][0]).trim() === loadingNo && String(allocData[i][1]).trim() === String(r.pickLine).trim()) {
        allocSh.getRange(i + 1, 11).setValue('CONFIRMED');
        allocSh.getRange(i + 1, 12).setValue(confirmQty);
        allocSh.getRange(i + 1, 13).setValue(shortQty);
        allocSh.getRange(i + 1, 14).setValue(r.remark || '');
        allocSh.getRange(i + 1, 15).setValue(user || '');
        allocSh.getRange(i + 1, 16).setValue(new Date());
        break;
      }
    }
  });

  updateLoadingStatus_(loadingNo, 'CONFIRMED');
  updateLoadingDetailsStatus_(loadingNo, 'CONFIRMED');
  refreshBinsFromLiveStock_();
  return { success: true, message: 'Loading confirmed. Stock updated.' };
}

function getLoadingStatus_(loadingNo) {
  const sh = SS.getSheetByName('LOADING_HEADER');
  if (!sh) return '';
  const data = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(loadingNo).trim()) return String(data[i][3]).trim().toUpperCase();
  }
  return '';
}

function updateLoadingStatus_(loadingNo, status) {
  const sh = SS.getSheetByName('LOADING_HEADER');
  if (!sh) return;
  const data = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(loadingNo).trim()) { sh.getRange(i + 1, 4).setValue(status); return; }
  }
}

function updateLoadingDetailsStatus_(loadingNo, status) {
  const sh = SS.getSheetByName('LOADING_DETAILS');
  if (!sh) return;
  const data = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(loadingNo).trim()) sh.getRange(i + 1, 10).setValue(status);
  }
}

function getLoadingPalletSize_(materialCode) {
  const sh = SS.getSheetByName('PRODUCT_MASTER');
  if (!sh) return 0;
  const data = sh.getDataRange().getDisplayValues();
  materialCode = String(materialCode || '').trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === materialCode) return Number(data[i][2] || 0);
  }
  return 0;
}

function getAvailableStockRowsForPick_(materialCode, batch, palletSize, mode) {
  const sh = SS.getSheetByName('LIVE_STOCK');
  if (!sh) return [];
  const data       = sh.getDataRange().getDisplayValues();
  const pendingMap = getPendingLoadingAllocMap_();
  const list       = [];
  palletSize = Number(palletSize || 0);
  mode       = String(mode || 'ALL').toUpperCase();

  for (let i = 1; i < data.length; i++) {
    const mc  = String(data[i][2]  || '').trim();
    const mn  = String(data[i][3]  || '').trim();
    const bt  = String(data[i][4]  || '').trim();
    const bin = String(data[i][5]  || '').trim();
    const bal = Number(data[i][6]  || 0);
    const type= String(data[i][14] || 'GOOD').trim().toUpperCase();
    if (mc !== String(materialCode).trim() || bt !== String(batch).trim() || type !== 'GOOD' || bal <= 0) continue;
    const key     = mc + '|' + bt + '|' + bin;
    const pending = (pendingMap[key] || 0) + (__runtimeLoadingPending[key] || 0);
    const avail   = bal - pending;
    if (avail <= 0) continue;
    const isFullPallet = palletSize > 0 && avail >= palletSize && avail % palletSize === 0;
    const fullMultiple = palletSize > 0 ? Math.floor(avail / palletSize) * palletSize : 0;
    list.push({ rowNo: i + 1, materialCode: mc, materialName: mn, batch: bt, binNo: bin, balance: bal, pending, available: avail, isFullPallet, fullMultipleQty: fullMultiple, looseQty: palletSize > 0 ? avail % palletSize : avail });
  }

  if (mode === 'FULL' && palletSize > 0) {
    return list.filter(r => r.fullMultipleQty > 0).sort((a, b) =>
      (b.isFullPallet ? 1 : 0) - (a.isFullPallet ? 1 : 0) ||
      b.fullMultipleQty - a.fullMultipleQty ||
      String(a.binNo).localeCompare(String(b.binNo), undefined, { numeric: true })
    );
  }
  if (mode === 'LOOSE' && palletSize > 0) {
    return list.sort((a, b) =>
      (a.isFullPallet ? 1 : 0) - (b.isFullPallet ? 1 : 0) ||
      a.available - b.available ||
      String(a.binNo).localeCompare(String(b.binNo), undefined, { numeric: true })
    );
  }
  return list.sort((a, b) => a.available - b.available || String(a.binNo).localeCompare(String(b.binNo), undefined, { numeric: true }));
}

function getPendingLoadingAllocMap_() {
  const sh = SS.getSheetByName('LOADING_ALLOC');
  if (!sh) return {};
  const data = sh.getDataRange().getDisplayValues();
  const map  = {};
  for (let i = 1; i < data.length; i++) {
    const st = String(data[i][10] || '').trim().toUpperCase();
    if (st !== 'PENDING') continue;
    const key = String(data[i][5] || '').trim() + '|' + String(data[i][7] || '').trim() + '|' + String(data[i][8] || '').trim();
    map[key]  = (map[key] || 0) + Number(data[i][9] || 0);
  }
  return map;
}

function getPendingPickMap_() {
  const sh = SS.getSheetByName('LOADING_ALLOC');
  if (!sh) return {};
  const data = sh.getDataRange().getDisplayValues();
  const map  = {};
  for (let i = 1; i < data.length; i++) {
    const status      = String(data[i][10] || '').trim().toUpperCase();
    if (status !== 'PENDING') continue;
    const materialCode= String(data[i][5]  || '').trim();
    const batch       = String(data[i][7]  || '').trim();
    const binNo       = String(data[i][8]  || '').trim();
    const qty         = Number(data[i][9]  || 0);
    const deliveryNo  = String(data[i][2]  || '').trim();
    const loadingNo   = String(data[i][0]  || '').trim();
    if (!materialCode || !batch || !binNo || qty <= 0) continue;
    const key = materialCode + '|' + batch + '|' + binNo;
    if (!map[key]) map[key] = { qty: 0, materialCode, batch, binNo, deliveryNos: {}, loadingNos: {} };
    map[key].qty += qty;
    if (deliveryNo) map[key].deliveryNos[deliveryNo] = true;
    if (loadingNo)  map[key].loadingNos[loadingNo]   = true;
  }
  return map;
}

function reduceLiveStock_(materialCode, batch, binNo, qty) {
  const sh = SS.getSheetByName('LIVE_STOCK');
  if (!sh) throw new Error('LIVE_STOCK sheet missing');
  const data = sh.getDataRange().getDisplayValues();
  let remaining = Number(qty || 0);

  for (let i = 1; i < data.length && remaining > 0; i++) {
    const mc   = String(data[i][2]  || '').trim();
    const bt   = String(data[i][4]  || '').trim();
    const bin  = String(data[i][5]  || '').trim();
    const type = String(data[i][14] || 'GOOD').trim().toUpperCase();
    const bal  = Number(data[i][6]  || 0);
    if (mc === String(materialCode).trim() && bt === String(batch).trim() && bin === String(binNo).trim() && type === 'GOOD' && bal > 0) {
      const cut = Math.min(remaining, bal);
      sh.getRange(i + 1, 7).setValue(bal - cut);
      remaining -= cut;
    }
  }
  if (remaining > 0) throw new Error('Insufficient stock in bin ' + binNo + ' for ' + materialCode + ' / ' + batch + '. Short by: ' + remaining);
}

/* ════════════════════════════════════════════════════════
   LOADING SHEET PRINT DATA
   ════════════════════════════════════════════════════════ */
function getLoadingSheetPrintData(loadingNoList) {
  if (!loadingNoList) loadingNoList = [];
  if (!Array.isArray(loadingNoList)) loadingNoList = [loadingNoList];
  loadingNoList = loadingNoList.map(String).map(s => s.trim()).filter(Boolean);
  if (!loadingNoList.length) return [];

  const allRows  = loadingNoList.flatMap(no => getLoadingAllocRows(no));
  const caseLotMap = getProductCaseLotMap_();
  const pageMap  = {};

  allRows.forEach(r => {
    const loadingNo  = String(r.loadingNo  || '').trim();
    const deliveryNo = String(r.deliveryNo || '').trim();
    const pageKey    = loadingNo + '|' + deliveryNo;

    if (!pageMap[pageKey]) pageMap[pageKey] = { loadingNo, deliveryNo, fromCode: String(r.cfaCode || ''), fromName: String(r.cfaName || ''), totalQty: 0, linesMap: {} };

    const mc       = String(r.materialCode || '').trim();
    const mn       = String(r.materialName || '').trim();
    const bt       = String(r.batch        || '').trim();
    const qty      = Number(r.pickQty || 0);
    const lineKey  = mc + '|' + mn + '|' + bt;

    if (!pageMap[pageKey].linesMap[lineKey]) pageMap[pageKey].linesMap[lineKey] = { materialCode: mc, materialName: mn, batch: bt, qty: 0, caseLot: Number(caseLotMap[mc] || 1) };
    pageMap[pageKey].linesMap[lineKey].qty += qty;
    pageMap[pageKey].totalQty += qty;
  });

  return Object.values(pageMap)
    .sort((a, b) => String(a.deliveryNo).localeCompare(String(b.deliveryNo), undefined, { numeric: true }))
    .map(p => {
      const lines = Object.values(p.linesMap).sort((a, b) => String(a.materialCode).localeCompare(String(b.materialCode), undefined, { numeric: true }));
      lines.forEach((l, idx) => {
        const cl = Number(l.caseLot || 1);
        l.srNo    = idx + 1;
        l.fullBox = cl > 0 ? Math.floor(Number(l.qty) / cl) : 0;
        l.looseQty= cl > 0 ? Number(l.qty) % cl : Number(l.qty);
      });
      delete p.linesMap;
      p.lines = lines;
      return p;
    });
}

function getProductCaseLotMap_() {
  const sh  = SS.getSheetByName('PRODUCT_MASTER');
  if (!sh) return {};
  const map = {};
  const data= sh.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    const code    = String(data[i][0] || '').trim();
    const caseLot = Number(data[i][4] || data[i][2] || 1);
    if (code) map[code] = caseLot > 0 ? caseLot : 1;
  }
  return map;
}

/* ════════════════════════════════════════════════════════
   SAP / OPENING STOCK DIFF
   ════════════════════════════════════════════════════════ */
function getOpeningStockDiffReport(sapRows) {
  sapRows = sapRows || [];
  const sapMap = {};

  function clean_(v)   { let s = String(v == null ? '' : v).trim(); if (s.endsWith('.0')) s = s.slice(0,-2); return s.toUpperCase(); }
  function matKey_(v)  { let s = clean_(v).replace(/\s+/g,''); if (/^\d+[A-Z]$/.test(s)) s = s.slice(0,-1); return s; }
  function batKey_(v)  { return clean_(v).replace(/\s+/g,''); }
  function qty_(v)     { const n = Number(String(v==null?'':v).replace(/,/g,'').trim()); return isNaN(n) ? 0 : n; }

  sapRows.forEach(r => {
    const mk = matKey_(r.materialCode || '');
    const bk = batKey_(r.batch || '');
    if (!mk || !bk) return;
    const key = mk + '|' + bk;
    if (!sapMap[key]) sapMap[key] = { materialCode: String(r.materialCode || ''), materialName: String(r.materialName || ''), batch: String(r.batch || ''), sapGood: 0, sapBlocked: 0 };
    sapMap[key].sapGood    += qty_(r.unrestricted || r.goodQty    || 0);
    sapMap[key].sapBlocked += qty_(r.blocked      || r.blockedQty || 0);
  });

  const liveSh = SS.getSheetByName('LIVE_STOCK');
  const wmsMap = {};
  if (liveSh) {
    const d = liveSh.getDataRange().getDisplayValues();
    for (let i = 1; i < d.length; i++) {
      const mc   = String(d[i][2] || '').trim();
      const mn   = String(d[i][3] || '').trim();
      const bt   = String(d[i][4] || '').trim();
      const qty  = Number(d[i][6] || 0);
      const type = clean_(d[i][14] || 'GOOD');
      const mk   = matKey_(mc), bk = batKey_(bt);
      if (!mk || !bk || qty <= 0) continue;
      const key = mk + '|' + bk;
      if (!wmsMap[key]) wmsMap[key] = { materialCode: mc, materialName: mn, batch: bt, wmsGood: 0, wmsDamage: 0, wmsCqa: 0, wmsOther: 0 };
      if (['GOOD','OPENING','UNRESTRICTED'].includes(type))         wmsMap[key].wmsGood   += qty;
      else if (['DAMAGE','DAMAGED','BLOCKED','BLOCK'].includes(type)) wmsMap[key].wmsDamage += qty;
      else if (['CQA','QUALITY','QI'].includes(type))                wmsMap[key].wmsCqa    += qty;
      else                                                            wmsMap[key].wmsOther  += qty;
    }
  }

  const keys = [...new Set([...Object.keys(sapMap), ...Object.keys(wmsMap)])];
  return keys.map(key => {
    const s = sapMap[key] || {};
    const w = wmsMap[key] || {};
    const sapGood    = Number(s.sapGood    || 0), sapBlocked = Number(s.sapBlocked || 0);
    const wmsGood    = Number(w.wmsGood    || 0), wmsDamage  = Number(w.wmsDamage  || 0);
    const wmsCqa     = Number(w.wmsCqa     || 0), wmsOther   = Number(w.wmsOther   || 0);
    const goodDiff   = wmsGood - sapGood, damageDiff = wmsDamage - sapBlocked;
    const totalDiff  = (wmsGood + wmsDamage + wmsCqa + wmsOther) - (sapGood + sapBlocked);
    return {
      materialCode: s.materialCode || w.materialCode || '',
      materialName: s.materialName || w.materialName || '',
      batch:        s.batch        || w.batch        || '',
      sapGood, sapBlocked, wmsGood, wmsDamage, wmsCqa, wmsOther,
      goodDiff, damageDiff, totalDiff,
      status: totalDiff === 0 && goodDiff === 0 && damageDiff === 0 ? 'OK' : 'DIFF'
    };
  }).sort((a, b) =>
    String(a.materialCode).localeCompare(String(b.materialCode), undefined, { numeric: true }) ||
    String(a.batch).localeCompare(String(b.batch), undefined, { numeric: true })
  );
}


/* ════════════════════════════════════════════════════════
   BULK UPLOAD — SKU MASTER & BIN MASTER (called from HTML)
   ════════════════════════════════════════════════════════ */

/**
 * Replaces all data in SKU_Master sheet with the uploaded rows.
 * Called by the HTML app via google.script.run.bulkSaveSKUMaster(rows)
 * rows = [{SKU Code, SKU Name, Case Pack (Units/Box), Box/Pallet (Boxes/Shelf), Classification (A/B/C)}]
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
function clearSheetData(sheetName) {
  try {
    if (!sheetName) return { success: false, message: 'sheetName required' };
    const sh = SS.getSheetByName(sheetName);
    if (!sh) return { success: false, message: 'Sheet not found: ' + sheetName };

    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return { success: true, message: 'Sheet already empty', count: 0 };

    // Read header first (batch read — avoids multiple API calls)
    const headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const headerBg  = sh.getRange(1, 1, 1, sh.getLastColumn()).getBackgrounds()[0];
    const headerFc  = sh.getRange(1, 1, 1, sh.getLastColumn()).getFontColors()[0];
    const headerFw  = sh.getRange(1, 1, 1, sh.getLastColumn()).getFontWeights()[0];

    const deletedRows = lastRow - 1;

    // Delete all rows below header in one operation
    sh.deleteRows(2, deletedRows);

    // Re-apply header style (deleteRows can sometimes strip formatting)
    const hr = sh.getRange(1, 1, 1, headerRow.length);
    hr.setValues([headerRow])
      .setBackgrounds([headerBg])
      .setFontColors([headerFc])
      .setFontWeights([headerFw]);
    sh.setFrozenRows(1);

    return { success: true, message: sheetName + ' cleared (' + deletedRows + ' rows deleted)', count: deletedRows };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/* ════════════════════════════════════════════════════════
   ⑪ NEW v7: refreshAllData
   Single call to reload ALL reference data from sheets.
   Returns everything loadAllWMSData returns PLUS
   inventory from INVENTORY_DUMP and gatepasses.
   Frontend uses this for the global Refresh button.
   ════════════════════════════════════════════════════════ */

/**
 * Cache helper using CacheService (up to 6h).
 * Skips sheet read if data was fetched recently.
 * 100KB limit per key — silently falls back to live read if too large.
 */
/* ── Chunked CacheService helpers (no size limit) ── */
const _CACHE_CHUNK_ = 95000; // bytes per chunk — under GAS 100KB per-key limit

function _cachedRead_(key, readFn, hours) {
  try {
    const hit = _cacheGetChunked_('wms_' + key);
    if (hit !== null) return hit;
  } catch(e) {}
  const data = readFn();
  try { _cacheSetChunked_('wms_' + key, data, (hours || 6) * 3600); } catch(e) {}
  return data;
}

function _invalidateCache_(key) {
  _cacheDeleteChunked_('wms_' + key);
}

function _cacheGetChunked_(cKey) {
  const cache = CacheService.getScriptCache();
  const meta  = cache.get(cKey + '__n');
  if (!meta) return null;
  const n = parseInt(meta);
  const keys = Array.from({length: n}, (_, i) => cKey + '__' + i);
  const chunks = cache.getAll(keys);
  const parts  = [];
  for (let i = 0; i < n; i++) {
    if (!chunks[cKey + '__' + i]) return null; // chunk expired
    parts.push(chunks[cKey + '__' + i]);
  }
  return JSON.parse(parts.join(''));
}

function _cacheSetChunked_(cKey, data, ttlSec) {
  const cache  = CacheService.getScriptCache();
  const json   = JSON.stringify(data);
  const n      = Math.ceil(json.length / _CACHE_CHUNK_);
  const entries = {[cKey + '__n']: String(n)};
  for (let i = 0; i < n; i++) {
    entries[cKey + '__' + i] = json.slice(i * _CACHE_CHUNK_, (i + 1) * _CACHE_CHUNK_);
  }
  cache.putAll(entries, ttlSec);
}

function _cacheDeleteChunked_(cKey) {
  try {
    const cache = CacheService.getScriptCache();
    const meta  = cache.get(cKey + '__n');
    const n     = meta ? parseInt(meta) : 8;
    const keys  = [cKey + '__n'];
    for (let i = 0; i < n; i++) keys.push(cKey + '__' + i);
    cache.removeAll(keys);
  } catch(e) {}
}

function refreshAllData() {
  try {
    const tz = Session.getScriptTimeZone();

    // ── SKU Master (cache 6h — changes only on bulk upload or inline add/edit) ──
    const skuMaster = _cachedRead_('skuMaster', getSKUMasterFromSheet, 6);

    // ── Bin Master (cache 6h — invalidated on every write) ──
    let binMaster = _cachedRead_('binMaster', getBinMasterFromSheet, 6);

    // ── Gatepasses (cache 2h) ──
    const gatepasses = _cachedRead_('gatepasses', getGatepassesFromSheet, 2);

    // ── Inventory Dump (all rows for cycle count + current inventory) ──
    const invSh = SS.getSheetByName('INVENTORY_DUMP');
    let inventory = [];
    if (invSh && invSh.getLastRow() >= 2) {
      const invLastRow = invSh.getLastRow();
      const invData = invLastRow > 1
        ? invSh.getRange(2, 1, invLastRow - 1, 14).getDisplayValues() // 14 cols, no getDataRange
        : [];
      inventory = invData.filter(r => r[0] && Number(r[9]) > 0).map(r => ({
        skuCode:   String(r[0]  || '').trim(),
        skuName:   String(r[1]  || '').trim(),
        batch:     String(r[2]  || '').trim(),
        mfgDate:   String(r[3]  || '').trim(),   // PASS AS-IS from INVENTORY_DUMP
        expDate:   String(r[4]  || '').trim(),   // PASS AS-IS from INVENTORY_DUMP
        mrp:       String(r[5]  || '').trim(),
        bin:       normalizeBinIdForMaster_(r[6]),
        stockType: String(r[7]  || 'GOOD').trim().toUpperCase(),
        skuStatus: String(r[8]  || 'ACTIVE').trim().toUpperCase(),
        qty:       Number(r[9]  || 0),
        ean:       String(r[10] || '').trim()
      }));
    }

    // ── Auto-update bin statuses from freshly loaded inventory ──
    // This ensures bins reflect current occupation without manual trigger
    if (inventory.length > 0) {
      try {
        const dumpRows = inventory.map(r => [
          r.skuCode, r.skuName, r.batch, r.mfgDate, r.expDate,
          r.mrp, r.bin, r.stockType, r.skuStatus, r.qty, r.ean, '', ''
        ]);
        uwUpdateBinStatus_(dumpRows);
        // Update bin status in-memory — no second sheet read needed
        const _occ=new Set(inventory.filter(r=>r.qty>0).map(r=>r.bin));
        const _sm={}; inventory.forEach(r=>{if(r.bin&&!_sm[r.bin])_sm[r.bin]=r.skuCode;});
        binMaster=binMaster.map(b=>{
          if(b.status==='Blocked') return b;
          return{...b,status:_occ.has(b.id)?'Occupied':(b.status==='Reserved'?'Reserved':'Empty'),skuFromDump:_sm[b.id]||''};
        });
      } catch(e) {
        Logger.log('refreshAllData bin update error (non-fatal): ' + e.message);
      }
    }

    return {
      success:    true,
      skuMaster,
      binMaster,
      gatepasses,
      inventory,
      refreshedAt: Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy HH:mm:ss')
    };
  } catch(e) {
    return { success: false, message: e.message, skuMaster: [], binMaster: [], gatepasses: [], inventory: [] };
  }
}


/* ════════════════════════════════════════════════════════
   BRAND-BASED BIN ALLOCATION ENGINE
   ════════════════════════════════════════════════════════ */

/**
 * Brand zone configuration.
 * Defines which Row + Column ranges belong to which brand.
 * Update this map if new rack zones are added — no other code changes needed.
 *
 * Structure: { brandName: [ {rows:[R1..R6], cols:[C1..C10]}, ... ] }
 */
const BRAND_ZONES = {
  'Bodywise': [
    { rowMin: 1,  rowMax: 6,  colMin: 1,  colMax: 10 },  // R1-R6,  C1-C10
    { rowMin: 7,  rowMax: 13, colMin: 1,  colMax: 20 },  // R7-R13, all cols
  ],
  'OWN': [
    { rowMin: 1,  rowMax: 6,  colMin: 11, colMax: 20 },  // R1-R6,  C11-C20
  ],
  'LittleJoys': [
    { rowMin: 14, rowMax: 23, colMin: 1,  colMax: 20 },  // R14-R23, all cols
  ],
  'ManMatters': [
    { rowMin: 24, rowMax: 26, colMin: 1,  colMax: 20 },  // R24-R26, all cols
  ],
  // R27-R35 = Overflow — not pre-assigned, allocated on user request
};

/** SKU prefix → Brand name */
const SKU_BRAND_MAP = {
  'MWBW': 'Bodywise',
  'OWN':  'OWN',
  'MWLJ': 'LittleJoys',
  'MWMM': 'ManMatters',
};

/**
 * Identify brand from SKU code prefix.
 * Returns brand name or null if unrecognised.
 */
function getBrandFromSKU(skuCode) {
  skuCode = String(skuCode || '').trim().toUpperCase();
  for (const prefix of Object.keys(SKU_BRAND_MAP)) {
    if (skuCode.startsWith(prefix.toUpperCase())) return SKU_BRAND_MAP[prefix];
  }
  return null; // unknown brand
}

/**
 * Parse a Bin ID like "R5-C12-003" into { row:5, col:12, level:3 }
 */
function parseBinId_(binId) {
  const m = String(binId).match(/^R(\d+)-C(\d+)-(\d+)$/i);
  if (!m) return null;
  return { row: parseInt(m[1]), col: parseInt(m[2]), level: parseInt(m[3]) };
}

/**
 * Check if a bin (parsed) falls within a brand zone definition.
 */
function binInZone_(parsed, zone) {
  return parsed.row >= zone.rowMin && parsed.row <= zone.rowMax &&
         parsed.col >= zone.colMin && parsed.col <= zone.colMax;
}

/**
 * Check if a bin belongs to a specific brand zone.
 */
function binBelongsToBrand_(binId, brand) {
  const parsed = parseBinId_(binId);
  if (!parsed) return false;
  const zones = BRAND_ZONES[brand];
  if (!zones) return false;
  return zones.some(z => binInZone_(parsed, z));
}

/**
 * Check if a bin is in the overflow zone (R27-R35).
 */
function binIsOverflow_(binId) {
  const parsed = parseBinId_(binId);
  if (!parsed) return false;
  return parsed.row >= 27 && parsed.row <= 35;
}

/**
 * Core allocation function — finds the best empty bin for a brand + FSN class.
 *
 * Priority:
 *   1. Brand's own zone, requested FSN class
 *   2. Brand's own zone, FSN B (if A was requested and full)
 *   3. Brand's own zone, FSN C (if B also full)
 *   4. Overflow zone (R27-R35), specific row chosen by user
 *   5. null → no bins available
 *
 * @param {string}   brand       — Brand name (from getBrandFromSKU)
 * @param {string}   fsnClass    — 'A', 'B', or 'C'
 * @param {Set}      usedBins    — bins already allocated in this GRN run
 * @param {number[]} overflowRows — overflow row numbers user selected (e.g. [27,28])
 * @returns {{ binId, fsn, zone, isOverflow, suggestion }} or null
 */
function findBestBin_(brand, fsnClass, usedBins, overflowRows, binMasterData) {
  fsnClass = (fsnClass || 'B').toUpperCase();
  const fsns = fsnClass === 'A' ? ['A','B','C'] :
               fsnClass === 'B' ? ['B','C','A'] : ['C','B','A'];

  // Try brand zone first
  for (const fsn of fsns) {
    const bin = findEmptyBinInZone_(brand, fsn, usedBins, binMasterData, false, []);
    if (bin) return { ...bin, isOverflow: false };
  }

  // Brand zone full — try overflow rows if user provided them
  if (overflowRows && overflowRows.length) {
    for (const fsn of fsns) {
      const bin = findEmptyBinInZone_(brand, fsn, usedBins, binMasterData, true, overflowRows);
      if (bin) return { ...bin, isOverflow: true };
    }
  }

  return null; // no bins available
}

/**
 * Scan BIN_MASTER_WMS data for an empty bin matching criteria.
 * binMasterData: array of rows from sheet (already read in one batch call).
 *
 * BIN_MASTER_WMS columns (0-based):
 *   0:BinID  1:Row  2:ColNo  3:Level  4:BrandZone  5:FSN  6:Capacity  7:Status  8:SKUinBin  9:UpdatedAt  10:UpdatedBy
 */
function findEmptyBinInZone_(brand, fsn, usedBins, binMasterData, overflow, overflowRows) {
  for (let i = 0; i < binMasterData.length; i++) {
    const row     = binMasterData[i];
    const binId   = String(row[0] || '').trim();
    const status  = String(row[6] || '').trim().toUpperCase(); // col G = Status
    const binFsn  = String(row[5] || '').trim().toUpperCase(); // col F = FSN
    const zoneBrand = String(row[4] || '').trim();             // col E = Brand Zone

    if (!binId || usedBins.has(binId)) continue;
    if (status === 'BLOCKED' || status === 'Blocked') continue;
    if (status !== 'EMPTY') continue;
    if (binFsn !== fsn) continue;

    const parsed = parseBinId_(binId);
    if (!parsed) continue;

    if (overflow) {
      // Must be in overflow rows
      if (!overflowRows.includes(parsed.row)) continue;
    } else {
      // Must belong to brand's zone
      if (zoneBrand !== brand) continue;
    }

    return {
      binId,
      fsn: binFsn,
      zone: overflow ? 'Overflow R' + parsed.row : brand,
      suggestion: overflow
        ? 'Overflow zone (R' + parsed.row + ') — brand zone full'
        : brand + ' zone — FSN ' + binFsn
    };
  }
  return null;
}

/**
 * Get Overflow status — how many empty bins exist per row in R27-R35.
 * Returns: { 27: 45, 28: 12, ... }
 */
function getOverflowAvailability() {
  try {
    const sh     = SS.getSheetByName('BIN_MASTER_WMS'); // SS not openById
    if (!sh || sh.getLastRow() < 2) return {};
    const data   = sh.getDataRange().getDisplayValues().slice(1);
    const counts = {};
    data.forEach(row => {
      const binId  = String(row[0] || '').trim();
      const status = String(row[7] || '').trim().toUpperCase();
      const parsed = parseBinId_(binId);
      if (!parsed || !binIsOverflow_(binId) || status !== 'EMPTY') return;
      counts[parsed.row] = (counts[parsed.row] || 0) + 1;
    });
    return counts;
  } catch(e) { return {}; }
}

/**
 * Called from frontend to get overflow row options.
 */
function getOverflowOptions() {
  try {
    const avail = getOverflowAvailability();
    const rows  = [];
    for (let r = 27; r <= 35; r++) {
      rows.push({ row: r, emptyBins: avail[r] || 0 });
    }
    return { success: true, rows };
  } catch(e) { return { success: false, rows: [], message: e.message }; }
}


function refreshDashboard() {
  const counts = getDashboardCounts();
  const msg = 'Bins: ' + counts.total + ' total | ' + counts.empty + ' empty | ' + counts.occupied + ' occupied\nTOs: ' + counts.pendingTO + ' pending | ' + counts.confirmedTO + ' confirmed\nStock (Good): ' + (counts.totalGoodStock || 0) + ' units';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Dashboard Summary\n\n' + msg); } catch(e) {}
}

function openStockReport() {
  const report = getAvailableStockReport();
  Logger.log('Stock Report: ' + report.length + ' lines.');
  try { SpreadsheetApp.getUi().alert('Stock Report: ' + report.length + ' lines. Open the web app for full view.'); } catch(e) {}
}


/* ════════════════════════════════════════════════════════
   ══════════════════════════════════════════════════════════════════
   UNIWARE AUTO-SYNC MODULE
   Automatically imports Shelfwise Inventory CSV from Gmail into
   INVENTORY_DUMP every 30 minutes.

   SETUP (run once from Apps Script editor):
     1. Run setupUniwareSync()  — creates trigger + Gmail label
     2. Run testUniwareSyncNow() — verify with a real email
     3. Done. Runs automatically every 30 min forever.
   ══════════════════════════════════════════════════════════════════
   ════════════════════════════════════════════════════════ */

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
