/**
 * Modular split from production-ready/Code.gs
 * File: 01-setup-auth.gs
 */

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
    }
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
    ['admin',    'admin123', 'Admin User',   'ADMIN',    'YES', '', 'ALL'],
    ['wms1',     'wms123',   'WMS User 1',   'OPERATOR', 'YES', '', 'INBOUND'],
    ['wms2',     'wms123',   'WMS User 2',   'OPERATOR', 'YES', '', 'OPERATIONS,PLANNING'],
    ['manager',  'mgr123',   'WMS Manager',  'MANAGER',  'YES', '', 'INBOUND,OPERATIONS,PLANNING'],
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
  const header = data[0].map(h => String(h || '').trim().toUpperCase());
  const sectionsIdx = header.indexOf('ALLOWED SECTIONS');
  for (let i = 1; i < data.length; i++) {
    const rowId   = String(data[i][0] || '').trim();
    const rowPw   = String(data[i][1] || '').trim();
    const rowName = String(data[i][2] || '').trim();
    const rowRole = String(data[i][3] || '').trim().toUpperCase();
    const active  = String(data[i][4] || '').trim().toUpperCase();
    const allowedSections = sectionsIdx >= 0
      ? String(data[i][sectionsIdx] || '').trim().toUpperCase()
      : (rowRole === 'ADMIN' ? 'ALL' : '');

    if (rowId === userId && rowPw === password) {
      if (active !== 'YES') return { success: false, message: 'Account inactive — contact Admin.' };

      // Write last login timestamp
      try {
        const tz = Session.getScriptTimeZone();
        sh.getRange(i + 1, 6).setValue(Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy HH:mm:ss'));
      } catch(e) {}

      return {
        success:  true,
        name:     rowName,
        role:     rowRole,   // ADMIN or anything else
        userId:   rowId,
        allowedSections,
      };
    }
  }
  return { success: false, message: 'Wrong User ID or Password' };
}

/* ════════════════════════════════════════════════════════
   DASHBOARD
   ════════════════════════════════════════════════════════ */
