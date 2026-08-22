/**
 * Modular split from Code.gs
 * File: 00-router.gs
 */

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
