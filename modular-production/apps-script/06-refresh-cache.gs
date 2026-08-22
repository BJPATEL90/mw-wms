/**
 * Modular split from production-ready/Code.gs
 * File: 06-refresh-cache.gs
 */

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
function _cachedRead_(key, readFn, hours) {
  const cache = CacheService.getScriptCache();
  const cKey  = 'wms_' + key;
  try {
    const hit = cache.get(cKey);
    if (hit) return JSON.parse(hit);
  } catch(e) {}
  const data = readFn();
  try {
    const str = JSON.stringify(data);
    if (str.length < 90000) cache.put(cKey, str, (hours || 6) * 3600);
  } catch(e) {}
  return data;
}

function _invalidateCache_(key) {
  try { CacheService.getScriptCache().remove('wms_' + key); } catch(e) {}
}

function refreshAllData() {
  try {
    const tz = Session.getScriptTimeZone();

    // ── SKU Master (CacheService 6h) ──
    const skuMaster = _cachedRead_('skuMaster', getSKUMasterFromSheet, 6);

    // ── Bin Master: read directly so stale/polluted cache cannot inflate counts ──
    let binMaster = getBinMasterFromSheet();

    // ── Gatepasses ──
    const gatepasses = getGatepassesFromSheet();

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
