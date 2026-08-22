/**
 * Modular split from Code.gs
 * File: 07-bin-allocation-dashboard.gs
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
