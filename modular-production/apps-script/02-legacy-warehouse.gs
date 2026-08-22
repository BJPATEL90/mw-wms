/**
 * Modular split from production-ready/Code.gs
 * File: 02-legacy-warehouse.gs
 */

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
